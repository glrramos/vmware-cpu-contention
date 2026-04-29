<#
.SYNOPSIS
    Diagnostico de match de VMs entre vCenter e vROps.

.DESCRIPTION
    Verifica por que Get-VMwareCPUContention.ps1 esta retornando dados
    vROps zerados. Para cada VM amostrada, testa 4 estrategias de busca
    no vROps e reporta qual funcionou:

      1. Match exato pelo nome (estrategia atual do script)
      2. Match por nome curto (stripping FQDN)
      3. Match por regex (substring)
      4. Match por instanceUuid via property condition

    Quando -ListStatKeys e fornecido, para a primeira VM encontrada
    lista todas as stat keys disponiveis e valida se as keys que o
    Get-VMwareCPUContention.ps1 pede realmente existem no vROps 8.18.

    Operacao read-only.

.PARAMETER VCenter
    FQDN ou IP do vCenter Server.

.PARAMETER VRopsServer
    FQDN ou IP do vROps.

.PARAMETER VRopsCredential
    PSCredential para vROps. Se omitido, prompt interativo.

.PARAMETER VRopsAuthSource
    Auth source. Default 'LOCAL'.

.PARAMETER VMNames
    Lista de nomes de VMs para testar. Se omitido, pega as 5 primeiras
    VMs ligadas do inventario.

.PARAMETER ListStatKeys
    Lista as stat keys disponiveis para a primeira VM encontrada e
    valida quais das keys usadas pelo script principal estao expostas.

.EXAMPLE
    .\Test-VRopsMatching.ps1 -VCenter vc01 -VRopsServer vrops01 `
        -VMNames "vm-app01","vm-db01","vm-edge01" -ListStatKeys

.EXAMPLE
    .\Test-VRopsMatching.ps1 -VCenter vc01 -VRopsServer vrops01

    Pega 5 VMs aleatorias do inventario e roda os 4 testes de match.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$VCenter,
    [Parameter(Mandatory)][string]$VRopsServer,
    [pscredential]$VRopsCredential,
    [string]$VRopsAuthSource = 'LOCAL',
    [string[]]$VMNames,
    [switch]$ListStatKeys
)

$ErrorActionPreference = 'Stop'

#region Helpers ------------------------------------------------------------------

function Set-CertificatePolicy {
    if ($PSVersionTable.PSVersion.Major -lt 6) {
        if (-not ('TrustAllCertsPolicy' -as [type])) {
            Add-Type -TypeDefinition @"
using System.Net;
using System.Security.Cryptography.X509Certificates;
public class TrustAllCertsPolicy : ICertificatePolicy {
    public bool CheckValidationResult(ServicePoint sp, X509Certificate cert, WebRequest req, int problem) {
        return true;
    }
}
"@
        }
        [System.Net.ServicePointManager]::CertificatePolicy = New-Object TrustAllCertsPolicy
        [System.Net.ServicePointManager]::SecurityProtocol  = [System.Net.SecurityProtocolType]::Tls12
    }
}

function Invoke-VRopsApi {
    param(
        [Parameter(Mandatory)][ValidateSet('GET','POST')][string]$Method,
        [Parameter(Mandatory)][string]$Uri,
        [object]$Body,
        [hashtable]$Headers
    )
    if (-not $Headers) { $Headers = @{} }
    if (-not $Headers.ContainsKey('Accept')) { $Headers['Accept'] = 'application/json' }
    $params = @{
        Method      = $Method
        Uri         = $Uri
        Headers     = $Headers
        ContentType = 'application/json'
        ErrorAction = 'Stop'
    }
    if ($PSBoundParameters.ContainsKey('Body') -and $null -ne $Body) {
        $params.Body = ($Body | ConvertTo-Json -Depth 10 -Compress)
    }
    if ($PSVersionTable.PSVersion.Major -ge 6) {
        $params.SkipCertificateCheck = $true
    }
    Invoke-RestMethod @params
}

function Connect-VRopsLocal {
    param([string]$Server, [pscredential]$Cred, [string]$AuthSource)
    $body = @{
        username   = $Cred.UserName
        password   = $Cred.GetNetworkCredential().Password
        authSource = $AuthSource
    }
    $uri  = "https://$Server/suite-api/api/auth/token/acquire"
    $resp = Invoke-VRopsApi -Method POST -Uri $uri -Body $body
    return [PSCustomObject]@{
        Server  = $Server
        Headers = @{ 'Authorization' = "vRealizeOpsToken $($resp.token)" }
    }
}

function Disconnect-VRopsLocal {
    param($Session)
    if (-not $Session) { return }
    try {
        $uri = "https://$($Session.Server)/suite-api/api/auth/token/release"
        Invoke-VRopsApi -Method POST -Uri $uri -Headers $Session.Headers | Out-Null
    } catch { }
}

function Get-MatchByName {
    param($Session, [string]$Name)
    $encoded = [uri]::EscapeDataString($Name)
    $uri = "https://$($Session.Server)/suite-api/api/resources?name=$encoded&resourceKind=VirtualMachine"
    try {
        $resp = Invoke-VRopsApi -Method GET -Uri $uri -Headers $Session.Headers
        return @($resp.resourceList)
    } catch {
        return $null
    }
}

function Get-MatchByRegex {
    param($Session, [string]$Pattern)
    $encoded = [uri]::EscapeDataString($Pattern)
    $uri = "https://$($Session.Server)/suite-api/api/resources?regex=$encoded&resourceKind=VirtualMachine&pageSize=10"
    try {
        $resp = Invoke-VRopsApi -Method GET -Uri $uri -Headers $Session.Headers
        return @($resp.resourceList)
    } catch {
        return $null
    }
}

function Get-MatchByUuid {
    param($Session, [string]$Uuid)
    # Tenta dois paths comuns para instanceUuid - varia por versao do MP
    $candidates = @('config|instanceUuid', 'summary|config|instanceUuid')
    foreach ($key in $candidates) {
        $body = @{
            resourceKind = @('VirtualMachine')
            propertyConditions = @{
                conditions = @(
                    @{
                        key         = $key
                        operator    = 'EQ'
                        stringValue = $Uuid
                    }
                )
            }
        }
        $uri = "https://$($Session.Server)/suite-api/api/resources/query"
        try {
            $resp  = Invoke-VRopsApi -Method POST -Uri $uri -Headers $Session.Headers -Body $body
            $items = @($resp.resourceList)
            if ($items.Count -gt 0) { return $items }
        } catch { }
    }
    return @()
}

function Test-VRopsMatch {
    param(
        $Session,
        [string]$VMName,
        [string]$ShortName,
        [string]$InstanceUuid
    )
    $r = [PSCustomObject]@{
        VMName       = $VMName
        ShortName    = $ShortName
        InstanceUuid = $InstanceUuid
        ByExact      = '-'
        ByShort      = '-'
        ByRegex      = '-'
        ByUuid       = '-'
        Recommended  = 'NENHUMA'
        ResourceId   = $null
    }

    # 1. Match exato (estrategia do script principal hoje)
    $items = Get-MatchByName -Session $Session -Name $VMName
    if ($null -eq $items) {
        $r.ByExact = 'erro'
    } elseif ($items.Count -gt 0) {
        $r.ByExact     = "$($items.Count) hit(s)"
        $r.Recommended = 'exact-name'
        $r.ResourceId  = $items[0].identifier
    } else {
        $r.ByExact = 'vazio'
    }

    # 2. Match por nome curto (so se diferente do exato)
    if ($ShortName -and $ShortName -ne $VMName) {
        $items = Get-MatchByName -Session $Session -Name $ShortName
        if ($null -eq $items) {
            $r.ByShort = 'erro'
        } elseif ($items.Count -gt 0) {
            $r.ByShort = "$($items.Count) hit(s)"
            if ($r.Recommended -eq 'NENHUMA') {
                $r.Recommended = 'short-name'
                $r.ResourceId  = $items[0].identifier
            }
        } else {
            $r.ByShort = 'vazio'
        }
    } else {
        $r.ByShort = 'n/a'
    }

    # 3. Match por regex (substring no nome curto)
    if ($ShortName) {
        $items = Get-MatchByRegex -Session $Session -Pattern $ShortName
        if ($null -eq $items) {
            $r.ByRegex = 'erro'
        } elseif ($items.Count -gt 0) {
            $r.ByRegex = "$($items.Count) hit(s)"
            if ($r.Recommended -eq 'NENHUMA') {
                $r.Recommended = 'regex'
                $r.ResourceId  = $items[0].identifier
            }
        } else {
            $r.ByRegex = 'vazio'
        }
    }

    # 4. Match por instanceUuid (caminho mais robusto se property existe)
    if ($InstanceUuid) {
        $items = Get-MatchByUuid -Session $Session -Uuid $InstanceUuid
        if ($items.Count -gt 0) {
            $r.ByUuid = "$($items.Count) hit(s)"
            if ($r.Recommended -eq 'NENHUMA') {
                $r.Recommended = 'uuid'
                $r.ResourceId  = $items[0].identifier
            }
        } else {
            $r.ByUuid = 'vazio'
        }
    } else {
        $r.ByUuid = 'n/a'
    }

    return $r
}

function Get-VRopsStatKeys {
    param($Session, [string]$ResourceId)
    $uri = "https://$($Session.Server)/suite-api/api/resources/$ResourceId/statkeys"
    try {
        $resp = Invoke-VRopsApi -Method GET -Uri $uri -Headers $Session.Headers
        return @($resp.'stat-key' | Where-Object { $_.key })
    } catch {
        Write-Warning "Falha listando stat keys: $($_.Exception.Message)"
        return @()
    }
}

function Test-VRopsStatHasData {
    param($Session, [string]$ResourceId, [string]$StatKey)
    $endMs   = [DateTimeOffset]::UtcNow.ToUnixTimeMilliseconds()
    $beginMs = [DateTimeOffset]::UtcNow.AddDays(-1).ToUnixTimeMilliseconds()
    $body = @{
        resourceId         = @($ResourceId)
        statKey            = @($StatKey)
        begin              = $beginMs
        end                = $endMs
        rollUpType         = 'AVG'
        intervalType       = 'HOURS'
        intervalQuantifier = 1
    }
    $uri = "https://$($Session.Server)/suite-api/api/resources/stats/query"
    try {
        $resp = Invoke-VRopsApi -Method POST -Uri $uri -Headers $Session.Headers -Body $body
        $values = @($resp.values)
        foreach ($v in $values) {
            $statList = @($v.'stat-list'.stat)
            foreach ($s in $statList) {
                $data = @($s.data)
                if ($data.Count -gt 0) { return [PSCustomObject]@{ HasData = $true; Sample = $data[0]; Count = $data.Count } }
            }
        }
        return [PSCustomObject]@{ HasData = $false; Sample = $null; Count = 0 }
    } catch {
        return [PSCustomObject]@{ HasData = $false; Sample = $null; Count = 0; Error = $_.Exception.Message }
    }
}

#endregion

#region Main ---------------------------------------------------------------------

Set-CertificatePolicy

Write-Host ""
Write-Host "Conectando ao vCenter $VCenter ..." -ForegroundColor Cyan
if (-not (Get-Module -ListAvailable -Name VMware.PowerCLI)) {
    Write-Error "VMware.PowerCLI nao encontrado. Instale com: Install-Module VMware.PowerCLI -Scope CurrentUser"
    exit 1
}
Import-Module VMware.PowerCLI -ErrorAction Stop | Out-Null
Set-PowerCLIConfiguration -InvalidCertificateAction Ignore -ParticipateInCEIP $false -Scope Session -Confirm:$false | Out-Null
$vc = Connect-VIServer -Server $VCenter -ErrorAction Stop
Write-Host "  Conectado: $($vc.Name)  versao $($vc.Version) build $($vc.Build)" -ForegroundColor Green

if (-not $VRopsCredential) {
    $VRopsCredential = Get-Credential -Message "Credenciais vROps em $VRopsServer (auth source: $VRopsAuthSource)"
}
Write-Host "Conectando ao vROps $VRopsServer (auth source: $VRopsAuthSource) ..." -ForegroundColor Cyan
$session = Connect-VRopsLocal -Server $VRopsServer -Cred $VRopsCredential -AuthSource $VRopsAuthSource
Write-Host "  Conectado em vROps" -ForegroundColor Green

try {
    # Selecao de VMs
    if (-not $VMNames -or $VMNames.Count -eq 0) {
        Write-Host ""
        Write-Host "Nenhum -VMNames informado, pegando 5 VMs ligadas do inventario..." -ForegroundColor Yellow
        $vms = @(Get-VM | Where-Object { $_.PowerState -eq 'PoweredOn' } | Select-Object -First 5)
    } else {
        $vms = @()
        foreach ($n in $VMNames) {
            $found = Get-VM -Name $n -ErrorAction SilentlyContinue
            if ($found) {
                if ($found -is [array]) { $vms += $found[0] } else { $vms += $found }
            } else {
                Write-Warning "VM '$n' nao encontrada no vCenter, pulando."
            }
        }
    }
    if ($vms.Count -eq 0) {
        Write-Error "Nenhuma VM para testar."
        exit 1
    }

    Write-Host ""
    Write-Host "Testando match no vROps para $($vms.Count) VM(s)..." -ForegroundColor Cyan
    Write-Host ""

    $results = @()
    foreach ($vm in $vms) {
        $vmName       = $vm.Name
        $instanceUuid = $null
        try { $instanceUuid = $vm.ExtensionData.Config.InstanceUuid } catch { }

        # Variantes de nome curto
        $shortName = $vmName
        if ($vmName.Contains('.')) {
            $shortName = $vmName.Substring(0, $vmName.IndexOf('.'))
        }
        $guestHost = $null
        try { $guestHost = $vm.ExtensionData.Guest.HostName } catch { }
        if ($guestHost -and $guestHost -ne $vmName -and $guestHost -ne $shortName) {
            # Se guest hostname existe e e diferente, prefere ele para o teste short
            $shortName = $guestHost
        }

        Write-Host "  $vmName  (uuid=$instanceUuid, short=$shortName)" -ForegroundColor Gray
        $r = Test-VRopsMatch -Session $session -VMName $vmName -ShortName $shortName -InstanceUuid $instanceUuid
        $results += $r
    }

    Write-Host ""
    Write-Host "===== Resultado por VM =====" -ForegroundColor Cyan
    $results |
        Format-Table VMName, ByExact, ByShort, ByRegex, ByUuid, Recommended -AutoSize |
        Out-Host

    # Resumo
    $semMatch = @($results | Where-Object { $_.Recommended -eq 'NENHUMA' })
    $byExact  = @($results | Where-Object { $_.Recommended -eq 'exact-name' }).Count
    $byShort  = @($results | Where-Object { $_.Recommended -eq 'short-name' }).Count
    $byRegex  = @($results | Where-Object { $_.Recommended -eq 'regex' }).Count
    $byUuid   = @($results | Where-Object { $_.Recommended -eq 'uuid' }).Count
    $tot = $results.Count

    Write-Host ""
    Write-Host "===== Resumo =====" -ForegroundColor Cyan
    $exactColor = if ($byExact -eq $tot) { 'Green' } elseif ($byExact -gt 0) { 'Yellow' } else { 'Red' }
    Write-Host ("Match exato (estrategia atual do script): {0} / {1}" -f $byExact, $tot) -ForegroundColor $exactColor
    Write-Host ("Match por nome curto:                     {0} / {1}" -f $byShort, $tot)
    Write-Host ("Match por regex (substring):              {0} / {1}" -f $byRegex, $tot)
    Write-Host ("Match por instanceUuid:                   {0} / {1}" -f $byUuid,  $tot)
    if ($semMatch.Count -gt 0) {
        Write-Host ""
        Write-Host "VMs sem nenhum match no vROps ($($semMatch.Count)):" -ForegroundColor Red
        $semMatch | ForEach-Object { Write-Host "  - $($_.VMName)" -ForegroundColor Red }
    }

    Write-Host ""
    Write-Host "===== Diagnostico =====" -ForegroundColor Cyan
    if ($byExact -eq $tot) {
        Write-Host "Match exato funciona para todas. O 'zerado' nao e match-por-nome." -ForegroundColor Green
        Write-Host "Hipoteses restantes: stat keys ausentes ou parsing. Rode com -ListStatKeys." -ForegroundColor Yellow
    } elseif (($byShort + $byRegex + $byUuid) -gt 0 -and $byExact -lt $tot) {
        Write-Host "Match exato falha para $($tot - $byExact)/$tot VMs - este e o bug." -ForegroundColor Red
        Write-Host "Fix: Get-VRopsResource precisa cair em fallback para short-name/regex/uuid." -ForegroundColor Yellow
    } elseif ($semMatch.Count -eq $tot) {
        Write-Host "Nenhuma estrategia achou as VMs no vROps." -ForegroundColor Red
        Write-Host "Hipoteses: vROps nao esta inventariando esse vCenter, ou auth source com escopo errado." -ForegroundColor Yellow
    }

    # Stat keys - so para a primeira VM que matchou
    if ($ListStatKeys) {
        $vmOk = $results | Where-Object { $_.ResourceId } | Select-Object -First 1
        if (-not $vmOk) {
            Write-Warning "Nenhuma VM teve match no vROps; nao da pra listar stat keys."
        } else {
            $rid = $vmOk.ResourceId
            Write-Host ""
            Write-Host ("===== Stat keys para '{0}' (resourceId {1}) =====" -f $vmOk.VMName, $rid) -ForegroundColor Cyan
            $keys = Get-VRopsStatKeys -Session $session -ResourceId $rid
            $allKeys = @($keys | ForEach-Object { $_.key })
            Write-Host "Total de keys expostas: $($allKeys.Count)" -ForegroundColor Gray

            $cpuKeys   = @($allKeys | Where-Object { $_ -like 'cpu|*' }     | Sort-Object)
            $memKeys   = @($allKeys | Where-Object { $_ -like 'mem|*' }     | Sort-Object)
            $stKeys    = @($allKeys | Where-Object { $_ -like 'storage|*' -or $_ -like 'diskspace|*' -or $_ -like 'virtualDisk|*' } | Sort-Object)
            $badgeKeys = @($allKeys | Where-Object { $_ -like 'badge|*' }   | Sort-Object)
            $sumKeys   = @($allKeys | Where-Object { $_ -like 'summary|*' } | Sort-Object)
            $sysKeys   = @($allKeys | Where-Object { $_ -like 'sys|*' }     | Sort-Object)

            Write-Host ""
            Write-Host "CPU keys ($($cpuKeys.Count)):" -ForegroundColor Yellow
            $cpuKeys | ForEach-Object { Write-Host "  $_" }
            Write-Host ""
            Write-Host "Mem keys ($($memKeys.Count)):" -ForegroundColor Yellow
            $memKeys | ForEach-Object { Write-Host "  $_" }
            Write-Host ""
            Write-Host "Storage/disk keys ($($stKeys.Count)):" -ForegroundColor Yellow
            $stKeys | ForEach-Object { Write-Host "  $_" }
            Write-Host ""
            Write-Host "Badge keys ($($badgeKeys.Count)):" -ForegroundColor Yellow
            $badgeKeys | ForEach-Object { Write-Host "  $_" }
            Write-Host ""
            Write-Host "Sys keys ($($sysKeys.Count)):" -ForegroundColor Yellow
            $sysKeys | ForEach-Object { Write-Host "  $_" }
            Write-Host ""
            Write-Host "Summary keys ($($sumKeys.Count)):" -ForegroundColor Yellow
            $sumKeys | ForEach-Object { Write-Host "  $_" }

            # Sanity check + amostra de dado para as keys que o script principal usa
            $expected = @(
                'cpu|capacity_contentionPct',
                'cpu|workload',
                'badge|stress',
                'summary|workload_indicator',
                'mem|capacity_contentionPct',
                'diskspace|usage_average',
                'storage|totalLatency_average',
                'sys|workload'
            )
            Write-Host ""
            Write-Host "===== Sanity check: keys que Get-VMwareCPUContention.ps1 usa =====" -ForegroundColor Cyan
            foreach ($e in $expected) {
                $exists = $allKeys -contains $e
                if (-not $exists) {
                    Write-Host ("  {0,-45} AUSENTE - key nao publicada nesta VM" -f $e) -ForegroundColor Red
                    continue
                }
                $sample = Test-VRopsStatHasData -Session $session -ResourceId $rid -StatKey $e
                if ($sample.HasData) {
                    Write-Host ("  {0,-45} OK     ($($sample.Count) pontos, ex: $($sample.Sample))" -f $e) -ForegroundColor Green
                } else {
                    Write-Host ("  {0,-45} VAZIA  - key existe mas sem dados na ultima 24h" -f $e) -ForegroundColor Yellow
                }
            }
        }
    }
} finally {
    Disconnect-VRopsLocal -Session $session
    Disconnect-VIServer -Server $vc -Confirm:$false -ErrorAction SilentlyContinue | Out-Null
    Write-Host ""
    Write-Host "Concluido." -ForegroundColor Green
}

#endregion
