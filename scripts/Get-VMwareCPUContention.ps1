<#
.SYNOPSIS
    Coleta e analisa contencao de CPU em ambientes VMware vSphere com regua diagnostica.

.DESCRIPTION
    Get-VMwareCPUContention.ps1 conecta em um vCenter via PowerCLI, coleta metricas
    de contencao de CPU (Ready, Co-stop, %MLMTD, Latency, SwapWait, Demand) por VM
    e por host na janela informada, classifica cada VM em OK / ATENCAO / CRITICO
    aplicando uma regua diferenciada para NSX Edges, e gera relatorios em JSON,
    CSV, HTML executivo e console colorido.

    O script identifica automaticamente situacoes em que metricas de Usage parecem
    saudaveis mas Ready/Contention estao criticos por causa de monster VMs, gang
    scheduling impossivel, NUMA imbalance ou CPU limits configurados.

    Quando -VRopsServer eh fornecido, enriquece VMs problematicas com historico
    de Aria Operations (vROps) via REST API, calculando trend (subindo/estavel/caindo)
    e cross-validation entre janela atual e media historica para classificar
    como CRONICO / INCIDENTE / MELHORANDO / OSCILANDO.

    Operacao read-only: nao modifica nada no vCenter ou no vROps.

.PARAMETER VCenter
    FQDN ou IP do vCenter Server. Obrigatorio exceto quando -ListVRopsAuthSources.

.PARAMETER VCenterCredential
    PSCredential para vCenter. Se omitido, prompt interativo. Use formato
    UPN (administrator@vsphere.local) ou DOMAIN\user. Evita o sub-prompt
    bugado do PowerCLI quando nao ha SSO passthrough.

.PARAMETER Hours
    Janela de coleta em horas. Default 1. Define automaticamente o rollup do
    vCenter: <=1h real-time (20s), <=24h past-day (5min), <=168h past-week (30min),
    >168h past-month (2h, com warning).

.PARAMETER Cluster
    Filtro opcional por nome de cluster. Aceita lista.

.PARAMETER Datacenter
    Filtro opcional por nome de datacenter. Aceita lista.

.PARAMETER VMName
    Filtro opcional por nome de VM. Aceita lista e wildcards.

.PARAMETER VMHost
    Filtro opcional por nome de host ESXi. Aceita lista.

.PARAMETER OutputPath
    Diretorio de saida. Default ./output. Criado se nao existir.

.PARAMETER SkipHosts
    Pula coleta de metricas em nivel de host.

.PARAMETER SkipVMs
    Pula coleta de metricas em nivel de VM.

.PARAMETER TopN
    Quantas VMs criticas exibir nas tabelas do console. Default 20.

.PARAMETER EdgePattern
    Regex para identificar NSX Edges pelo nome. Default 'edge|nsx-edge|nsxe'.
    Edges recebem regua mais conservadora (5%/2% em vez de 10%/5%).

.PARAMETER CompareWith
    Caminho para JSON de execucao anterior. Quando fornecido, gera CSV com
    deltas de Contention/Ready/Co-stop por VM e veredicto MELHOROU/PIOROU/ESTAVEL.

.PARAMETER VRopsServer
    FQDN ou IP do servidor Aria Operations (vROps). Habilita enriquecimento
    historico das VMs problematicas.

.PARAMETER VRopsCredential
    PSCredential para autenticacao no vROps.

.PARAMETER VRopsAuthSource
    Nome do auth source no vROps. Default 'LOCAL'. Em ambientes vIDM o nome
    costuma ser o FQDN do vIDM. Use -ListVRopsAuthSources para descobrir.

.PARAMETER VRopsHistoryDays
    Quantos dias de historico buscar no vROps. Default 30.

.PARAMETER ListVRopsAuthSources
    Modo helper: lista todos os auth sources configurados no vROps e sai.
    Pede credencial admin LOCAL via Get-Credential. Util em ambientes com
    vIDM / Workspace ONE Access onde o nome do auth source nao eh obvio.

.PARAMETER OraclePattern
    Regex para identificar VMs Oracle pelo nome. Default fallback generico
    '\b(oracle|oradb|orcl\d*|oraebs|oradg)\b'. Customize para a convencao
    do cliente. Exemplos: 'CRMDB', 'CRMDB|FINAPP|ORCL', '^ORA',
    '\bORADB\d+\b', '^(PRD|HML|DEV)-ORA'.

.PARAMETER OracleDeepAnalysis
    Ativa coleta extra de metricas de memoria, storage e network para VMs
    Oracle (ou para todas, se nenhuma VM for identificada como Oracle).
    Adiciona ~30% no tempo de coleta.

.PARAMETER OracleVMTag
    Lista de tags do vCenter que marcam VM como Oracle. Aceita formato
    'TagName' ou 'Categoria=TagName'. Alternativa/complemento ao -OraclePattern.

.PARAMETER BurstyThreshold
    Delta entre Max e Avg de Contention que caracteriza padrao bursty.
    Default 20 (ou seja, Max - Avg > 20pp). Ambientes sensiveis podem usar 10.

.PARAMETER ListOracleMatches
    Modo dry-run: lista quais VMs seriam classificadas como Oracle pelo
    pattern/tag/notes atual e quais sao "possivelmente Oracle" (vCPU>=8,
    RAM>=16GB, nome com db/database/rdbms). Sai sem coletar metricas.
    Use para validar a deteccao antes da coleta completa.

.EXAMPLE
    .\Get-VMwareCPUContention.ps1 -VCenter vc01.lab.local

    Coleta basica da ultima hora em todo o inventario do vCenter.

.EXAMPLE
    $vrCred = Get-Credential
    .\Get-VMwareCPUContention.ps1 -VCenter vc01 -Cluster "Prod-DB" -Hours 24 `
        -VRopsServer vrops01 -VRopsCredential $vrCred -VRopsHistoryDays 60

    Cluster especifico, janela de 24h, enriquecimento via vROps com 60 dias
    de historico.

.EXAMPLE
    .\Get-VMwareCPUContention.ps1 -VRopsServer vrops01 -ListVRopsAuthSources

    Descobre os auth sources do vROps. Util quando autenticacao via vIDM
    falha porque o nome do auth source eh diferente do esperado.

.EXAMPLE
    .\Get-VMwareCPUContention.ps1 -VCenter vc01 -Hours 1 `
        -CompareWith ./baseline/cpucontention-20260420-140000.json

    Compara janela atual com baseline anterior, gerando CSV de deltas e
    destacando no console as VMs que pioraram.

.EXAMPLE
    .\Get-VMwareCPUContention.ps1 -VCenter vc01 `
        -OraclePattern 'CRMDB' -OracleDeepAnalysis -Hours 24

    Cliente cuja convencao Oracle eh CRMDB. Ativa analise profunda
    (memoria, storage, network) e regua especializada Oracle.

.EXAMPLE
    .\Get-VMwareCPUContention.ps1 -VCenter vc01 `
        -OraclePattern 'CRMDB|FINAPP' -ListOracleMatches

    Dry-run: lista quais VMs seriam classificadas como Oracle e quais
    sao candidatas a revisao manual, sem coletar metricas. Use sempre
    antes da analise completa para validar o pattern.

.EXAMPLE
    .\Get-VMwareCPUContention.ps1 -VCenter vc01 `
        -OracleVMTag "Workload=Oracle","Type=Database" `
        -OracleDeepAnalysis -Hours 4

    Identifica VMs Oracle por tag do vCenter (alternativa a regex).

.EXAMPLE
    $cred = Get-Credential
    .\Get-VMwareCPUContention.ps1 -VCenter vc01 `
        -OraclePattern 'CRMDB|FINAPP' -OracleDeepAnalysis `
        -VRopsServer vrops01 -VRopsCredential $cred -VRopsHistoryDays 60 `
        -CompareWith ./baseline.json

    Combinacao completa: regua Oracle + vROps com 60d + comparacao
    com baseline anterior.

.NOTES
    Requisitos:
      - PowerShell 5.1+ ou 7+
      - VMware.PowerCLI 12+
      - vSphere 6.5+
      - vROps 7.5+ (opcional)

    Operacao read-only. Nao modifica nada no vCenter ou no vROps.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$VCenter,

    [pscredential]$VCenterCredential,

    [int]$Hours = 1,

    [string[]]$Cluster,
    [string[]]$Datacenter,
    [string[]]$VMName,
    [string[]]$VMHost,

    [string]$OutputPath = "./output",

    [switch]$SkipHosts,
    [switch]$SkipVMs,

    [int]$TopN = 20,

    [string]$EdgePattern = 'edge|nsx-edge|nsxe',

    [string]$CompareWith,

    # vROps opcional
    [string]$VRopsServer,
    [pscredential]$VRopsCredential,
    [string]$VRopsAuthSource = 'LOCAL',
    [int]$VRopsHistoryDays = 30,
    [switch]$ListVRopsAuthSources,

    # Oracle Database analysis
    [string]$OraclePattern = '\b(oracle|oradb|orcl\d*|oraebs|oradg)\b',
    [switch]$OracleDeepAnalysis,
    [string[]]$OracleVMTag,
    [int]$BurstyThreshold = 20,
    [switch]$ListOracleMatches
)

#region Setup inicial e validacao -------------------------------------------------

$ErrorActionPreference = 'Stop'
$ProgressPreference    = 'SilentlyContinue'

# Validacao: VCenter eh obrigatorio salvo modo helper
if (-not $ListVRopsAuthSources -and -not $VCenter) {
    Write-Error "Parametro -VCenter eh obrigatorio (exceto quando -ListVRopsAuthSources)."
    exit 1
}
if ($ListVRopsAuthSources -and -not $VRopsServer) {
    Write-Error "Para usar -ListVRopsAuthSources, informe -VRopsServer."
    exit 1
}
if ($ListOracleMatches -and -not $VCenter) {
    Write-Error "Para usar -ListOracleMatches, informe -VCenter."
    exit 1
}

# OutputPath
if (-not (Test-Path -LiteralPath $OutputPath)) {
    New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null
}
$OutputPath = (Resolve-Path -LiteralPath $OutputPath).Path

# Run label
$runTimestamp = Get-Date
$runLabel     = "cpucontention-{0:yyyyMMdd-HHmmss}" -f $runTimestamp

#endregion

#region Helpers - certificados e Invoke-RestMethod compatibility ------------------

function Set-CertificatePolicy {
    # PS 5.1: aceita certs auto-assinados via TrustAllCertsPolicy + TLS 1.2
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
        [Parameter(Mandatory)][ValidateSet('GET','POST','PUT','DELETE')]
        [string]$Method,
        [Parameter(Mandatory)][string]$Uri,
        [object]$Body,
        [hashtable]$Headers
    )
    if (-not $Headers) { $Headers = @{} }
    if (-not $Headers.ContainsKey('Accept'))      { $Headers['Accept']      = 'application/json' }
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

#endregion

#region Helpers - vROps API ------------------------------------------------------

function Connect-VRops {
    param(
        [Parameter(Mandatory)][string]$Server,
        [Parameter(Mandatory)][pscredential]$Credential,
        [Parameter(Mandatory)][string]$AuthSource
    )
    $body = @{
        username   = $Credential.UserName
        password   = $Credential.GetNetworkCredential().Password
        authSource = $AuthSource
    }
    $uri = "https://$Server/suite-api/api/auth/token/acquire"
    try {
        $resp = Invoke-VRopsApi -Method POST -Uri $uri -Body $body
    } catch {
        $statusCode = $null
        if ($_.Exception.Response) { $statusCode = [int]$_.Exception.Response.StatusCode }
        if ($statusCode -eq 401 -or $statusCode -eq 403) {
            Write-Warning "Falha de autenticacao no vROps ($statusCode). Verifique:"
            Write-Warning "  1. Senha correta para o usuario '$($Credential.UserName)'"
            Write-Warning "  2. Nome do auth source: '$AuthSource'"
            Write-Warning "     - Em ambientes com vIDM, costuma ser o FQDN do vIDM"
            Write-Warning "     - Para descobrir: re-execute com -ListVRopsAuthSources"
            Write-Warning "  3. Permissoes do usuario para a Suite API"
        } else {
            Write-Warning "Erro conectando ao vROps em $Server : $($_.Exception.Message)"
        }
        throw
    }
    return [PSCustomObject]@{
        Server  = $Server
        Token   = $resp.token
        Headers = @{ 'Authorization' = "vRealizeOpsToken $($resp.token)" }
    }
}

function Disconnect-VRops {
    param([Parameter(Mandatory)]$Session)
    if (-not $Session) { return }
    $uri = "https://$($Session.Server)/suite-api/api/auth/token/release"
    try {
        Invoke-VRopsApi -Method POST -Uri $uri -Headers $Session.Headers | Out-Null
    } catch {
        Write-Verbose "Falha ao liberar token vROps: $($_.Exception.Message)"
    }
}

function Get-VRopsAuthSources {
    param(
        [Parameter(Mandatory)][string]$Server,
        [Parameter(Mandatory)][pscredential]$Credential
    )
    Set-CertificatePolicy
    Write-Host ""
    Write-Host "Conectando como admin LOCAL para listar auth sources..." -ForegroundColor Cyan
    $session = Connect-VRops -Server $Server -Credential $Credential -AuthSource 'LOCAL'
    try {
        $uri    = "https://$Server/suite-api/api/auth/sources"
        $resp   = Invoke-VRopsApi -Method GET -Uri $uri -Headers $session.Headers
        $items  = @($resp.sources)
        if ($items.Count -eq 0) {
            Write-Warning "Nenhum auth source retornado."
            return
        }
        Write-Host ""
        Write-Host "Auth sources configurados em $Server :" -ForegroundColor Green
        Write-Host ""
        foreach ($s in $items) {
            $name     = $s.name
            $type     = $s.sourceType.name
            $hint     = ''
            if ($type -match 'IDM') { $hint = '  (vIDM/Workspace ONE Access)' }
            $line = "  {0,-32}  tipo: {1,-22}  Use: -VRopsAuthSource '{2}'{3}" -f $name, $type, $name, $hint
            Write-Host $line
        }
        Write-Host ""
    } finally {
        Disconnect-VRops -Session $session
    }
}

function Get-VRopsResource {
    param(
        [Parameter(Mandatory)]$Session,
        [Parameter(Mandatory)][string]$Name
    )
    $encoded = [uri]::EscapeDataString($Name)
    $uri = "https://$($Session.Server)/suite-api/api/resources?name=$encoded&resourceKind=VirtualMachine"
    try {
        $resp = Invoke-VRopsApi -Method GET -Uri $uri -Headers $Session.Headers
        $items = @($resp.resourceList)
        if ($items.Count -eq 0) { return $null }
        return $items[0]
    } catch {
        Write-Verbose "Resource '$Name' nao encontrado no vROps: $($_.Exception.Message)"
        return $null
    }
}

function Get-VRopsStats {
    param(
        [Parameter(Mandatory)]$Session,
        [Parameter(Mandatory)][string]$ResourceId,
        [Parameter(Mandatory)][string[]]$StatKeys,
        [Parameter(Mandatory)][int]$HistoryDays
    )
    $endMs   = [DateTimeOffset]::UtcNow.ToUnixTimeMilliseconds()
    $beginMs = [DateTimeOffset]::UtcNow.AddDays(-$HistoryDays).ToUnixTimeMilliseconds()
    $body = @{
        resourceId         = @($ResourceId)
        statKey            = $StatKeys
        begin              = $beginMs
        end                = $endMs
        rollUpType         = 'AVG'
        intervalType       = 'HOURS'
        intervalQuantifier = 1
    }
    $uri    = "https://$($Session.Server)/suite-api/api/resources/stats/query"
    $result = @{}
    foreach ($k in $StatKeys) { $result[$k] = $null }
    try {
        $resp = Invoke-VRopsApi -Method POST -Uri $uri -Headers $Session.Headers -Body $body
        $values = @($resp.values)
        if ($values.Count -eq 0) { return $result }
        foreach ($v in $values) {
            $statList = @($v.'stat-list'.stat)
            foreach ($s in $statList) {
                $key = $s.statKey.key
                $data = @($s.data)
                if ($data.Count -eq 0) { continue }
                $avg    = ($data | Measure-Object -Average).Average
                $max    = ($data | Measure-Object -Maximum).Maximum
                $min    = ($data | Measure-Object -Minimum).Minimum
                $latest = $data[-1]
                $result[$key] = [PSCustomObject]@{
                    Avg    = [Math]::Round($avg, 2)
                    Max    = [Math]::Round($max, 2)
                    Min    = [Math]::Round($min, 2)
                    Latest = [Math]::Round($latest, 2)
                    Count  = $data.Count
                }
            }
        }
    } catch {
        Write-Verbose "Falha em stats/query para $ResourceId : $($_.Exception.Message)"
    }
    return $result
}

function Get-VRopsProperties {
    param(
        [Parameter(Mandatory)]$Session,
        [Parameter(Mandatory)][string]$ResourceId
    )
    $uri = "https://$($Session.Server)/suite-api/api/resources/$ResourceId/properties"
    $out = [PSCustomObject]@{
        RecommendedVCPU  = $null
        RecommendedRAMGB = $null
    }
    try {
        $resp = Invoke-VRopsApi -Method GET -Uri $uri -Headers $Session.Headers
        $props = @($resp.property)
        foreach ($p in $props) {
            switch ($p.name) {
                'summary|recommendedConfig|cpu' {
                    if ($p.value) { $out.RecommendedVCPU = [int][double]$p.value }
                }
                'summary|recommendedConfig|memory' {
                    if ($p.value) { $out.RecommendedRAMGB = [Math]::Round([double]$p.value / 1024, 2) }
                }
            }
        }
    } catch {
        Write-Verbose "Falha em properties para $ResourceId : $($_.Exception.Message)"
    }
    return $out
}

#endregion

#region Branch helper -ListVRopsAuthSources --------------------------------------

if ($ListVRopsAuthSources) {
    $cred = Get-Credential -Message "Credenciais admin LOCAL do vROps em $VRopsServer"
    if (-not $cred) {
        Write-Error "Credenciais nao informadas."
        exit 1
    }
    Get-VRopsAuthSources -Server $VRopsServer -Credential $cred
    exit 0
}

#endregion

#region Helpers - calculo de metricas --------------------------------------------

function Convert-SummationToPercent {
    param(
        [double]$ValueMs,
        [double]$IntervalSeconds,
        [int]$NumVcpus
    )
    if ($null -eq $ValueMs -or $IntervalSeconds -le 0 -or $NumVcpus -le 0) { return 0 }
    return [Math]::Round(($ValueMs / ($IntervalSeconds * 1000 * $NumVcpus)) * 100, 2)
}

function Test-IsEdgeVM {
    param(
        [string]$Name,
        [string]$Pattern
    )
    if ([string]::IsNullOrWhiteSpace($Name) -or [string]::IsNullOrWhiteSpace($Pattern)) { return $false }
    return ($Name -match $Pattern)
}

function Get-ContentionClass {
    param(
        [bool]$IsEdge,
        [double]$ContentionPct,
        [double]$ReadyPctPerVcpu,
        [double]$CostopPct,
        [double]$MaxLimitedPct,
        [double]$SwapWaitPct
    )
    if ($IsEdge) {
        $crit = 5;  $warn = 2
    } else {
        $crit = 10; $warn = 5
    }
    $level    = 'OK'
    $findings = New-Object System.Collections.ArrayList
    $worst    = [Math]::Max([double]$ContentionPct, [double]$ReadyPctPerVcpu)
    if ($worst -ge $crit) {
        $level = 'CRITICO'
    } elseif ($worst -ge $warn) {
        $level = 'ATENCAO'
    }
    if ($ContentionPct -ge $warn) {
        [void]$findings.Add("Contention agregada $([Math]::Round($ContentionPct,2))%")
    }
    if ($ReadyPctPerVcpu -ge $warn) {
        [void]$findings.Add("Ready $([Math]::Round($ReadyPctPerVcpu,2))% por vCPU")
    }
    if ($CostopPct -gt 3) {
        [void]$findings.Add("Co-stop ($([Math]::Round($CostopPct,2))%) - VM SMP grande demais")
        if ($level -eq 'OK') { $level = 'ATENCAO' }
    }
    if ($MaxLimitedPct -gt 0) {
        [void]$findings.Add("CPU Limit ativo (%MLMTD $([Math]::Round($MaxLimitedPct,2))%)")
        if ($level -eq 'OK') { $level = 'ATENCAO' }
    }
    if ($SwapWaitPct -gt 0) {
        [void]$findings.Add("Swap wait ($([Math]::Round($SwapWaitPct,2))%) - pressao de memoria")
        if ($level -eq 'OK') { $level = 'ATENCAO' }
    }
    if ($IsEdge -and $level -ne 'OK') {
        [void]$findings.Add("Avaliado como NSX Edge (regua conservadora)")
    }
    return [PSCustomObject]@{
        Level    = $level
        Findings = ($findings -join '; ')
    }
}

function Get-CPUStats {
    param(
        [Parameter(Mandatory)]$Entity,
        [Parameter(Mandatory)][string]$StatName,
        [Parameter(Mandatory)][hashtable]$IntervalInfo,
        [Parameter(Mandatory)][int]$Hours
    )
    try {
        $statParams = @{
            Entity      = $Entity
            Stat        = $StatName
            ErrorAction = 'Stop'
        }
        if ($IntervalInfo.Realtime) {
            $samples = [int][Math]::Ceiling(($Hours * 3600) / 20.0)
            if ($samples -lt 1)   { $samples = 1 }
            if ($samples -gt 360) { $samples = 360 }   # vCenter retem ~1h em real-time
            $statParams.Realtime   = $true
            $statParams.MaxSamples = $samples
        } else {
            $statParams.Start  = (Get-Date).AddHours(-$Hours)
            $statParams.Finish = Get-Date
        }
        $stats = Get-Stat @statParams
        if (-not $stats) { return $null }
        $values = @($stats | Where-Object { $null -ne $_.Value } | Select-Object -ExpandProperty Value)
        if ($values.Count -eq 0) { return $null }
        $avg    = ($values | Measure-Object -Average).Average
        $max    = ($values | Measure-Object -Maximum).Maximum
        $sorted = $values | Sort-Object
        $idx    = [int][Math]::Floor($sorted.Count * 0.95)
        if ($idx -ge $sorted.Count) { $idx = $sorted.Count - 1 }
        if ($idx -lt 0)             { $idx = 0 }
        $p95 = $sorted[$idx]
        return [PSCustomObject]@{
            Avg   = [Math]::Round([double]$avg, 2)
            Max   = [Math]::Round([double]$max, 2)
            P95   = [Math]::Round([double]$p95, 2)
            Count = $sorted.Count
        }
    } catch {
        Write-Verbose "Get-Stat falhou para '$StatName' em '$($Entity.Name)': $($_.Exception.Message)"
        return $null
    }
}

#endregion

#region Helpers - Oracle Database ------------------------------------------------

# Cache de NUMA por host (key = host name)
$script:HostNumaCache = @{}

function Get-HostNumaInfo {
    param([Parameter(Mandatory)]$VMHostObj)
    if ($script:HostNumaCache.ContainsKey($VMHostObj.Name)) {
        return $script:HostNumaCache[$VMHostObj.Name]
    }
    $threads = [int]$VMHostObj.NumCpu
    $sockets = 0
    if ($VMHostObj.ExtensionData -and $VMHostObj.ExtensionData.Hardware -and $VMHostObj.ExtensionData.Hardware.CpuInfo) {
        $sockets = [int]$VMHostObj.ExtensionData.Hardware.CpuInfo.NumCpuPackages
    }
    $numNodes = if ($sockets -gt 0) { $sockets } else { 1 }
    try {
        if ($VMHostObj.ExtensionData -and $VMHostObj.ExtensionData.Hardware -and $VMHostObj.ExtensionData.Hardware.NumaInfo) {
            $ni = $VMHostObj.ExtensionData.Hardware.NumaInfo
            if ($ni.NumNodes -and [int]$ni.NumNodes -gt 0) { $numNodes = [int]$ni.NumNodes }
        }
    } catch {
        Write-Verbose "Falha lendo NumaInfo de '$($VMHostObj.Name)': $($_.Exception.Message)"
    }
    if ($numNodes -le 0) { $numNodes = 1 }
    $threadsPerNode = [int][Math]::Floor($threads / $numNodes)
    $totalRamGB     = [Math]::Round([double]$VMHostObj.MemoryTotalGB, 2)
    $ramPerNodeGB   = [Math]::Round($totalRamGB / $numNodes, 2)
    $info = [PSCustomObject]@{
        Threads        = $threads
        Sockets        = $sockets
        NumNodes       = $numNodes
        ThreadsPerNode = $threadsPerNode
        TotalRamGB     = $totalRamGB
        RamPerNodeGB   = $ramPerNodeGB
    }
    $script:HostNumaCache[$VMHostObj.Name] = $info
    return $info
}

function Test-IsOracleVM {
    param(
        [Parameter(Mandatory)]$VM,
        [string]$Pattern,
        [string[]]$Tags
    )
    # 1. Tag (alta confianca)
    if ($Tags -and $Tags.Count -gt 0) {
        try {
            $vmTags = @(Get-TagAssignment -Entity $VM -ErrorAction SilentlyContinue)
            foreach ($t in $vmTags) {
                $tagName = "$($t.Tag.Name)"
                $catTagName = ''
                if ($t.Tag.Category) { $catTagName = "$($t.Tag.Category.Name)=$($t.Tag.Name)" }
                foreach ($wanted in $Tags) {
                    if ($tagName -eq $wanted -or $catTagName -eq $wanted) {
                        return [PSCustomObject]@{ IsOracle = $true; DetectionMethod = 'tag' }
                    }
                }
            }
        } catch {
            Write-Verbose "Falha em Get-TagAssignment para '$($VM.Name)': $($_.Exception.Message)"
        }
    }
    # 2. Nome (regex)
    if ($Pattern -and ($VM.Name -match $Pattern)) {
        return [PSCustomObject]@{ IsOracle = $true; DetectionMethod = 'name_pattern' }
    }
    # 3. Notes / annotation
    if ($VM.Notes -and ($VM.Notes -match '(?i)oracle|database|rdbms')) {
        return [PSCustomObject]@{ IsOracle = $true; DetectionMethod = 'notes' }
    }
    return [PSCustomObject]@{ IsOracle = $false; DetectionMethod = 'none' }
}

function Test-IsPossibleOracleVM {
    param(
        [Parameter(Mandatory)]$VM,
        [bool]$AlreadyOracle
    )
    if ($AlreadyOracle) { return $false }
    $numVcpu  = [int]$VM.NumCpu
    $memGB    = [double]$VM.MemoryGB
    $nameHit  = ($VM.Name -match '(?i)\b(db|database|rdbms)\b')
    return ($numVcpu -ge 8 -and $memGB -ge 16 -and $nameHit)
}

function Get-VMDisparity {
    param(
        [Parameter(Mandatory)]$VM,
        [Parameter(Mandatory)][hashtable]$IntervalInfo,
        [Parameter(Mandatory)][int]$Hours
    )
    try {
        $params = @{
            Entity      = $VM
            Stat        = 'cpu.usage.average'
            Instance    = '*'
            ErrorAction = 'Stop'
        }
        if ($IntervalInfo.Realtime) {
            $samples = [int][Math]::Ceiling(($Hours * 3600) / 20.0)
            if ($samples -lt 1)   { $samples = 1 }
            if ($samples -gt 360) { $samples = 360 }
            $params.Realtime   = $true
            $params.MaxSamples = $samples
        } else {
            $params.Start  = (Get-Date).AddHours(-$Hours)
            $params.Finish = Get-Date
        }
        $stats = Get-Stat @params
        if (-not $stats) { return $null }
        # Agrupa por timestamp e considera so instances especificas (vCPUs individuais)
        $perVcpu = $stats | Where-Object { -not [string]::IsNullOrEmpty($_.Instance) }
        if (-not $perVcpu) { return $null }
        $grouped = $perVcpu | Group-Object Timestamp
        $disparities = New-Object System.Collections.ArrayList
        foreach ($g in $grouped) {
            $vals = @($g.Group | Select-Object -ExpandProperty Value)
            if ($vals.Count -lt 2) { continue }
            $vMax = ($vals | Measure-Object -Maximum).Maximum
            $vMin = ($vals | Measure-Object -Minimum).Minimum
            $vAvg = ($vals | Measure-Object -Average).Average
            if ($vAvg -le 0) { continue }
            [void]$disparities.Add((($vMax - $vMin) / $vAvg) * 100)
        }
        if ($disparities.Count -eq 0) { return $null }
        $a = ($disparities | Measure-Object -Average).Average
        $m = ($disparities | Measure-Object -Maximum).Maximum
        return [PSCustomObject]@{
            Avg = [Math]::Round([double]$a, 2)
            Max = [Math]::Round([double]$m, 2)
        }
    } catch {
        Write-Verbose "Falha em Get-VMDisparity para '$($VM.Name)': $($_.Exception.Message)"
        return $null
    }
}

function Get-VMDeepStats {
    param(
        [Parameter(Mandatory)]$VM,
        [Parameter(Mandatory)][hashtable]$IntervalInfo,
        [Parameter(Mandatory)][int]$Hours
    )
    $deepNames = @(
        'mem.active.average', 'mem.consumed.average', 'mem.granted.average',
        'mem.shared.average', 'mem.swapped.average', 'mem.swapout.average',
        'mem.vmmemctl.average', 'mem.latency.average', 'mem.compressed.average',
        'disk.maxTotalLatency.latest', 'disk.usage.average',
        'disk.commandsAborted.summation',
        'virtualDisk.totalReadLatency.average',
        'virtualDisk.totalWriteLatency.average',
        'net.droppedRx.summation', 'net.droppedTx.summation',
        'net.usage.average'
    )
    $out = @{}
    foreach ($n in $deepNames) {
        $out[$n] = Get-CPUStats -Entity $VM -StatName $n -IntervalInfo $IntervalInfo -Hours $Hours
    }
    return $out
}

function Get-OracleHealthClass {
    param(
        [Parameter(Mandatory)][hashtable]$D
    )
    $level    = 'OK'
    $findings = New-Object System.Collections.ArrayList

    # CRITICO
    $critical = $false
    if ($D.IsOversized) {
        [void]$findings.Add("Oversized: $($D.NumVcpu) vCPUs em host de $($D.HostThreads) threads")
        $critical = $true
    }
    if ([double]$D.ContentionMax -gt 30) {
        [void]$findings.Add("Pico severo de Contention: $([Math]::Round([double]$D.ContentionMax,2))%")
        $critical = $true
    }
    if ([double]$D.CostopMax -gt 15) {
        [void]$findings.Add("Co-stop pico de $([Math]::Round([double]$D.CostopMax,2))% indica gang scheduling falhando")
        $critical = $true
    }
    if ([double]$D.MemLatencyAvg -gt 10) {
        [void]$findings.Add("Memory latency $([Math]::Round([double]$D.MemLatencyAvg,2))% pode degradar SGA hits")
        $critical = $true
    }
    if ([double]$D.MemBalloonedMB -gt 0) {
        [void]$findings.Add("Ballooning ativo ($([Math]::Round([double]$D.MemBalloonedMB,0)) MB) - host com pressao de memoria")
        $critical = $true
    }
    if ([double]$D.MemSwappedMB -gt 0 -or [double]$D.MemCompressedMB -gt 0) {
        [void]$findings.Add("Memoria swapped/compressed - catastrofico para Oracle")
        $critical = $true
    }
    if ($D.DisparityAvg -ne $null -and [double]$D.DisparityAvg -gt 60) {
        [void]$findings.Add("Disparity alta ($([Math]::Round([double]$D.DisparityAvg,2))%) sugere workload mal paralelizado - revisar PARALLEL hints e parallel_max_servers")
        $critical = $true
    }
    if ($D.CrossesNUMABoundary -and [int]$D.NumVcpu -gt 4) {
        [void]$findings.Add("Cruza fronteira NUMA: $($D.NumVcpu) vCPUs / $([Math]::Round([double]$D.MemoryGB,0)) GB excedem $($D.ThreadsPerNode) threads / $([Math]::Round([double]$D.RamPerNodeGB,0)) GB por no")
        $critical = $true
    }
    if ($critical) { $level = 'CRITICO' }

    # ATENCAO (so se nao for CRITICO ainda)
    if (-not $critical) {
        $warn = $false
        if ($D.IsBursty) {
            [void]$findings.Add("Bursty: Contention oscila entre $([Math]::Round([double]$D.ContentionAvg,2))% (avg) e $([Math]::Round([double]$D.ContentionMax,2))% (max)")
            $warn = $true
        }
        if ($D.DisparityAvg -ne $null -and [double]$D.DisparityAvg -gt 30 -and [double]$D.DisparityAvg -le 60) {
            [void]$findings.Add("Disparity moderada ($([Math]::Round([double]$D.DisparityAvg,2))%) - revisar paralelismo Oracle")
            $warn = $true
        }
        if ([double]$D.DemandAvgPct -gt 80) {
            [void]$findings.Add("Demand sustentado acima de 80% - possivel que SQLs precisem de tuning")
            $warn = $true
        }
        if ([double]$D.DiskLatencyP95 -gt 20) {
            [void]$findings.Add("Storage latency P95 $([Math]::Round([double]$D.DiskLatencyP95,2)) ms - LGWR pode estar travando commits")
            $warn = $true
        }
        if ([double]$D.DiskCommandsAborted -gt 0) {
            [void]$findings.Add("Disk commands aborted ($([int]$D.DiskCommandsAborted)) - problemas de storage")
            $warn = $true
        }
        if ([double]$D.NetDropped -gt 0) {
            [void]$findings.Add("Network drops ($([int]$D.NetDropped)) - revisar throughput e queues")
            $warn = $true
        }
        if ($D.IsMonster) {
            [void]$findings.Add("VM monster (vCPU >= host threads) - alto risco de gang scheduling")
            $warn = $true
        }
        if ($warn) { $level = 'ATENCAO' }
    }

    # Recommendation (regra de prioridade)
    $reco = ''
    $disparityHigh = ($D.DisparityAvg -ne $null -and [double]$D.DisparityAvg -gt 60)
    if ($D.IsOversized -and $disparityHigh) {
        $reco = "Reduzir para metade dos vCPUs atuais. Workload nao se beneficia do paralelismo configurado. Validar parallel_max_servers no Oracle."
    } elseif ($D.IsOversized) {
        $reco = "Reduzir para no maximo $($D.HostThreads) vCPUs. Considerar -1 para reservar core ao SO."
    } elseif ($D.CrossesNUMABoundary -and [double]$D.MemoryGB -ge 32) {
        $reco = "Considerar dividir em multiplas VMs menores ou migrar para host com nos NUMA maiores. SGA cruzando fronteira NUMA degrada cache hit ratio."
    } elseif ([double]$D.MemBalloonedMB -gt 0 -or [double]$D.MemSwappedMB -gt 0) {
        $reco = "Pressao de memoria no host. Validar reservation de memoria da VM Oracle. SGA em swap = catastrofe de performance."
    } elseif ([double]$D.DiskLatencyP95 -gt 20) {
        $reco = "LGWR pode estar travando commits. Validar latencia de redo log files. Considerar mover redo para storage mais rapido."
    } elseif ($D.IsBursty) {
        $reco = "Investigar pico em janela de tempo. Cruzar com AWR/Statspack do Oracle no horario do pico."
    } elseif ([double]$D.DemandAvgPct -gt 80 -and [double]$D.ContentionAvg -lt 5) {
        $reco = "VM esta entregando o que precisa, mas com demanda alta sustentada. Investigar tuning de SQL no Oracle."
    }

    return [PSCustomObject]@{
        Level          = $level
        Findings       = ($findings -join '; ')
        Recommendation = $reco
    }
}

#endregion

#region Granularidade adaptativa -------------------------------------------------

if ($Hours -le 1) {
    $intervalInfo = @{ Name = 'real-time'; IntervalSecs = 20;   Realtime = $true;  Note = '' }
} elseif ($Hours -le 24) {
    $intervalInfo = @{ Name = 'past-day';  IntervalSecs = 300;  Realtime = $false; Note = '' }
} elseif ($Hours -le 168) {
    $intervalInfo = @{ Name = 'past-week'; IntervalSecs = 1800; Realtime = $false; Note = '' }
} else {
    $intervalInfo = @{
        Name = 'past-month'; IntervalSecs = 7200; Realtime = $false
        Note = 'Janela longa - rollup de 2h nao recomendado para diagnostico de contention pontual'
    }
    Write-Warning $intervalInfo.Note
}

#endregion

#region PowerCLI - import e configuracao -----------------------------------------

try {
    if (-not (Get-Module -ListAvailable -Name VMware.PowerCLI)) {
        throw "Modulo VMware.PowerCLI nao encontrado. Instale com: Install-Module VMware.PowerCLI -Scope CurrentUser"
    }
    Import-Module VMware.PowerCLI -ErrorAction Stop | Out-Null
    Set-PowerCLIConfiguration -InvalidCertificateAction Ignore -ParticipateInCEIP $false -Scope Session -Confirm:$false | Out-Null
} catch {
    Write-Error $_.Exception.Message
    exit 1
}

#endregion

#region Conexoes -----------------------------------------------------------------

Set-CertificatePolicy

Write-Host ""
Write-Host "Conectando em vCenter $VCenter ..." -ForegroundColor Cyan
if (-not $VCenterCredential) {
    $VCenterCredential = Get-Credential -Message "Credenciais vCenter ($VCenter) - use UPN administrator@vsphere.local ou DOMAIN\user"
    if (-not $VCenterCredential) {
        Write-Error "Credenciais vCenter nao informadas."
        exit 1
    }
}
try {
    $vcConnection = Connect-VIServer -Server $VCenter -Credential $VCenterCredential -ErrorAction Stop
} catch {
    Write-Error "Falha conectando ao vCenter '$VCenter': $($_.Exception.Message)"
    Write-Warning "Cheque: (1) usuario UPN (user@vsphere.local) ou DOMAIN\user; (2) permissao de leitura no vSphere; (3) senha sem caracteres que o terminal possa estar comendo."
    exit 1
}
$vCenterVersion = "$($vcConnection.Version) build $($vcConnection.Build)"
Write-Host "  Conectado: $($vcConnection.Name)  versao $vCenterVersion" -ForegroundColor Green

# vROps opcional
$vRopsSession = $null
$vRopsEnabled = $false
if ($VRopsServer) {
    if (-not $VRopsCredential) {
        $VRopsCredential = Get-Credential -Message "Credenciais para vROps em $VRopsServer (auth source: $VRopsAuthSource)"
    }
    if ($VRopsCredential) {
        Write-Host "Conectando em vROps $VRopsServer (auth source: $VRopsAuthSource) ..." -ForegroundColor Cyan
        try {
            $vRopsSession = Connect-VRops -Server $VRopsServer -Credential $VRopsCredential -AuthSource $VRopsAuthSource
            $vRopsEnabled = $true
            Write-Host "  Conectado em vROps" -ForegroundColor Green
        } catch {
            Write-Warning "Seguindo sem enriquecimento vROps."
            $vRopsEnabled = $false
        }
    } else {
        Write-Warning "Credenciais vROps nao informadas - seguindo sem enriquecimento."
    }
}

#endregion

#region Filtros de inventario ----------------------------------------------------

# Resolve hosts
$hostFilter = @{}
if ($Datacenter) { $hostFilter.Location = (Get-Datacenter -Name $Datacenter -ErrorAction SilentlyContinue) }
$allHosts = @()
if ($SkipHosts -and $SkipVMs) {
    Write-Error "-SkipHosts e -SkipVMs juntos nao geram nada para coletar."
    exit 1
}

# Inventario de hosts (ainda que SkipHosts, precisamos para mapear VM->host->NumCpu)
$hostQuery = Get-VMHost @hostFilter
if ($Cluster) {
    $clusterObjs = Get-Cluster -Name $Cluster -ErrorAction SilentlyContinue
    if ($clusterObjs) {
        $clusterHosts = $clusterObjs | Get-VMHost
        $hostQuery = $hostQuery | Where-Object { $clusterHosts -contains $_ }
    }
}
if ($VMHost) {
    $hostQuery = $hostQuery | Where-Object {
        $h = $_
        ($VMHost | Where-Object { $h.Name -like $_ }).Count -gt 0
    }
}
$allHosts = @($hostQuery)

# Inventario de VMs
$vmQuery = $null
if (-not $SkipVMs) {
    $vmParams = @{}
    if ($Cluster -and $clusterObjs) {
        $vmParams.Location = $clusterObjs
    } elseif ($Datacenter -and $hostFilter.Location) {
        $vmParams.Location = $hostFilter.Location
    } elseif ($allHosts.Count -gt 0 -and ($VMHost -or $Cluster -or $Datacenter)) {
        $vmParams.Location = $allHosts
    }
    if ($VMName) {
        $vmParams.Name = $VMName
    }
    $vmQuery = @(Get-VM @vmParams | Where-Object { $_.PowerState -eq 'PoweredOn' })
    # Filtro adicional por host quando aplicavel
    if ($VMHost) {
        $vmQuery = $vmQuery | Where-Object {
            $h = $_.VMHost.Name
            ($VMHost | Where-Object { $h -like $_ }).Count -gt 0
        }
        $vmQuery = @($vmQuery)
    }
}

Write-Host ""
Write-Host ("Janela: {0}h  |  granularidade: {1} ({2}s)" -f $Hours, $intervalInfo.Name, $intervalInfo.IntervalSecs) -ForegroundColor Cyan
Write-Host ("Hosts no escopo: {0}  |  VMs no escopo: {1}" -f $allHosts.Count, ($vmQuery | Measure-Object).Count) -ForegroundColor Cyan
if ($vRopsEnabled) {
    Write-Host ("vROps: enriquecimento habilitado, historico {0}d" -f $VRopsHistoryDays) -ForegroundColor Cyan
}

#endregion

#region Branch helper -ListOracleMatches -----------------------------------------

if ($ListOracleMatches) {
    Write-Host ""
    Write-Host "===== Validacao de deteccao Oracle =====" -ForegroundColor Cyan
    Write-Host ("Pattern usado: {0}" -f $OraclePattern)
    if ($OracleVMTag) {
        Write-Host ("Tags consideradas: {0}" -f ($OracleVMTag -join ', '))
    }
    Write-Host ""

    $candidates = if ($vmQuery) { @($vmQuery) } else { @(Get-VM | Where-Object { $_.PowerState -eq 'PoweredOn' }) }

    $oracleHits   = New-Object System.Collections.ArrayList
    $possibleHits = New-Object System.Collections.ArrayList
    $i = 0
    $tot = $candidates.Count
    foreach ($vm in $candidates) {
        $i++
        Write-Progress -Activity "Avaliando deteccao Oracle" -Status $vm.Name -PercentComplete (($i / [Math]::Max($tot,1)) * 100)
        $det = Test-IsOracleVM -VM $vm -Pattern $OraclePattern -Tags $OracleVMTag
        $clName = ''
        try {
            $cl = $vm | Get-Cluster -ErrorAction SilentlyContinue
            if ($cl) { $clName = $cl.Name }
        } catch {}
        if ($det.IsOracle) {
            [void]$oracleHits.Add([PSCustomObject]@{
                Name     = $vm.Name
                vCPU     = [int]$vm.NumCpu
                MemoryGB = [Math]::Round([double]$vm.MemoryGB, 2)
                Cluster  = $clName
                Method   = $det.DetectionMethod
            })
        } elseif (Test-IsPossibleOracleVM -VM $vm -AlreadyOracle $false) {
            [void]$possibleHits.Add([PSCustomObject]@{
                Name     = $vm.Name
                vCPU     = [int]$vm.NumCpu
                MemoryGB = [Math]::Round([double]$vm.MemoryGB, 2)
                Cluster  = $clName
                Notes    = if ($vm.Notes) { ($vm.Notes -replace '\s+', ' ').Substring(0, [Math]::Min(80, $vm.Notes.Length)) } else { '' }
            })
        }
    }
    Write-Progress -Activity "Avaliando deteccao Oracle" -Completed

    Write-Host ("VMs identificadas como Oracle ({0}):" -f $oracleHits.Count) -ForegroundColor Green
    if ($oracleHits.Count -gt 0) {
        $oracleHits | Sort-Object Name | Format-Table Name, vCPU, MemoryGB, Cluster, Method -AutoSize | Out-Host
    } else {
        Write-Host "  (nenhuma)" -ForegroundColor Yellow
    }

    Write-Host ""
    Write-Host ("VMs possivelmente Oracle (nao identificadas mas com perfil de banco) ({0}):" -f $possibleHits.Count) -ForegroundColor Yellow
    if ($possibleHits.Count -gt 0) {
        $possibleHits | Sort-Object Name | Format-Table Name, vCPU, MemoryGB, Cluster, Notes -AutoSize -Wrap | Out-Host
    } else {
        Write-Host "  (nenhuma)" -ForegroundColor DarkGray
    }

    Write-Host ""
    Write-Host "Confirme com o cliente se a lista esta correta antes de rodar analise completa." -ForegroundColor Cyan

    if ($vcConnection) {
        Disconnect-VIServer -Server $vcConnection -Confirm:$false -ErrorAction SilentlyContinue | Out-Null
    }
    exit 0
}

#endregion

#region Coleta de VMs -------------------------------------------------------------

$vmResults = @()
if (-not $SkipVMs -and $vmQuery -and $vmQuery.Count -gt 0) {
    $idx = 0
    $total = $vmQuery.Count
    foreach ($vm in $vmQuery) {
        $idx++
        Write-Progress -Activity "Coletando metricas de VMs" -Status $vm.Name -PercentComplete (($idx / $total) * 100)

        $vmHostObj = $vm.VMHost
        $hostThreads = if ($vmHostObj) { [int]$vmHostObj.NumCpu } else { 0 }
        $numVcpu     = [int]$vm.NumCpu

        $isMonster   = ($hostThreads -gt 0 -and $numVcpu -ge $hostThreads)
        $isOversized = ($hostThreads -gt 0 -and $numVcpu -gt $hostThreads)
        $isEdge      = Test-IsEdgeVM -Name $vm.Name -Pattern $EdgePattern

        # Deteccao Oracle (multi-camada)
        $oracleDet = Test-IsOracleVM -VM $vm -Pattern $OraclePattern -Tags $OracleVMTag
        $isOracle  = [bool]$oracleDet.IsOracle
        $oracleMethod = "$($oracleDet.DetectionMethod)"
        $isPossibleOracle = Test-IsPossibleOracleVM -VM $vm -AlreadyOracle $isOracle

        # NUMA boundary
        $numaInfo = $null
        $crossesNuma = $false
        $threadsPerNumaNode = 0
        $ramPerNumaNodeGB   = 0
        if ($vmHostObj) {
            $numaInfo = Get-HostNumaInfo -VMHostObj $vmHostObj
            $threadsPerNumaNode = [int]$numaInfo.ThreadsPerNode
            $ramPerNumaNodeGB   = [double]$numaInfo.RamPerNodeGB
            if ($threadsPerNumaNode -gt 0 -and $numVcpu -gt $threadsPerNumaNode)        { $crossesNuma = $true }
            if ($ramPerNumaNodeGB -gt 0 -and [double]$vm.MemoryGB -gt $ramPerNumaNodeGB) { $crossesNuma = $true }
        }

        # Coleta das 8 metricas
        $usagePct  = Get-CPUStats -Entity $vm -StatName 'cpu.usage.average'      -IntervalInfo $intervalInfo -Hours $Hours
        $usageMHz  = Get-CPUStats -Entity $vm -StatName 'cpu.usagemhz.average'   -IntervalInfo $intervalInfo -Hours $Hours
        $demand    = Get-CPUStats -Entity $vm -StatName 'cpu.demand.average'     -IntervalInfo $intervalInfo -Hours $Hours
        $readyMs   = Get-CPUStats -Entity $vm -StatName 'cpu.ready.summation'    -IntervalInfo $intervalInfo -Hours $Hours
        $costopMs  = Get-CPUStats -Entity $vm -StatName 'cpu.costop.summation'   -IntervalInfo $intervalInfo -Hours $Hours
        $mlmtdMs   = Get-CPUStats -Entity $vm -StatName 'cpu.maxlimited.summation' -IntervalInfo $intervalInfo -Hours $Hours
        $swapMs    = Get-CPUStats -Entity $vm -StatName 'cpu.swapwait.summation' -IntervalInfo $intervalInfo -Hours $Hours
        $latency   = Get-CPUStats -Entity $vm -StatName 'cpu.latency.average'    -IntervalInfo $intervalInfo -Hours $Hours

        # Conversoes summation -> %
        $intervalSec = [double]$intervalInfo.IntervalSecs
        $readyAvgPct  = if ($readyMs)  { Convert-SummationToPercent -ValueMs $readyMs.Avg  -IntervalSeconds $intervalSec -NumVcpus $numVcpu } else { 0 }
        $readyMaxPct  = if ($readyMs)  { Convert-SummationToPercent -ValueMs $readyMs.Max  -IntervalSeconds $intervalSec -NumVcpus $numVcpu } else { 0 }
        $readyP95Pct  = if ($readyMs)  { Convert-SummationToPercent -ValueMs $readyMs.P95  -IntervalSeconds $intervalSec -NumVcpus $numVcpu } else { 0 }
        $costopAvgPct = if ($costopMs) { Convert-SummationToPercent -ValueMs $costopMs.Avg -IntervalSeconds $intervalSec -NumVcpus $numVcpu } else { 0 }
        $costopMaxPct = if ($costopMs) { Convert-SummationToPercent -ValueMs $costopMs.Max -IntervalSeconds $intervalSec -NumVcpus $numVcpu } else { 0 }
        $mlmtdAvgPct  = if ($mlmtdMs)  { Convert-SummationToPercent -ValueMs $mlmtdMs.Avg  -IntervalSeconds $intervalSec -NumVcpus $numVcpu } else { 0 }
        $mlmtdMaxPct  = if ($mlmtdMs)  { Convert-SummationToPercent -ValueMs $mlmtdMs.Max  -IntervalSeconds $intervalSec -NumVcpus $numVcpu } else { 0 }
        $swapAvgPct   = if ($swapMs)   { Convert-SummationToPercent -ValueMs $swapMs.Avg   -IntervalSeconds $intervalSec -NumVcpus $numVcpu } else { 0 }
        $swapMaxPct   = if ($swapMs)   { Convert-SummationToPercent -ValueMs $swapMs.Max   -IntervalSeconds $intervalSec -NumVcpus $numVcpu } else { 0 }

        $contentionAvg = if ($latency) { [double]$latency.Avg } else { 0 }
        $contentionMax = if ($latency) { [double]$latency.Max } else { 0 }
        $contentionP95 = if ($latency) { [double]$latency.P95 } else { 0 }

        $cls = Get-ContentionClass -IsEdge $isEdge `
            -ContentionPct $contentionAvg `
            -ReadyPctPerVcpu $readyAvgPct `
            -CostopPct $costopAvgPct `
            -MaxLimitedPct $mlmtdAvgPct `
            -SwapWaitPct $swapAvgPct

        # Bursty: delta entre Max e Avg de contention/ready/costop
        $contentionBurst = [Math]::Round($contentionMax - $contentionAvg, 2)
        $readyBurst      = [Math]::Round($readyMaxPct - $readyAvgPct, 2)
        $costopBurst     = [Math]::Round($costopMaxPct - $costopAvgPct, 2)
        $isBursty        = ($contentionBurst -gt $BurstyThreshold)

        # Disparity (custo extra: so coleta para Oracle, monster ou OracleDeepAnalysis global)
        $disparityAvg = $null
        $disparityMax = $null
        $needDisparity = ($isOracle -or $isMonster -or $OracleDeepAnalysis.IsPresent)
        if ($needDisparity) {
            $disp = Get-VMDisparity -VM $vm -IntervalInfo $intervalInfo -Hours $Hours
            if ($disp) {
                $disparityAvg = $disp.Avg
                $disparityMax = $disp.Max
            }
        }

        # Deep stats (memoria, storage, network) so para Oracle ou se flag global
        $memActiveGB = 0; $memConsumedGB = 0; $memGrantedGB = 0
        $memSharedMB = 0; $memSwappedMB  = 0; $memSwapoutMB = 0
        $memBalloonMB= 0; $memLatencyAvg = 0; $memCompressedMB = 0
        $diskMaxLatencyMs = 0; $diskUsageAvg = 0; $diskAborted = 0
        $vDiskReadLatencyMs = 0; $vDiskWriteLatencyMs = 0
        $netDroppedRx = 0; $netDroppedTx = 0; $netUsageAvg = 0
        $diskMaxLatencyP95 = 0
        $needDeep = ($isOracle -or $OracleDeepAnalysis.IsPresent)
        if ($needDeep) {
            $deep = Get-VMDeepStats -VM $vm -IntervalInfo $intervalInfo -Hours $Hours
            $tmp = $deep['mem.active.average'];        if ($tmp) { $memActiveGB     = [Math]::Round($tmp.Avg / 1024, 2) }
            $tmp = $deep['mem.consumed.average'];      if ($tmp) { $memConsumedGB   = [Math]::Round($tmp.Avg / 1024, 2) }
            $tmp = $deep['mem.granted.average'];       if ($tmp) { $memGrantedGB    = [Math]::Round($tmp.Avg / 1024, 2) }
            $tmp = $deep['mem.shared.average'];        if ($tmp) { $memSharedMB     = [Math]::Round($tmp.Avg, 2) }
            $tmp = $deep['mem.swapped.average'];       if ($tmp) { $memSwappedMB    = [Math]::Round($tmp.Max, 2) }
            $tmp = $deep['mem.swapout.average'];       if ($tmp) { $memSwapoutMB    = [Math]::Round($tmp.Max, 2) }
            $tmp = $deep['mem.vmmemctl.average'];      if ($tmp) { $memBalloonMB    = [Math]::Round($tmp.Max, 2) }
            $tmp = $deep['mem.latency.average'];       if ($tmp) { $memLatencyAvg   = [Math]::Round($tmp.Avg, 2) }
            $tmp = $deep['mem.compressed.average'];    if ($tmp) { $memCompressedMB = [Math]::Round($tmp.Max, 2) }
            $tmp = $deep['disk.maxTotalLatency.latest']; if ($tmp) {
                $diskMaxLatencyMs  = [Math]::Round($tmp.Max, 2)
                $diskMaxLatencyP95 = [Math]::Round($tmp.P95, 2)
            }
            $tmp = $deep['disk.usage.average'];        if ($tmp) { $diskUsageAvg = [Math]::Round($tmp.Avg, 2) }
            $tmp = $deep['disk.commandsAborted.summation']; if ($tmp) { $diskAborted = [int]$tmp.Max }
            $tmp = $deep['virtualDisk.totalReadLatency.average'];  if ($tmp) { $vDiskReadLatencyMs  = [Math]::Round($tmp.Avg, 2) }
            $tmp = $deep['virtualDisk.totalWriteLatency.average']; if ($tmp) { $vDiskWriteLatencyMs = [Math]::Round($tmp.Avg, 2) }
            $tmp = $deep['net.droppedRx.summation'];   if ($tmp) { $netDroppedRx = [int]$tmp.Max }
            $tmp = $deep['net.droppedTx.summation'];   if ($tmp) { $netDroppedTx = [int]$tmp.Max }
            $tmp = $deep['net.usage.average'];         if ($tmp) { $netUsageAvg  = [Math]::Round($tmp.Avg, 2) }
        }

        # Oracle health class (so para VMs Oracle)
        $oracleLevel = $null
        $oracleFindings = $null
        $oracleReco = $null
        if ($isOracle) {
            $oracleData = @{
                NumVcpu            = $numVcpu
                HostThreads        = $hostThreads
                MemoryGB           = [Math]::Round([double]$vm.MemoryGB, 2)
                IsOversized        = $isOversized
                IsMonster          = $isMonster
                IsBursty           = $isBursty
                CrossesNUMABoundary= $crossesNuma
                ThreadsPerNode     = $threadsPerNumaNode
                RamPerNodeGB       = $ramPerNumaNodeGB
                ContentionAvg      = $contentionAvg
                ContentionMax      = $contentionMax
                CostopMax          = $costopMaxPct
                MemLatencyAvg      = $memLatencyAvg
                MemBalloonedMB     = $memBalloonMB
                MemSwappedMB       = $memSwappedMB
                MemCompressedMB    = $memCompressedMB
                DisparityAvg       = $disparityAvg
                DemandAvgPct       = if ($demand -and $vmHostObj -and $vmHostObj.CpuTotalMhz -gt 0 -and $numVcpu -gt 0) {
                    # demand em MHz / (MHz por core * vCPU) * 100
                    $mhzPerCore = $vmHostObj.CpuTotalMhz / [Math]::Max($hostThreads, 1)
                    [Math]::Round(($demand.Avg / ($mhzPerCore * $numVcpu)) * 100, 2)
                } else { 0 }
                DiskLatencyP95     = $diskMaxLatencyP95
                DiskCommandsAborted= $diskAborted
                NetDropped         = ($netDroppedRx + $netDroppedTx)
            }
            $oh = Get-OracleHealthClass -D $oracleData
            $oracleLevel    = $oh.Level
            $oracleFindings = $oh.Findings
            $oracleReco     = $oh.Recommendation
        }

        # Cluster pode nao existir (VM em folder direto)
        $clusterName = ''
        try {
            $cl = $vm | Get-Cluster -ErrorAction SilentlyContinue
            if ($cl) { $clusterName = $cl.Name }
        } catch {}

        $vmResults += [PSCustomObject]@{
            Name              = $vm.Name
            PowerState        = "$($vm.PowerState)"
            Cluster           = $clusterName
            VMHost            = if ($vmHostObj) { $vmHostObj.Name } else { '' }
            NumCpu            = $numVcpu
            CoresPerSocket    = [int]$vm.CoresPerSocket
            MemoryGB          = [Math]::Round([double]$vm.MemoryGB, 2)
            HostThreads       = $hostThreads
            IsMonsterVM       = $isMonster
            IsOversized       = $isOversized
            IsEdge            = $isEdge
            IsOracleVM            = $isOracle
            OracleDetectionMethod = $oracleMethod
            IsPossibleOracle      = $isPossibleOracle
            ThreadsPerNUMANode    = $threadsPerNumaNode
            RamPerNUMANodeGB      = $ramPerNumaNodeGB
            CrossesNUMABoundary   = $crossesNuma
            UsageAvgPct       = if ($usagePct) { $usagePct.Avg } else { 0 }
            UsageMaxPct       = if ($usagePct) { $usagePct.Max } else { 0 }
            UsageMHzAvg       = if ($usageMHz) { $usageMHz.Avg } else { 0 }
            DemandMHzAvg      = if ($demand)   { $demand.Avg }   else { 0 }
            DemandMHzMax      = if ($demand)   { $demand.Max }   else { 0 }
            ReadyAvgPct       = $readyAvgPct
            ReadyMaxPct       = $readyMaxPct
            ReadyP95Pct       = $readyP95Pct
            CostopAvgPct      = $costopAvgPct
            CostopMaxPct      = $costopMaxPct
            MaxLimitedAvgPct  = $mlmtdAvgPct
            MaxLimitedMaxPct  = $mlmtdMaxPct
            SwapWaitAvgPct    = $swapAvgPct
            SwapWaitMaxPct    = $swapMaxPct
            ContentionAvgPct  = [Math]::Round($contentionAvg, 2)
            ContentionMaxPct  = [Math]::Round($contentionMax, 2)
            ContentionP95Pct  = [Math]::Round($contentionP95, 2)
            ContentionBurstDelta = $contentionBurst
            ReadyBurstDelta      = $readyBurst
            CostopBurstDelta     = $costopBurst
            IsBurstyContention   = $isBursty
            DisparityAvg      = $disparityAvg
            DisparityMax      = $disparityMax
            MemoryActiveGB    = $memActiveGB
            MemoryConsumedGB  = $memConsumedGB
            MemoryGrantedGB   = $memGrantedGB
            MemorySharedMB    = $memSharedMB
            MemorySwappedMB   = $memSwappedMB
            MemorySwapoutMB   = $memSwapoutMB
            MemoryBalloonedMB = $memBalloonMB
            MemoryLatencyPct  = $memLatencyAvg
            MemoryCompressedMB= $memCompressedMB
            DiskMaxLatencyMs  = $diskMaxLatencyMs
            DiskMaxLatencyP95Ms = $diskMaxLatencyP95
            DiskUsageAvgKBps  = $diskUsageAvg
            DiskCommandsAborted = $diskAborted
            VirtualDiskReadLatencyMs  = $vDiskReadLatencyMs
            VirtualDiskWriteLatencyMs = $vDiskWriteLatencyMs
            NetDroppedRx      = $netDroppedRx
            NetDroppedTx      = $netDroppedTx
            NetUsageAvgKBps   = $netUsageAvg
            Level             = $cls.Level
            Findings          = $cls.Findings
            OracleHealthLevel    = $oracleLevel
            OracleFindings       = $oracleFindings
            OracleRecommendation = $oracleReco
            VRops_ResourceId          = $null
            VRops_ContentionAvg30d    = $null
            VRops_ContentionMax30d    = $null
            VRops_ContentionTrend     = $null
            VRops_WorkloadAvg         = $null
            VRops_StressAvg           = $null
            VRops_RecommendedVCPU     = $null
            VRops_RecommendedRAMGB    = $null
            VRops_HistoryDataAvailable = $false
            VRops_MemContentionAvg    = $null
            VRops_MemContentionMax    = $null
            VRops_StorageLatencyAvg   = $null
            CrossValidation           = $null
        }
    }
    Write-Progress -Activity "Coletando metricas de VMs" -Completed
}

#endregion

#region Enriquecimento vROps (seletivo) ------------------------------------------

if ($vRopsEnabled -and $vmResults.Count -gt 0) {
    $candidates = @($vmResults | Where-Object {
        $_.Level -ne 'OK' -or $_.IsMonsterVM -or $_.IsEdge -or $_.IsOracleVM
    })
    Write-Host ""
    Write-Host ("Enriquecendo {0} VMs com historico vROps ({1}d)..." -f $candidates.Count, $VRopsHistoryDays) -ForegroundColor Cyan

    $statKeys = @(
        'cpu|capacity_contentionPct',
        'cpu|workload',
        'badge|stress',
        'summary|workload_indicator'
    )
    $oracleExtraKeys = @(
        'mem|capacity_contentionPct',
        'diskspace|usage_average',
        'storage|totalLatency_average',
        'sys|workload'
    )

    $i = 0
    $tot = $candidates.Count
    foreach ($vmRow in $candidates) {
        $i++
        Write-Progress -Activity "Enriquecimento vROps" -Status $vmRow.Name -PercentComplete (($i / [Math]::Max($tot,1)) * 100)
        $res = Get-VRopsResource -Session $vRopsSession -Name $vmRow.Name
        if (-not $res) {
            Write-Verbose "VM '$($vmRow.Name)' nao encontrada no vROps."
            continue
        }
        $vmRow.VRops_ResourceId = $res.identifier
        $keysForThisVm = if ($vmRow.IsOracleVM) { $statKeys + $oracleExtraKeys } else { $statKeys }
        $stats = Get-VRopsStats -Session $vRopsSession -ResourceId $res.identifier -StatKeys $keysForThisVm -HistoryDays $VRopsHistoryDays
        $contStat = $stats['cpu|capacity_contentionPct']
        $wlStat   = $stats['cpu|workload']
        $stStat   = $stats['badge|stress']
        if ($contStat) {
            $vmRow.VRops_ContentionAvg30d     = $contStat.Avg
            $vmRow.VRops_ContentionMax30d     = $contStat.Max
            $delta = [double]$contStat.Latest - [double]$contStat.Avg
            if ([Math]::Abs($delta) -lt 1) {
                $vmRow.VRops_ContentionTrend = 'ESTAVEL'
            } elseif ($delta -gt 0) {
                $vmRow.VRops_ContentionTrend = 'SUBINDO'
            } else {
                $vmRow.VRops_ContentionTrend = 'CAINDO'
            }
            $vmRow.VRops_HistoryDataAvailable = $true
        }
        if ($wlStat) { $vmRow.VRops_WorkloadAvg = $wlStat.Avg }
        if ($stStat) { $vmRow.VRops_StressAvg   = $stStat.Avg }
        if ($vmRow.IsOracleVM) {
            $memCont = $stats['mem|capacity_contentionPct']
            $stLat   = $stats['storage|totalLatency_average']
            if ($memCont) {
                $vmRow.VRops_MemContentionAvg = $memCont.Avg
                $vmRow.VRops_MemContentionMax = $memCont.Max
            }
            if ($stLat) {
                $vmRow.VRops_StorageLatencyAvg = $stLat.Avg
            }
        }
        $props = Get-VRopsProperties -Session $vRopsSession -ResourceId $res.identifier
        $vmRow.VRops_RecommendedVCPU  = $props.RecommendedVCPU
        $vmRow.VRops_RecommendedRAMGB = $props.RecommendedRAMGB
    }
    Write-Progress -Activity "Enriquecimento vROps" -Completed

    # Cross-validation
    foreach ($vmRow in $vmResults) {
        if (-not $vmRow.VRops_HistoryDataAvailable) { continue }
        $current = [double]$vmRow.ContentionAvgPct
        $hist    = [double]$vmRow.VRops_ContentionAvg30d
        $delta   = $current - $hist
        if ([Math]::Abs($delta) -lt 2) {
            $vmRow.CrossValidation = "CRONICO - alinhado com historico ${VRopsHistoryDays}d (delta $([Math]::Round($delta,2)) pp)"
        } elseif ($delta -gt 5) {
            $vmRow.CrossValidation = "INCIDENTE - atual $([Math]::Round($current,2))% acima da media $([Math]::Round($hist,2))%"
        } elseif ($delta -lt -5) {
            $vmRow.CrossValidation = "MELHORANDO - atual $([Math]::Round($current,2))% abaixo da media $([Math]::Round($hist,2))%"
        } else {
            $vmRow.CrossValidation = "OSCILANDO - variacao normal (delta $([Math]::Round($delta,2)) pp)"
        }
    }
}

#endregion

#region Coleta de hosts ----------------------------------------------------------

$hostResults = @()
if (-not $SkipHosts -and $allHosts.Count -gt 0) {
    $idx = 0
    foreach ($h in $allHosts) {
        $idx++
        Write-Progress -Activity "Coletando metricas de hosts" -Status $h.Name -PercentComplete (($idx / $allHosts.Count) * 100)

        $threads      = [int]$h.NumCpu
        $sockets      = if ($h.ExtensionData -and $h.ExtensionData.Hardware -and $h.ExtensionData.Hardware.CpuInfo) {
            [int]$h.ExtensionData.Hardware.CpuInfo.NumCpuPackages
        } else { 0 }
        $coresPerSock = if ($sockets -gt 0) {
            [int]($h.ExtensionData.Hardware.CpuInfo.NumCpuCores / $sockets)
        } else { 0 }
        $cpuModel     = if ($h.ExtensionData -and $h.ExtensionData.Summary.Hardware) {
            $h.ExtensionData.Summary.Hardware.CpuModel
        } else { '' }
        $cpuMhz       = [int]$h.CpuTotalMhz
        $usedMhz      = [int]$h.CpuUsageMhz

        # vCPUs alocados nas VMs ligadas
        $hostVMs       = @(Get-VM -Location $h | Where-Object { $_.PowerState -eq 'PoweredOn' })
        $allocatedVcpu = ($hostVMs | Measure-Object -Property NumCpu -Sum).Sum
        if (-not $allocatedVcpu) { $allocatedVcpu = 0 }
        $monsterCount  = ($hostVMs | Where-Object { $_.NumCpu -ge $threads }).Count
        $overcommit    = if ($threads -gt 0) { [Math]::Round($allocatedVcpu / $threads, 2) } else { 0 }

        # Metricas
        $usagePct = Get-CPUStats -Entity $h -StatName 'cpu.usage.average'    -IntervalInfo $intervalInfo -Hours $Hours
        $usageMHz = Get-CPUStats -Entity $h -StatName 'cpu.usagemhz.average' -IntervalInfo $intervalInfo -Hours $Hours
        $demand   = Get-CPUStats -Entity $h -StatName 'cpu.demand.average'   -IntervalInfo $intervalInfo -Hours $Hours
        $latency  = Get-CPUStats -Entity $h -StatName 'cpu.latency.average'  -IntervalInfo $intervalInfo -Hours $Hours

        # Cluster do host
        $clName = ''
        try {
            $cl = $h | Get-Cluster -ErrorAction SilentlyContinue
            if ($cl) { $clName = $cl.Name }
        } catch {}

        $hostResults += [PSCustomObject]@{
            Name             = $h.Name
            Cluster          = $clName
            ConnectionState  = "$($h.ConnectionState)"
            PowerState       = "$($h.PowerState)"
            CpuModel         = $cpuModel
            Sockets          = $sockets
            CoresPerSocket   = $coresPerSock
            Threads          = $threads
            CpuMhzPerCore    = if ($threads -gt 0) { [int]($cpuMhz / $threads) } else { 0 }
            CpuTotalMhz      = $cpuMhz
            CpuUsageMhz      = $usedMhz
            AllocatedVcpu    = [int]$allocatedVcpu
            OvercommitRatio  = $overcommit
            MonsterVMsHosted = [int]$monsterCount
            UsageAvgPct      = if ($usagePct) { $usagePct.Avg } else { 0 }
            UsageMaxPct      = if ($usagePct) { $usagePct.Max } else { 0 }
            UsageP95Pct      = if ($usagePct) { $usagePct.P95 } else { 0 }
            UsageMHzAvg      = if ($usageMHz) { $usageMHz.Avg } else { 0 }
            DemandMHzAvg     = if ($demand)   { $demand.Avg }   else { 0 }
            DemandMHzMax     = if ($demand)   { $demand.Max }   else { 0 }
            ContentionAvgPct = if ($latency)  { [Math]::Round($latency.Avg, 2) } else { 0 }
            ContentionMaxPct = if ($latency)  { [Math]::Round($latency.Max, 2) } else { 0 }
            ContentionP95Pct = if ($latency)  { [Math]::Round($latency.P95, 2) } else { 0 }
        }
    }
    Write-Progress -Activity "Coletando metricas de hosts" -Completed
}

#endregion

#region Sumario por cluster ------------------------------------------------------

$clusterSummary = @()
if ($vmResults.Count -gt 0) {
    $byCluster = $vmResults | Group-Object -Property Cluster
    foreach ($g in $byCluster) {
        $vms = $g.Group
        $contAvg = ($vms | Measure-Object -Property ContentionAvgPct -Average).Average
        $rdyAvg  = ($vms | Measure-Object -Property ReadyAvgPct -Average).Average
        $clusterSummary += [PSCustomObject]@{
            Cluster        = if ($g.Name) { $g.Name } else { '(sem cluster)' }
            TotalVMs       = $vms.Count
            CriticalVMs    = ($vms | Where-Object { $_.Level -eq 'CRITICO' }).Count
            WarningVMs     = ($vms | Where-Object { $_.Level -eq 'ATENCAO' }).Count
            MonsterVMs     = ($vms | Where-Object { $_.IsMonsterVM }).Count
            OversizedVMs   = ($vms | Where-Object { $_.IsOversized }).Count
            EdgeVMs        = ($vms | Where-Object { $_.IsEdge }).Count
            AvgContention  = [Math]::Round([double]$contAvg, 2)
            AvgReady       = [Math]::Round([double]$rdyAvg, 2)
        }
    }
}

#endregion

#region Saida JSON ---------------------------------------------------------------

$oracleVMs        = @($vmResults | Where-Object { $_.IsOracleVM })
$oracleByTag      = @($oracleVMs | Where-Object { $_.OracleDetectionMethod -eq 'tag' }).Count
$oracleByName     = @($oracleVMs | Where-Object { $_.OracleDetectionMethod -eq 'name_pattern' }).Count
$oracleByNotes    = @($oracleVMs | Where-Object { $_.OracleDetectionMethod -eq 'notes' }).Count
$oracleCritical   = @($oracleVMs | Where-Object { $_.OracleHealthLevel -eq 'CRITICO' }).Count
$possibleOracle   = @($vmResults | Where-Object { $_.IsPossibleOracle }).Count
$burstyVMs        = @($vmResults | Where-Object { $_.IsBurstyContention }).Count
$numaCrossing     = @($vmResults | Where-Object { $_.CrossesNUMABoundary }).Count

$summary = [PSCustomObject]@{
    totalVMs            = $vmResults.Count
    criticalVMs         = ($vmResults | Where-Object { $_.Level -eq 'CRITICO' }).Count
    warningVMs          = ($vmResults | Where-Object { $_.Level -eq 'ATENCAO' }).Count
    monsterVMs          = ($vmResults | Where-Object { $_.IsMonsterVM }).Count
    oversizedVMs        = ($vmResults | Where-Object { $_.IsOversized }).Count
    edgeVMs             = ($vmResults | Where-Object { $_.IsEdge }).Count
    totalHosts          = $hostResults.Count
    vRopsEnriched       = ($vmResults | Where-Object { $_.VRops_HistoryDataAvailable }).Count
    oracleVMs           = $oracleVMs.Count
    oracleVMsByTag      = $oracleByTag
    oracleVMsByName     = $oracleByName
    oracleVMsByNotes    = $oracleByNotes
    oracleCriticalVMs   = $oracleCritical
    possibleOracleVMs   = $possibleOracle
    burstyContentionVMs = $burstyVMs
    numaCrossingVMs     = $numaCrossing
}

$metadata = [PSCustomObject]@{
    runLabel              = $runLabel
    timestamp             = $runTimestamp.ToString('o')
    vCenter               = $VCenter
    vCenterVersion        = $vCenterVersion
    vRopsServer           = $VRopsServer
    vRopsEnriched         = $vRopsEnabled
    vRopsHistoryDays      = $VRopsHistoryDays
    windowHours           = $Hours
    windowStart           = $runTimestamp.AddHours(-$Hours).ToString('o')
    windowEnd             = $runTimestamp.ToString('o')
    granularity           = $intervalInfo.Name
    intervalSecs          = $intervalInfo.IntervalSecs
    filters               = [PSCustomObject]@{
        cluster    = $Cluster
        datacenter = $Datacenter
        vmName     = $VMName
        vmHost     = $VMHost
    }
    edgePattern           = $EdgePattern
    oracleAnalysisEnabled = ($oracleVMs.Count -gt 0 -or $OracleDeepAnalysis.IsPresent)
    oraclePattern         = $OraclePattern
    oracleVMTag           = $OracleVMTag
    oracleDeepAnalysis    = $OracleDeepAnalysis.IsPresent
    burstyThreshold       = $BurstyThreshold
}

$jsonOutput = [PSCustomObject]@{
    schemaVersion = '3.2'
    metadata      = $metadata
    summary       = $summary
    clusters      = @($clusterSummary)
    vms           = @($vmResults)
    hosts         = @($hostResults)
}

$jsonPath = Join-Path $OutputPath "$runLabel.json"
$jsonOutput | ConvertTo-Json -Depth 10 | Set-Content -Path $jsonPath -Encoding UTF8
Write-Host ""
Write-Host "JSON salvo em: $jsonPath" -ForegroundColor Green

#endregion

#region Saida CSV ----------------------------------------------------------------

$vmsCsvPath   = Join-Path $OutputPath "$runLabel-vms.csv"
$hostsCsvPath = Join-Path $OutputPath "$runLabel-hosts.csv"

if ($vmResults.Count -gt 0) {
    $vmResults |
        Sort-Object -Property ContentionAvgPct -Descending |
        Export-Csv -Path $vmsCsvPath -NoTypeInformation -Encoding UTF8
    Write-Host "CSV de VMs salvo em: $vmsCsvPath" -ForegroundColor Green
}
if ($hostResults.Count -gt 0) {
    $hostResults |
        Sort-Object -Property OvercommitRatio -Descending |
        Export-Csv -Path $hostsCsvPath -NoTypeInformation -Encoding UTF8
    Write-Host "CSV de hosts salvo em: $hostsCsvPath" -ForegroundColor Green
}

#endregion

#region Modo comparacao ----------------------------------------------------------

$compareResults = @()
if ($CompareWith) {
    if (-not (Test-Path -LiteralPath $CompareWith)) {
        Write-Error "Arquivo de baseline nao encontrado: $CompareWith"
        exit 1
    }
    Write-Host ""
    Write-Host "Comparando com baseline $CompareWith ..." -ForegroundColor Cyan
    try {
        $baseline = Get-Content -LiteralPath $CompareWith -Raw -Encoding UTF8 | ConvertFrom-Json
    } catch {
        Write-Error "Falha lendo baseline JSON: $($_.Exception.Message)"
        exit 1
    }
    $baselineVms = @{}
    foreach ($bv in $baseline.vms) { $baselineVms[$bv.Name] = $bv }

    foreach ($cur in $vmResults) {
        if (-not $baselineVms.ContainsKey($cur.Name)) { continue }
        $base = $baselineVms[$cur.Name]
        $dCont   = [Math]::Round([double]$cur.ContentionAvgPct - [double]$base.ContentionAvgPct, 2)
        $dReady  = [Math]::Round([double]$cur.ReadyAvgPct      - [double]$base.ReadyAvgPct, 2)
        $dCostop = [Math]::Round([double]$cur.CostopAvgPct     - [double]$base.CostopAvgPct, 2)
        $verdict = 'ESTAVEL'
        if ($dCont -gt 2)      { $verdict = 'PIOROU' }
        elseif ($dCont -lt -2) { $verdict = 'MELHOROU' }
        $compareResults += [PSCustomObject]@{
            Name                 = $cur.Name
            Cluster              = $cur.Cluster
            ContentionAvgBefore  = [double]$base.ContentionAvgPct
            ContentionAvgAfter   = [double]$cur.ContentionAvgPct
            DeltaContention      = $dCont
            ReadyAvgBefore       = [double]$base.ReadyAvgPct
            ReadyAvgAfter        = [double]$cur.ReadyAvgPct
            DeltaReady           = $dReady
            CostopAvgBefore      = [double]$base.CostopAvgPct
            CostopAvgAfter       = [double]$cur.CostopAvgPct
            DeltaCostop          = $dCostop
            LevelBefore          = "$($base.Level)"
            LevelAfter           = $cur.Level
            Verdict              = $verdict
        }
    }
    if ($compareResults.Count -gt 0) {
        $comparePath = Join-Path $OutputPath "$runLabel-compare.csv"
        $compareResults |
            Sort-Object -Property DeltaContention -Descending |
            Export-Csv -Path $comparePath -NoTypeInformation -Encoding UTF8
        Write-Host "CSV de comparacao salvo em: $comparePath" -ForegroundColor Green
    } else {
        Write-Warning "Nenhuma VM em comum entre baseline e execucao atual."
    }
}

#endregion

#region Resumo no console --------------------------------------------------------

Write-Host ""
Write-Host "==================== RESUMO EXECUTIVO ====================" -ForegroundColor White
Write-Host ("Run:        {0}" -f $runLabel)
Write-Host ("vCenter:    {0}  ({1})" -f $VCenter, $vCenterVersion)
if ($vRopsEnabled) {
    Write-Host ("vROps:      {0}  (historico {1}d)" -f $VRopsServer, $VRopsHistoryDays)
}
Write-Host ("Janela:     {0}h ({1}, {2}s)" -f $Hours, $intervalInfo.Name, $intervalInfo.IntervalSecs)
Write-Host ""
Write-Host ("Total VMs:        {0}" -f $summary.totalVMs)
Write-Host ("  Criticas:       {0}" -f $summary.criticalVMs) -ForegroundColor Red
Write-Host ("  Atencao:        {0}" -f $summary.warningVMs)  -ForegroundColor Yellow
Write-Host ("  Monster VMs:    {0}" -f $summary.monsterVMs)  -ForegroundColor Magenta
Write-Host ("  Oversized:      {0}" -f $summary.oversizedVMs)
Write-Host ("  NSX Edges:      {0}" -f $summary.edgeVMs)     -ForegroundColor Cyan
Write-Host ("Total hosts:      {0}" -f $summary.totalHosts)
if ($vRopsEnabled) {
    Write-Host ("vROps enriched:   {0}" -f $summary.vRopsEnriched) -ForegroundColor Cyan
}

if ($clusterSummary.Count -gt 1) {
    Write-Host ""
    Write-Host "----- Sumario por cluster -----" -ForegroundColor White
    $clusterSummary |
        Sort-Object -Property CriticalVMs -Descending |
        Format-Table Cluster, TotalVMs, CriticalVMs, WarningVMs, MonsterVMs, EdgeVMs, AvgContention, AvgReady -AutoSize | Out-Host
}

# Secao Oracle (so se houver VMs Oracle ou possiveis Oracle)
if ($oracleVMs.Count -gt 0 -or $possibleOracle -gt 0) {
    Write-Host ""
    Write-Host "===== Analise Oracle =====" -ForegroundColor Cyan
    Write-Host ("Pattern usado: {0}" -f $OraclePattern)
    if ($OracleVMTag) { Write-Host ("Tags consideradas: {0}" -f ($OracleVMTag -join ', ')) }
    Write-Host ("Oracle VMs identificadas: {0}" -f $oracleVMs.Count) -ForegroundColor Green
    Write-Host ("  Por tag:        {0}" -f $oracleByTag)
    Write-Host ("  Por nome:       {0}" -f $oracleByName)
    Write-Host ("  Por anotacao:   {0}" -f $oracleByNotes)
    Write-Host ("  Em CRITICO:     {0}" -f $oracleCritical) -ForegroundColor Red
    Write-Host ("  Bursty:         {0}" -f ($oracleVMs | Where-Object { $_.IsBurstyContention }).Count) -ForegroundColor Yellow
    Write-Host ("  NUMA mismatch:  {0}" -f ($oracleVMs | Where-Object { $_.CrossesNUMABoundary }).Count) -ForegroundColor Yellow
    Write-Host ("  Memory press.:  {0}" -f ($oracleVMs | Where-Object { $_.MemoryBalloonedMB -gt 0 -or $_.MemorySwappedMB -gt 0 }).Count) -ForegroundColor Yellow
    Write-Host ("Possivelmente Oracle (revisar com cliente): {0}" -f $possibleOracle) -ForegroundColor Yellow

    $oracleProblems = @($oracleVMs | Where-Object { $_.OracleHealthLevel -eq 'CRITICO' -or $_.OracleHealthLevel -eq 'ATENCAO' } |
        Sort-Object -Property @{Expression='OracleHealthLevel'; Descending=$true}, @{Expression='ContentionAvgPct'; Descending=$true})
    if ($oracleProblems.Count -gt 0) {
        Write-Host ""
        Write-Host ("----- Top Oracle VMs com problemas ({0}) -----" -f $oracleProblems.Count) -ForegroundColor Red
        $oracleProblems |
            Select-Object -First $TopN |
            Format-Table Name, NumCpu, MemoryGB, OracleDetectionMethod, ContentionAvgPct, ContentionMaxPct, CostopMaxPct, MemoryLatencyPct, MemoryBalloonedMB, OracleHealthLevel, OracleRecommendation -AutoSize -Wrap | Out-Host
    }

    if ($possibleOracle -gt 0) {
        Write-Host ""
        Write-Host ("----- Possivelmente Oracle (nao identificadas) -----") -ForegroundColor Yellow
        $vmResults | Where-Object { $_.IsPossibleOracle } |
            Sort-Object -Property NumCpu -Descending |
            Select-Object -First $TopN |
            Format-Table Name, Cluster, NumCpu, MemoryGB, ContentionAvgPct, IsBurstyContention -AutoSize | Out-Host
    }
}

# Secao bursty (VMs - Oracle ou nao - com IsBurstyContention)
$burstyAll = @($vmResults | Where-Object { $_.IsBurstyContention } | Sort-Object -Property ContentionBurstDelta -Descending)
if ($burstyAll.Count -gt 0) {
    Write-Host ""
    Write-Host ("----- VMs com padrao bursty (Max-Avg > {0}pp) ({1}) -----" -f $BurstyThreshold, $burstyAll.Count) -ForegroundColor Yellow
    $burstyAll |
        Select-Object -First $TopN |
        Format-Table Name, Cluster, NumCpu, ContentionAvgPct, ContentionMaxPct, ContentionBurstDelta, IsOracleVM, Level -AutoSize | Out-Host
}

$criticalVMs = @($vmResults | Where-Object { $_.Level -eq 'CRITICO' } | Sort-Object -Property ContentionAvgPct -Descending)
if ($criticalVMs.Count -gt 0) {
    Write-Host ""
    Write-Host ("----- Top {0} VMs criticas -----" -f [Math]::Min($TopN, $criticalVMs.Count)) -ForegroundColor Red
    $criticalVMs |
        Select-Object -First $TopN |
        Format-Table Name, Cluster, NumCpu, ContentionAvgPct, ReadyAvgPct, CostopAvgPct, MaxLimitedAvgPct, IsMonsterVM, IsEdge, Findings -AutoSize -Wrap | Out-Host
}

if ($hostResults.Count -gt 0) {
    Write-Host ""
    Write-Host "----- Hosts ordenados por overcommit -----" -ForegroundColor White
    $hostResults |
        Sort-Object -Property OvercommitRatio -Descending |
        Select-Object -First $TopN |
        Format-Table Name, Cluster, Threads, AllocatedVcpu, OvercommitRatio, MonsterVMsHosted, UsageAvgPct, ContentionAvgPct -AutoSize | Out-Host
}

if ($compareResults.Count -gt 0) {
    $worsened = @($compareResults | Where-Object { $_.Verdict -eq 'PIOROU' } | Sort-Object -Property DeltaContention -Descending)
    if ($worsened.Count -gt 0) {
        Write-Host ""
        Write-Host ("----- VMs que pioraram vs baseline ({0}) -----" -f $worsened.Count) -ForegroundColor Red
        $worsened |
            Select-Object -First $TopN |
            Format-Table Name, Cluster, ContentionAvgBefore, ContentionAvgAfter, DeltaContention, DeltaReady, LevelBefore, LevelAfter -AutoSize | Out-Host
    }
}

#endregion

#region HTML executivo -----------------------------------------------------------

$css = @"
<style>
:root {
  --bg: #fafafa;
  --fg: #1a1a1a;
  --muted: #555;
  --border: #e5e5e5;
  --card-bg: #ffffff;
  --critical: #a32d2d;
  --critical-bg: #fdecec;
  --warning: #854f0b;
  --warning-bg: #fdf3df;
  --monster: #993556;
  --monster-bg: #fde7ee;
  --edge: #0f6e56;
  --edge-bg: #def4ec;
  --accent: #2563eb;
}
@media (prefers-color-scheme: dark) {
  :root {
    --bg: #161616;
    --fg: #ececec;
    --muted: #a8a8a8;
    --border: #2e2e2e;
    --card-bg: #1f1f1f;
    --critical: #ff7a7a;
    --critical-bg: #3a1717;
    --warning: #f0c66f;
    --warning-bg: #3a2c0d;
    --monster: #ff8fa9;
    --monster-bg: #3a1822;
    --edge: #5fd6b8;
    --edge-bg: #0f3329;
    --accent: #79a7ff;
  }
}
* { box-sizing: border-box; }
body {
  margin: 0;
  font-family: -apple-system, system-ui, "Segoe UI", Roboto, sans-serif;
  background: var(--bg);
  color: var(--fg);
  line-height: 1.55;
  font-size: 14px;
}
.wrap { max-width: 1200px; margin: 0 auto; padding: 32px 24px 80px; }
h1 { font-weight: 500; font-size: 28px; margin: 0 0 4px; }
h2 { font-weight: 500; font-size: 20px; margin: 36px 0 12px; padding-bottom: 6px; border-bottom: 1px solid var(--border); }
h3 { font-weight: 500; font-size: 16px; margin: 24px 0 8px; }
.meta { color: var(--muted); font-size: 13px; margin-bottom: 20px; }
.badge {
  display: inline-block;
  padding: 2px 10px;
  border-radius: 999px;
  font-size: 12px;
  font-weight: 500;
  margin-left: 8px;
  vertical-align: middle;
  background: var(--edge-bg);
  color: var(--edge);
  border: 1px solid var(--edge);
}
.kpis { display: grid; grid-template-columns: repeat(auto-fit, minmax(140px, 1fr)); gap: 12px; margin: 16px 0 32px; }
.kpi {
  background: var(--card-bg);
  border: 1px solid var(--border);
  border-radius: 8px;
  padding: 14px 16px;
}
.kpi .lbl { color: var(--muted); font-size: 12px; text-transform: uppercase; letter-spacing: 0.04em; }
.kpi .val { font-size: 26px; font-weight: 500; margin-top: 4px; }
.kpi.crit .val { color: var(--critical); }
.kpi.warn .val { color: var(--warning); }
.kpi.mon  .val { color: var(--monster); }
.kpi.edge .val { color: var(--edge); }
table {
  width: 100%;
  border-collapse: collapse;
  margin: 8px 0 20px;
  background: var(--card-bg);
  border: 1px solid var(--border);
  border-radius: 8px;
  overflow: hidden;
  font-size: 13px;
}
th, td { padding: 10px 12px; text-align: left; border-bottom: 1px solid var(--border); }
th { background: var(--card-bg); color: var(--muted); font-weight: 500; text-transform: uppercase; font-size: 11px; letter-spacing: 0.04em; }
tr:last-child td { border-bottom: none; }
tr:hover td { background: rgba(127,127,127,0.05); }
.tag { display: inline-block; padding: 2px 8px; border-radius: 4px; font-size: 11px; font-weight: 500; }
.tag.crit { background: var(--critical-bg); color: var(--critical); }
.tag.warn { background: var(--warning-bg); color: var(--warning); }
.tag.ok   { background: var(--edge-bg);    color: var(--edge); }
.tag.mon  { background: var(--monster-bg); color: var(--monster); }
.tag.edge { background: var(--edge-bg);    color: var(--edge); }
.note {
  border-left: 3px solid var(--accent);
  padding: 10px 14px;
  background: var(--card-bg);
  border-radius: 0 8px 8px 0;
  margin: 12px 0;
  color: var(--muted);
  font-size: 13px;
}
.ruler { display: grid; grid-template-columns: 1fr 1fr; gap: 16px; margin: 12px 0 24px; }
.ruler .col {
  background: var(--card-bg);
  border: 1px solid var(--border);
  border-radius: 8px;
  padding: 14px 16px;
}
.ruler ul { margin: 6px 0 0; padding-left: 18px; }
.ruler li { margin: 4px 0; }
.foot { color: var(--muted); font-size: 12px; margin-top: 48px; padding-top: 16px; border-top: 1px solid var(--border); }
</style>
"@

function ConvertTo-VmHtmlRow {
    param($vm)
    $levelTag = switch ($vm.Level) {
        'CRITICO' { '<span class="tag crit">CRITICO</span>' }
        'ATENCAO' { '<span class="tag warn">ATENCAO</span>' }
        default   { '<span class="tag ok">OK</span>' }
    }
    $tags = @()
    if ($vm.IsMonsterVM) { $tags += '<span class="tag mon">monster</span>' }
    if ($vm.IsEdge)      { $tags += '<span class="tag edge">edge</span>' }
    $tagsHtml = ($tags -join ' ')
    [PSCustomObject]@{
        Name             = $vm.Name
        Cluster          = $vm.Cluster
        vCPU             = $vm.NumCpu
        Level            = $levelTag
        Tags             = $tagsHtml
        Contention       = "$($vm.ContentionAvgPct)% / $($vm.ContentionMaxPct)%"
        Ready            = "$($vm.ReadyAvgPct)% / $($vm.ReadyMaxPct)%"
        Costop           = "$($vm.CostopAvgPct)%"
        MLMTD            = "$($vm.MaxLimitedAvgPct)%"
        Findings         = $vm.Findings
    }
}

$vRopsBadge = if ($vRopsEnabled) { '<span class="badge">vROps enriched</span>' } else { '' }

# KPIs HTML
$kpiHtml = @"
<div class="kpis">
  <div class="kpi"><div class="lbl">Total VMs</div><div class="val">$($summary.totalVMs)</div></div>
  <div class="kpi crit"><div class="lbl">Criticas</div><div class="val">$($summary.criticalVMs)</div></div>
  <div class="kpi warn"><div class="lbl">Atencao</div><div class="val">$($summary.warningVMs)</div></div>
  <div class="kpi mon"><div class="lbl">Monster VMs</div><div class="val">$($summary.monsterVMs)</div></div>
  <div class="kpi edge"><div class="lbl">NSX Edges</div><div class="val">$($summary.edgeVMs)</div></div>
  <div class="kpi"><div class="lbl">Hosts</div><div class="val">$($summary.totalHosts)</div></div>
"@
if ($vRopsEnabled) {
    $kpiHtml += @"
  <div class="kpi edge"><div class="lbl">vROps Enriched</div><div class="val">$($summary.vRopsEnriched)</div></div>
"@
}
if ($summary.oracleVMs -gt 0) {
    $kpiHtml += @"
  <div class="kpi"><div class="lbl">Oracle VMs</div><div class="val">$($summary.oracleVMs)</div></div>
  <div class="kpi crit"><div class="lbl">Oracle Criticas</div><div class="val">$($summary.oracleCriticalVMs)</div></div>
"@
}
if ($summary.burstyContentionVMs -gt 0) {
    $kpiHtml += @"
  <div class="kpi warn"><div class="lbl">Bursty</div><div class="val">$($summary.burstyContentionVMs)</div></div>
"@
}
if ($summary.possibleOracleVMs -gt 0) {
    $kpiHtml += @"
  <div class="kpi warn"><div class="lbl">Possivel Oracle</div><div class="val">$($summary.possibleOracleVMs)</div></div>
"@
}
$kpiHtml += "</div>"

# Cluster summary
$clusterHtml = ''
if ($clusterSummary.Count -gt 0) {
    $clusterHtml = "<h2>Sumario por cluster</h2>" + (
        $clusterSummary | Sort-Object -Property CriticalVMs -Descending |
            ConvertTo-Html -Fragment -Property Cluster, TotalVMs, CriticalVMs, WarningVMs, MonsterVMs, OversizedVMs, EdgeVMs, AvgContention, AvgReady
    )
}

# Critical VMs
$critFragment = ''
$critList = @($vmResults | Where-Object { $_.Level -eq 'CRITICO' } | Sort-Object -Property ContentionAvgPct -Descending)
if ($critList.Count -gt 0) {
    $rows = $critList | ForEach-Object { ConvertTo-VmHtmlRow $_ }
    $critFragment = "<h2>VMs criticas ($($critList.Count))</h2>" + (
        $rows | ConvertTo-Html -Fragment -Property Name, Cluster, vCPU, Level, Tags, Contention, Ready, Costop, MLMTD, Findings
    )
    # Decode tags HTML que ConvertTo-Html escapa
    $critFragment = $critFragment -replace '&lt;', '<' -replace '&gt;', '>' -replace '&quot;', '"'
}

# Warning VMs
$warnFragment = ''
$warnList = @($vmResults | Where-Object { $_.Level -eq 'ATENCAO' } | Sort-Object -Property ContentionAvgPct -Descending)
if ($warnList.Count -gt 0) {
    $rows = $warnList | ForEach-Object { ConvertTo-VmHtmlRow $_ }
    $warnFragment = "<h2>VMs em atencao ($($warnList.Count))</h2>" + (
        $rows | ConvertTo-Html -Fragment -Property Name, Cluster, vCPU, Level, Tags, Contention, Ready, Costop, MLMTD, Findings
    )
    $warnFragment = $warnFragment -replace '&lt;', '<' -replace '&gt;', '>' -replace '&quot;', '"'
}

# Monster VMs
$monsterFragment = ''
$monsterList = @($vmResults | Where-Object { $_.IsMonsterVM } | Sort-Object -Property NumCpu -Descending)
if ($monsterList.Count -gt 0) {
    $monsterFragment = "<h2>Monster VMs ($($monsterList.Count))</h2>" +
        '<div class="note">Monster VMs tem vCPU >= total de threads do host. Gang scheduling fica praticamente impossivel: o scheduler precisa esperar todas as threads ficarem livres simultaneamente.</div>' + (
        $monsterList | ConvertTo-Html -Fragment -Property Name, Cluster, VMHost, NumCpu, HostThreads, ContentionAvgPct, ReadyAvgPct, CostopAvgPct, Level
    )
}

# Edge VMs
$edgeFragment = ''
$edgeList = @($vmResults | Where-Object { $_.IsEdge } | Sort-Object -Property ContentionAvgPct -Descending)
if ($edgeList.Count -gt 0) {
    $edgeFragment = "<h2>NSX Edges ($($edgeList.Count))</h2>" +
        '<div class="note">NSX Edges usam DPDK em poll mode 100% e nao toleram contention. Regua mais conservadora: 5% para CRITICO e 2% para ATENCAO.</div>' + (
        $edgeList | ConvertTo-Html -Fragment -Property Name, Cluster, NumCpu, ContentionAvgPct, ReadyAvgPct, MaxLimitedAvgPct, Level, Findings
    )
}

# Hosts
$hostsFragment = ''
if ($hostResults.Count -gt 0) {
    $hostsFragment = "<h2>Hosts e overcommit</h2>" + (
        $hostResults | Sort-Object -Property OvercommitRatio -Descending |
            ConvertTo-Html -Fragment -Property Name, Cluster, CpuModel, Sockets, CoresPerSocket, Threads, AllocatedVcpu, OvercommitRatio, MonsterVMsHosted, UsageAvgPct, ContentionAvgPct
    )
}

# vROps section
$vropsFragment = ''
if ($vRopsEnabled) {
    $enriched = @($vmResults | Where-Object { $_.VRops_HistoryDataAvailable } | Sort-Object -Property VRops_ContentionAvg30d -Descending)
    if ($enriched.Count -gt 0) {
        $vropsFragment = "<h2>Enriquecimento vROps - cross-validation</h2>" +
            "<div class='note'>Comparacao entre janela atual ($($Hours)h) e media historica ($($VRopsHistoryDays)d) extraida do vROps. Permite distinguir incidentes pontuais de problemas cronicos.</div>" + (
            $enriched | ConvertTo-Html -Fragment -Property Name, Cluster, ContentionAvgPct, VRops_ContentionAvg30d, VRops_ContentionMax30d, VRops_ContentionTrend, VRops_StressAvg, VRops_RecommendedVCPU, VRops_RecommendedRAMGB, CrossValidation
        )
    }
}

# Compare
$compareFragment = ''
if ($compareResults.Count -gt 0) {
    $compareFragment = "<h2>Comparacao com baseline</h2>" +
        "<div class='note'>Baseline: $CompareWith</div>" + (
        $compareResults | Sort-Object -Property DeltaContention -Descending |
            ConvertTo-Html -Fragment -Property Name, Cluster, ContentionAvgBefore, ContentionAvgAfter, DeltaContention, DeltaReady, DeltaCostop, LevelBefore, LevelAfter, Verdict
    )
}

# Oracle: secao consolidada
$oracleFragment = ''
$oracleConsolidatedRows = @($oracleVMs | Sort-Object -Property @{Expression='OracleHealthLevel'; Descending=$true}, @{Expression='ContentionAvgPct'; Descending=$true})
if ($oracleConsolidatedRows.Count -gt 0) {
    $oracleHeader = "<h2>Analise Oracle Database</h2>"
    $oracleNote = "<p>VMs identificadas como Oracle por tag, regex no nome, ou anotacao. Pattern usado: <code>$([System.Web.HttpUtility]::HtmlEncode($OraclePattern))</code>. Aplicada regua especializada considerando bursty workload, NUMA, memoria e padrao de paralelismo.</p>"
    # HtmlEncode pode nao estar disponivel em PS5.1 sem assembly; usar fallback simples
    if (-not ('System.Web.HttpUtility' -as [type])) {
        $oraclePatternEsc = ($OraclePattern -replace '&','&amp;' -replace '<','&lt;' -replace '>','&gt;')
        $oracleNote = "<p>VMs identificadas como Oracle por tag, regex no nome, ou anotacao. Pattern usado: <code>$oraclePatternEsc</code>. Aplicada regua especializada considerando bursty workload, NUMA, memoria e padrao de paralelismo.</p>"
    }
    $oracleConsolidatedTable = $oracleConsolidatedRows | ConvertTo-Html -Fragment -Property Name, Cluster, NumCpu, MemoryGB, OracleDetectionMethod, ContentionAvgPct, ContentionMaxPct, CostopMaxPct, OracleHealthLevel, OracleFindings, OracleRecommendation
    $oracleFragment = $oracleHeader + $oracleNote + "<h3>Oracle VMs - estado consolidado</h3>" + $oracleConsolidatedTable

    # Deep analysis
    $oracleDeepRows = @($oracleVMs | Where-Object { $_.MemoryConsumedGB -gt 0 -or $_.DiskMaxLatencyMs -gt 0 })
    if ($oracleDeepRows.Count -gt 0) {
        $deepTable = $oracleDeepRows |
            Sort-Object -Property OracleHealthLevel -Descending |
            ConvertTo-Html -Fragment -Property Name, MemoryActiveGB, MemoryConsumedGB, MemoryBalloonedMB, MemorySwappedMB, MemoryCompressedMB, MemoryLatencyPct, DiskMaxLatencyP95Ms, VirtualDiskReadLatencyMs, VirtualDiskWriteLatencyMs, NetDroppedRx, NetDroppedTx, DisparityAvg, DisparityMax
        $oracleFragment += "<h3>Memoria, storage e paralelismo (Oracle deep analysis)</h3>" + $deepTable
    }
}

# Bursty (todas as VMs, Oracle ou nao)
$burstyFragment = ''
$burstyRows = @($vmResults | Where-Object { $_.IsBurstyContention } | Sort-Object -Property ContentionBurstDelta -Descending)
if ($burstyRows.Count -gt 0) {
    $burstyFragment = "<h2>Padrao bursty detectado ($($burstyRows.Count))</h2>" +
        "<div class='note'>VMs onde Max - Avg de Contention excede $($BurstyThreshold)pp. Picos episodicos que escapam de reguas baseadas em medias.</div>" + (
        $burstyRows | ConvertTo-Html -Fragment -Property Name, Cluster, NumCpu, IsOracleVM, ContentionAvgPct, ContentionMaxPct, ContentionBurstDelta, ReadyBurstDelta, CostopBurstDelta, Level
    )
}

# Possivelmente Oracle
$possibleOracleFragment = ''
$possibleOracleRows = @($vmResults | Where-Object { $_.IsPossibleOracle } | Sort-Object -Property NumCpu -Descending)
if ($possibleOracleRows.Count -gt 0) {
    $possibleOracleFragment = "<h2>Possivelmente Oracle ($($possibleOracleRows.Count)) - revisar com cliente</h2>" +
        "<div class='note'>VMs com perfil de banco grande (vCPU>=8, RAM>=16GB, nome com db/database/rdbms) mas nao identificadas pelo pattern atual. Podem ser Oracle nao-padronizadas ou outros bancos. Validar com cliente e ajustar -OraclePattern se necessario.</div>" + (
        $possibleOracleRows | ConvertTo-Html -Fragment -Property Name, Cluster, NumCpu, MemoryGB, ContentionAvgPct, IsBurstyContention, Level
    )
}

# Regua
$reguaFragment = @"
<h2>Regua diagnostica aplicada</h2>
<div class="ruler">
  <div class="col">
    <h3>VMs gerais</h3>
    <ul>
      <li><strong>CRITICO:</strong> Contention &gt; 10% ou Ready &gt; 10%/vCPU</li>
      <li><strong>ATENCAO:</strong> Contention 5-10% ou Ready 5-10%/vCPU</li>
      <li><strong>OK:</strong> abaixo de 5% e sem findings adicionais</li>
    </ul>
  </div>
  <div class="col">
    <h3>NSX Edges (regex: <code>$EdgePattern</code>)</h3>
    <ul>
      <li><strong>CRITICO:</strong> Contention &gt; 5% ou Ready &gt; 5%/vCPU</li>
      <li><strong>ATENCAO:</strong> Contention 2-5% ou Ready 2-5%/vCPU</li>
      <li>Limites menores porque DPDK nao tolera contention.</li>
    </ul>
  </div>
</div>
<div class="note">Findings adicionais (acumulam em qualquer VM): Co-stop &gt; 3% sinaliza VM SMP grande demais; %MLMTD &gt; 0 indica CPU Limit ativo; SwapWait &gt; 0 indica pressao de memoria; monster VM = vCPU &gt;= threads do host (gang scheduling impossivel); oversized = vCPU &gt; threads.</div>
"@

$filtersBits = @()
if ($Cluster)    { $filtersBits += "cluster=$($Cluster -join ',')" }
if ($Datacenter) { $filtersBits += "datacenter=$($Datacenter -join ',')" }
if ($VMName)     { $filtersBits += "vmName=$($VMName -join ',')" }
if ($VMHost)     { $filtersBits += "vmHost=$($VMHost -join ',')" }
$filtersText = if ($filtersBits.Count -gt 0) { $filtersBits -join '  |  ' } else { 'sem filtros' }

$html = @"
<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>CPU Contention - $runLabel</title>
$css
</head>
<body>
<div class="wrap">
  <h1>Analise de contencao de CPU - VMware vSphere $vRopsBadge</h1>
  <div class="meta">
    <strong>Run:</strong> $runLabel<br>
    <strong>vCenter:</strong> $VCenter ($vCenterVersion)<br>
    <strong>Janela:</strong> $($Hours)h ($($intervalInfo.Name), $($intervalInfo.IntervalSecs)s)<br>
    <strong>Filtros:</strong> $filtersText<br>
    <strong>Gerado em:</strong> $($runTimestamp.ToString('yyyy-MM-dd HH:mm:ss'))
  </div>

  <h2>Visao geral</h2>
  $kpiHtml

  $clusterHtml
  $critFragment
  $warnFragment
  $monsterFragment
  $edgeFragment
  $oracleFragment
  $burstyFragment
  $possibleOracleFragment
  $hostsFragment
  $vropsFragment
  $compareFragment
  $reguaFragment

  <div class="foot">
    Gerado por Get-VMwareCPUContention.ps1 - schemaVersion 3.2 - operacao read-only
  </div>
</div>
</body>
</html>
"@

$htmlPath = Join-Path $OutputPath "$runLabel.html"
$html | Set-Content -Path $htmlPath -Encoding UTF8
Write-Host "HTML salvo em: $htmlPath" -ForegroundColor Green

#endregion

#region Cleanup ------------------------------------------------------------------

if ($vRopsEnabled -and $vRopsSession) {
    Disconnect-VRops -Session $vRopsSession
}
if ($vcConnection) {
    Disconnect-VIServer -Server $vcConnection -Confirm:$false -ErrorAction SilentlyContinue | Out-Null
}

Write-Host ""
Write-Host "Concluido." -ForegroundColor Green

#endregion
