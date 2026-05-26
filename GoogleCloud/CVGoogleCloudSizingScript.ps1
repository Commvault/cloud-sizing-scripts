<#
.SYNOPSIS
    GCP Cloud Sizing Script - API-only inventory for local machine execution.

.DESCRIPTION
    - Inventories GCP Compute, Storage, Filestore, Database, and GKE resources using
      Google REST APIs and Cloud Monitoring APIs.
    - Designed for local machine runs and API-based authentication.
    - Does not require gcloud, gsutil, or kubectl for core inventory paths.
    - Supports output streaming, worker-pool scaling, and resilient CSV writes.

.PARAMETER Types
    Optional. Restrict inventory to specific resource types.
    Valid values: VM, Storage, Fileshare, DB, GKE
    Accepts comma-separated or array input.

.PARAMETER Projects
    Optional. Target specific projects by ID/name. If omitted, script discovers
    accessible ACTIVE projects via Cloud Resource Manager API.

.PARAMETER ProjectsFile
    Optional. Text file containing projects to include (one project per line).
    Lines beginning with # are ignored.

.PARAMETER SkipProjects
    Optional. Projects to skip even if discovered or explicitly included.
    Accepts comma-separated or array input.

.PARAMETER SkipProjectsFile
    Optional. Text file containing projects to skip (one project per line).
    Lines beginning with # are ignored.

.PARAMETER ServiceAccount
    Optional. Path to service account JSON key file. Preferred for unattended local runs.

.PARAMETER AccessToken
    Optional. Pre-provided OAuth bearer token. If omitted, token is acquired from
    service account or metadata endpoint.

.PARAMETER OutputDir
    Optional. Output working directory. If omitted, script creates timestamped folder.

.PARAMETER SkipZip
    Optional. Do not create final ZIP archive.

.PARAMETER KeepOutputDir
    Optional. Keep working output directory after completion.

.OUTPUTS
    Creates timestamped inventory files under output directory, including:
    - gcp_inventory_summary_*.csv
    - gcp_inventory_errors_*.csv
    - resource-specific CSV files for selected workloads
    - gcp_sizing_script_output_*.log
    Optionally creates gcp_sizing_api_*.zip unless -SkipZip is used.

.NOTES
    Recommended execution target: local machine with network access to Google APIs.
    Required: PowerShell and GCP IAM permissions for listed resources.

.EXAMPLE
    .\CVGoogleCloudSizingScript_api.ps1 -ServiceAccount "C:\keys\gcp-sa.json" -Projects "proj-a,proj-b" -Types VM,Storage -SkipZip -KeepOutputDir

.EXAMPLE
    .\CVGoogleCloudSizingScript_api.ps1 -ServiceAccount "C:\keys\gcp-sa.json" -ProjectsFile .\projects.txt -SkipProjectsFile .\skip-projects.txt -Types DB,Fileshare
#>

param(
    [string[]]$Types,
    [string[]]$Projects,
    [string]$ProjectsFile,
    [string[]]$SkipProjects,
    [string]$SkipProjectsFile,
    [string]$ServiceAccount,
    [string]$AccessToken,
    [string]$OutputDir,
    [switch]$LargeScaleMode,
    [switch]$EnableProjectWorkerPool,
    [switch]$DisableProjectWorkerPool,
    [int]$ProjectWorkers = 0,
    [int]$MaxThreads = 6,
    [int]$ApiTimeoutSec = 120,
    [int]$MetricLookbackHours = 24,
    [int]$CsvFlushThreshold = 500,
    [switch]$SkipZip,
    [switch]$KeepOutputDir
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# -------------------------
# Normalization and defaults
# -------------------------
if ($ProjectsFile) {
    if (-not (Test-Path $ProjectsFile)) { throw "Projects file not found: $ProjectsFile" }
    $fromFile = @(Get-Content -Path $ProjectsFile -ErrorAction Stop | ForEach-Object { $_.Trim() } | Where-Object { $_ -and -not $_.StartsWith('#') })
    if ($fromFile.Count -gt 0) { $Projects = $fromFile }
}

if ($SkipProjectsFile) {
    if (-not (Test-Path $SkipProjectsFile)) { throw "Skip projects file not found: $SkipProjectsFile" }
    $skipFromFile = @(Get-Content -Path $SkipProjectsFile -ErrorAction Stop | ForEach-Object { $_.Trim() } | Where-Object { $_ -and -not $_.StartsWith('#') })
    if ($skipFromFile.Count -gt 0) {
        if ($SkipProjects -and $SkipProjects.Count -gt 0) {
            $SkipProjects = @($SkipProjects + $skipFromFile)
        } else {
            $SkipProjects = $skipFromFile
        }
    }
}

if ($Projects -and $Projects.Count -eq 1 -and $Projects[0] -match ',') {
    $Projects = $Projects[0].Split(',') | ForEach-Object { $_.Trim() } | Where-Object { $_ }
}
if ($Types -and $Types.Count -eq 1 -and $Types[0] -match ',') {
    $Types = $Types[0].Split(',') | ForEach-Object { $_.Trim() } | Where-Object { $_ }
}
if ($SkipProjects -and $SkipProjects.Count -eq 1 -and $SkipProjects[0] -match ',') {
    $SkipProjects = $SkipProjects[0].Split(',') | ForEach-Object { $_.Trim() } | Where-Object { $_ }
}
if ($SkipProjects -and $SkipProjects.Count -gt 0) {
    $SkipProjects = @($SkipProjects | ForEach-Object { $_.Trim() } | Where-Object { $_ } | Select-Object -Unique)
}

$maxThreadsProvided = $PSBoundParameters.ContainsKey('MaxThreads')
$hostCpuCores = [Math]::Max(1,[Environment]::ProcessorCount)
$hostMemoryGb = 0
try {
    $totalMemoryBytes = [double](Get-CimInstance -ClassName Win32_ComputerSystem -ErrorAction Stop | Select-Object -ExpandProperty TotalPhysicalMemory)
    if ($totalMemoryBytes -gt 0) {
        $hostMemoryGb = [int][Math]::Round(($totalMemoryBytes / 1GB),0,[System.MidpointRounding]::AwayFromZero)
    }
} catch {
    try {
        $gcInfo = [System.GC]::GetGCMemoryInfo()
        $availBytes = [double]$gcInfo.TotalAvailableMemoryBytes
        if ($availBytes -gt 0) {
            $hostMemoryGb = [int][Math]::Round(($availBytes / 1GB),0,[System.MidpointRounding]::AwayFromZero)
        }
    } catch {
        $hostMemoryGb = 0
    }
}
if ($hostMemoryGb -lt 1) { $hostMemoryGb = 8 }

# Resource-aware limits: keep inventory stable on laptops while still allowing useful concurrency.
$cpuBasedLimit = [Math]::Max(1,[int][Math]::Floor($hostCpuCores * 0.5))
$memBasedLimit = [Math]::Max(1,[int][Math]::Floor(($hostMemoryGb - 4) / 3))
$recommendedMaxThreads = [Math]::Max(1,[Math]::Min($cpuBasedLimit,$memBasedLimit))
$resourceHardCap = [Math]::Max(2,[Math]::Min(12,$hostCpuCores))
if ($recommendedMaxThreads -gt $resourceHardCap) { $recommendedMaxThreads = $resourceHardCap }

if ($LargeScaleMode) {
    if (-not $maxThreadsProvided) { $MaxThreads = $recommendedMaxThreads }
    if (-not $PSBoundParameters.ContainsKey('MetricLookbackHours')) { $MetricLookbackHours = 48 }
} else {
    if (-not $maxThreadsProvided) { $MaxThreads = $recommendedMaxThreads }
}

if ($MaxThreads -lt 1) { $MaxThreads = 1 }
$script:MaxThreadsClamped = $false
if ($MaxThreads -gt $resourceHardCap) {
    $MaxThreads = $resourceHardCap
    $script:MaxThreadsClamped = $true
}
if ($MetricLookbackHours -lt 1) { $MetricLookbackHours = 1 }
if ($CsvFlushThreshold -lt 1) { $CsvFlushThreshold = 1 }
if ($CsvFlushThreshold -gt 20000) { $CsvFlushThreshold = 20000 }
if ($ProjectWorkers -lt 0) { $ProjectWorkers = 0 }
if ($ProjectWorkers -gt 4) { $ProjectWorkers = 4 }

$script:HostCpuCores = $hostCpuCores
$script:HostMemoryGb = $hostMemoryGb
$script:RecommendedMaxThreads = $recommendedMaxThreads
$script:ResourceHardCap = $resourceHardCap
$script:ParallelCapability = if ($PSVersionTable.PSVersion.Major -ge 7) { 'PS7ParallelAvailable' } else { 'PS5SequentialOnly' }

$script:CsvFlushThreshold = $CsvFlushThreshold
$script:CsvFlushedRowsByFile = @{}
$script:ApiAdaptiveDelayMs = 0
$script:ApiAdaptiveDelayMaxMs = 5000
$script:ApiLastCallUtc = [datetime]::MinValue

$resourceTypeMap = @{
    'VM' = 'VM'
    'STORAGE' = 'STORAGE'
    'FILESHARE' = 'FILESHARE'
    'DB' = 'DB'
    'GKE' = 'GKE'
}
if ($Types) {
    $selected = @{}
    foreach ($t in $Types) {
        $k = $t.Trim().ToUpperInvariant()
        if ($resourceTypeMap.ContainsKey($k)) { $selected[$k] = $true }
    }
    if ($selected.Count -eq 0) { throw 'No valid resource types selected.' }
} else {
    $selected = @{}
    foreach ($k in $resourceTypeMap.Keys) { $selected[$k] = $true }
}

# -------------------------
# Output + logger
# -------------------------
$dateStr = (Get-Date).ToString('yyyy-MM-dd_HHmmss')
$outDir = if ($OutputDir) { $OutputDir } else { Join-Path -Path $PWD -ChildPath ("gcp-inv-api-" + $dateStr) }
New-Item -Path $outDir -ItemType Directory -Force | Out-Null

$script:LogFile = Join-Path $outDir ("gcp_sizing_script_output_" + $dateStr + ".log")
$script:LogLock = New-Object object

function Write-Log {
    param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet('INFO','WARN','ERROR','DEBUG')][string]$Level = 'INFO'
    )
    $ts = (Get-Date).ToString('s')
    $line = "[$ts] [$Level] $Message"
    [System.Threading.Monitor]::Enter($script:LogLock)
    try {
        $canWriteLogFile = $false
        if ($script:LogFile) {
            try {
                $logDir = Split-Path -Path $script:LogFile -Parent
                $canWriteLogFile = (Test-Path $logDir)
            } catch {
                $canWriteLogFile = $false
            }
        }
        if ($canWriteLogFile) {
            try {
                Add-Content -Path $script:LogFile -Value $line -Encoding UTF8 -ErrorAction Stop
            } catch {
                # Best-effort fallback for transient file lock contention.
                try {
                    $fileStream = [System.IO.File]::Open($script:LogFile,[System.IO.FileMode]::OpenOrCreate,[System.IO.FileAccess]::Write,[System.IO.FileShare]::ReadWrite)
                    try {
                        [void]$fileStream.Seek(0,[System.IO.SeekOrigin]::End)
                        $bytes = [System.Text.Encoding]::UTF8.GetBytes($line + [Environment]::NewLine)
                        $fileStream.Write($bytes,0,$bytes.Length)
                        $fileStream.Flush()
                    } finally {
                        $fileStream.Dispose()
                    }
                } catch {
                    # Do not throw from logger; console output still carries the message.
                }
            }
        }
        switch ($Level) {
            'ERROR' { Write-Host $line -ForegroundColor Red }
            'WARN'  { Write-Host $line -ForegroundColor Yellow }
            'DEBUG' { Write-Host $line -ForegroundColor DarkGray }
            default { Write-Host $line -ForegroundColor Gray }
        }
    }
    finally {
        [System.Threading.Monitor]::Exit($script:LogLock)
    }
}

# -------------------------
# Decimal helpers
# -------------------------
function Convert-ToDecimal {
    param($Value)
    if ($null -eq $Value) { return [decimal]0 }
    try { return [decimal]$Value } catch { return [decimal]0 }
}

function Round-Decimal {
    param([decimal]$Value,[int]$Scale=6)
    return [decimal]::Round($Value,$Scale,[System.MidpointRounding]::AwayFromZero)
}

function Get-MeasureSumSafe {
    param(
        $Rows,
        [Parameter(Mandatory)][string]$Property
    )

    $m = @($Rows) | Measure-Object -Property $Property -Sum
    if ($null -eq $m) { return [decimal]0 }
    $sumProp = $m.PSObject.Properties['Sum']
    if ($null -eq $sumProp) { return [decimal]0 }
    return (Convert-ToDecimal $sumProp.Value)
}

function Convert-BytesToDecimalGB { param($Bytes,[int]$Scale=6) return (Round-Decimal -Value ((Convert-ToDecimal $Bytes)/[decimal]1000000000) -Scale $Scale) }
function Convert-BytesToDecimalMB { param($Bytes,[int]$Scale=6) return (Round-Decimal -Value ((Convert-ToDecimal $Bytes)/[decimal]1000000) -Scale $Scale) }
function Convert-BytesToDecimalTB { param($Bytes,[int]$Scale=6) return (Round-Decimal -Value ((Convert-ToDecimal $Bytes)/[decimal]1000000000000) -Scale $Scale) }
function Convert-BytesToTiB { param($Bytes,[int]$Scale=6) return (Round-Decimal -Value ((Convert-ToDecimal $Bytes)/[decimal]1099511627776) -Scale $Scale) }
function Convert-GBToDecimalTB { param($GB,[int]$Scale=6) return (Round-Decimal -Value ((Convert-ToDecimal $GB)/[decimal]1000) -Scale $Scale) }
function Convert-GBToDecimalMB { param($GB,[int]$Scale=6) return (Round-Decimal -Value ((Convert-ToDecimal $GB)*[decimal]1000) -Scale $Scale) }
function Format-WholeNumber {
    param($Value)
    return ([decimal]::Round((Convert-ToDecimal $Value),0,[System.MidpointRounding]::AwayFromZero)).ToString('0',[System.Globalization.CultureInfo]::InvariantCulture)
}

function Convert-ToInvariantString {
    param($Value)
    if ($null -eq $Value) { return '' }
    if ($Value -is [datetime]) { return $Value.ToString('o',[System.Globalization.CultureInfo]::InvariantCulture) }
    if ($Value -is [decimal] -or $Value -is [double] -or $Value -is [single] -or $Value -is [float] -or $Value -is [int64] -or $Value -is [int32]) {
        return [string]::Format([System.Globalization.CultureInfo]::InvariantCulture,'{0}',$Value)
    }
    return $Value.ToString()
}

function Convert-RowToCsvLine {
    param([Parameter(Mandatory)]$Row,[Parameter(Mandatory)][string[]]$Columns)
    $vals = foreach ($c in $Columns) {
        $v = $Row.$c
        if ($null -eq $v) { '' }
        else { ((Convert-ToInvariantString -Value $v) -replace '"','') -replace '[\r\n]',' ' -replace ',', ';' }
    }
    return ($vals -join ',')
}

function Test-IsRetryableFileIoError {
    param($Exception)

    if ($null -eq $Exception) { return $false }
    if ($Exception -is [System.IO.IOException] -or $Exception -is [System.UnauthorizedAccessException]) { return $true }
    $msg = [string]$Exception.Message
    if ($msg -match 'being used by another process|cannot access the file|sharing violation|denied') { return $true }
    return $false
}

function Get-PendingFilePath {
    param([Parameter(Mandatory)][string]$Path)

    $dir = Split-Path -Path $Path -Parent
    if (-not $dir) { $dir = $PWD.Path }
    $name = [System.IO.Path]::GetFileNameWithoutExtension($Path)
    $ext = [System.IO.Path]::GetExtension($Path)
    return (Join-Path $dir ("{0}.pending{1}" -f $name,$ext))
}

function Set-ContentSafe {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)]$Value,
        [string]$Encoding = 'UTF8',
        [int]$MaxRetries = 6,
        [int]$RetryMs = 300,
        [switch]$AllowFallback,
        [switch]$PassThru
    )

    $attempt = 0
    while ($true) {
        $attempt++
        try {
            Microsoft.PowerShell.Management\Set-Content -Path $Path -Value $Value -Encoding $Encoding -ErrorAction Stop
            if ($PassThru) { return $Path }
            return
        } catch {
            if ($attempt -lt $MaxRetries -and (Test-IsRetryableFileIoError -Exception $_.Exception)) {
                Start-Sleep -Milliseconds $RetryMs
                continue
            }
            if ($AllowFallback) {
                $fallbackPath = Get-PendingFilePath -Path $Path
                Microsoft.PowerShell.Management\Set-Content -Path $fallbackPath -Value $Value -Encoding $Encoding -ErrorAction Stop
                Write-Log -Level WARN -Message ("[CSV-OpenFallback] set-content redirected to pending file. Target={0} Pending={1}" -f (Split-Path $Path -Leaf),(Split-Path $fallbackPath -Leaf))
                if ($PassThru) { return $fallbackPath }
                return
            }
            throw
        }
    }
}

function Add-ContentSafe {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)]$Value,
        [string]$Encoding = 'UTF8',
        [int]$MaxRetries = 6,
        [int]$RetryMs = 300,
        [bool]$AllowFallback = $true,
        [switch]$PassThru
    )

    $attempt = 0
    while ($true) {
        $attempt++
        try {
            Microsoft.PowerShell.Management\Add-Content -Path $Path -Value $Value -Encoding $Encoding -ErrorAction Stop
            if ($PassThru) { return $Path }
            return
        } catch {
            if ($attempt -lt $MaxRetries -and (Test-IsRetryableFileIoError -Exception $_.Exception)) {
                Start-Sleep -Milliseconds $RetryMs
                continue
            }
            if ($AllowFallback) {
                $fallbackPath = Get-PendingFilePath -Path $Path
                Microsoft.PowerShell.Management\Add-Content -Path $fallbackPath -Value $Value -Encoding $Encoding -ErrorAction Stop
                Write-Log -Level WARN -Message ("[CSV-OpenFallback] add-content redirected to pending file. Target={0} Pending={1}" -f (Split-Path $Path -Leaf),(Split-Path $fallbackPath -Leaf))
                if ($PassThru) { return $fallbackPath }
                return
            }
            throw
        }
    }
}

function Write-PlainCsv {
    param(
        [Parameter(Mandatory)]$Data,
        [Parameter(Mandatory)][string]$Path
    )
    $rows = @($Data)
    if ($rows.Count -eq 0) { return }
    $first = $rows[0]
    $cols = @($first.PSObject.Properties | Where-Object { $_.MemberType -eq 'NoteProperty' } | Select-Object -ExpandProperty Name)
    $targetPath = Set-ContentSafe -Path $Path -Value ($cols -join ',') -Encoding UTF8 -AllowFallback -PassThru
    $enc = New-Object System.Text.UTF8Encoding($false)
    $sw = $null
    $attempt = 0
    while (-not $sw -and $attempt -lt 6) {
        $attempt++
        try {
            $sw = New-Object System.IO.StreamWriter($targetPath,$true,$enc)
        } catch {
            if ($attempt -ge 6 -or -not (Test-IsRetryableFileIoError -Exception $_.Exception)) { break }
            Start-Sleep -Milliseconds 300
        }
    }
    if (-not $sw) {
        $fallbackPath = Get-PendingFilePath -Path $Path
        if ($fallbackPath -ne $targetPath) {
            $targetPath = Set-ContentSafe -Path $fallbackPath -Value ($cols -join ',') -Encoding UTF8 -AllowFallback -PassThru
        }
        $sw = New-Object System.IO.StreamWriter($targetPath,$true,$enc)
        Write-Log -Level WARN -Message ("[CSV-OpenFallback] stream writer redirected to pending file. Target={0} Pending={1}" -f (Split-Path $Path -Leaf),(Split-Path $targetPath -Leaf))
    }
    try {
        foreach ($r in $rows) {
            $sw.WriteLine((Convert-RowToCsvLine -Row $r -Columns $cols))
        }
        $sw.Flush()
    } finally {
        $sw.Dispose()
    }
}

function Copy-FileContentStreaming {
    param(
        [Parameter(Mandatory)][string]$SourcePath,
        [Parameter(Mandatory)][string]$DestinationPath
    )

    if (-not (Test-Path $SourcePath)) { return 0 }

    $copied = 0
    $targetPath = $DestinationPath
    $enc = New-Object System.Text.UTF8Encoding($false)
    $sr = New-Object System.IO.StreamReader($SourcePath,$enc,$true)
    $sw = $null
    $attempt = 0
    while (-not $sw -and $attempt -lt 6) {
        $attempt++
        try {
            $sw = New-Object System.IO.StreamWriter($targetPath,$true,$enc)
        } catch {
            if ($attempt -ge 6 -or -not (Test-IsRetryableFileIoError -Exception $_.Exception)) { break }
            Start-Sleep -Milliseconds 300
        }
    }
    if (-not $sw) {
        $targetPath = Get-PendingFilePath -Path $DestinationPath
        $sw = New-Object System.IO.StreamWriter($targetPath,$true,$enc)
        Write-Log -Level WARN -Message ("[CSV-OpenFallback] stream-copy redirected to pending file. Target={0} Pending={1}" -f (Split-Path $DestinationPath -Leaf),(Split-Path $targetPath -Leaf))
    }
    try {
        while ($null -ne ($line = $sr.ReadLine())) {
            $sw.WriteLine($line)
            $copied++
        }
        $sw.Flush()
    } finally {
        $sr.Dispose()
        $sw.Dispose()
    }

    return $copied
}

function Initialize-CsvFile {
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string[]]$Columns
    )

    [void](Set-ContentSafe -Path $Path -Value ($Columns -join ',') -Encoding UTF8 -AllowFallback)
}

function Append-CsvRows {
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)]$Rows,
        [Parameter(Mandatory)][string[]]$Columns,
        [int]$FlushThreshold = 0,
        [string]$Context = ''
    )

    $arr = @($Rows)
    if ($arr.Count -eq 0) { return }

    $threshold = [int]$FlushThreshold
    if ($threshold -lt 1) { $threshold = [int]$script:CsvFlushThreshold }
    if ($threshold -lt 1) { $threshold = 1 }

    $targetPath = $Path
    $enc = New-Object System.Text.UTF8Encoding($false)
    $sw = $null
    $attempt = 0
    while (-not $sw -and $attempt -lt 6) {
        $attempt++
        try {
            $sw = New-Object System.IO.StreamWriter($targetPath,$true,$enc)
        } catch {
            if ($attempt -ge 6 -or -not (Test-IsRetryableFileIoError -Exception $_.Exception)) { break }
            Start-Sleep -Milliseconds 300
        }
    }
    if (-not $sw) {
        $targetPath = Get-PendingFilePath -Path $Path
        $sw = New-Object System.IO.StreamWriter($targetPath,$true,$enc)
        Write-Log -Level WARN -Message ("[CSV-OpenFallback] append redirected to pending file. Target={0} Pending={1}" -f (Split-Path $Path -Leaf),(Split-Path $targetPath -Leaf))
    }
    try {
        $chunkCount = 0
        $fileKey = [string]$targetPath
        if (-not $script:CsvFlushedRowsByFile.ContainsKey($fileKey)) { $script:CsvFlushedRowsByFile[$fileKey] = 0 }

        foreach ($r in $arr) {
            $sw.WriteLine((Convert-RowToCsvLine -Row $r -Columns $Columns))
            $chunkCount++

            if ($chunkCount -ge $threshold) {
                $sw.Flush()
                $script:CsvFlushedRowsByFile[$fileKey] = [int]$script:CsvFlushedRowsByFile[$fileKey] + $chunkCount
                $ctx = if ($Context) { $Context } else { 'General' }
                Write-Log -Level INFO -Message ("[CSV-Flush] File={0} | Context={1} | ChunkRows={2} | TotalFlushedRows={3}" -f (Split-Path $Path -Leaf),$ctx,$chunkCount,$script:CsvFlushedRowsByFile[$fileKey])
                $chunkCount = 0
            }
        }

        if ($chunkCount -gt 0) {
            $sw.Flush()
            $script:CsvFlushedRowsByFile[$fileKey] = [int]$script:CsvFlushedRowsByFile[$fileKey] + $chunkCount
            $ctx = if ($Context) { $Context } else { 'General' }
            Write-Log -Level INFO -Message ("[CSV-Flush] File={0} | Context={1} | ChunkRows={2} | TotalFlushedRows={3}" -f (Split-Path $Path -Leaf),$ctx,$chunkCount,$script:CsvFlushedRowsByFile[$fileKey])
        }
    } finally {
        $sw.Dispose()
    }
}

function Normalize-Key {
    param(
        [string]$Key,
        [string]$Default = 'Unknown'
    )

    if ([string]::IsNullOrWhiteSpace($Key)) { return $Default }
    return $Key
}

function Add-CountToMap {
    param(
        [Parameter(Mandatory)][hashtable]$Map,
        [string]$Key,
        [int]$Increment = 1,
        [string]$Default = 'Unknown'
    )

    $k = Normalize-Key -Key $Key -Default $Default
    if (-not $Map.ContainsKey($k)) { $Map[$k] = 0 }
    $Map[$k] = [int]$Map[$k] + $Increment
}

function Add-DecimalToMap {
    param(
        [Parameter(Mandatory)][hashtable]$Map,
        [string]$Key,
        [decimal]$Increment,
        [string]$Default = 'Unknown'
    )

    $k = Normalize-Key -Key $Key -Default $Default
    if (-not $Map.ContainsKey($k)) { $Map[$k] = [decimal]0 }
    $Map[$k] = [decimal]$Map[$k] + (Convert-ToDecimal $Increment)
}

function Get-MapCount {
    param(
        [Parameter(Mandatory)][hashtable]$Map,
        [string]$Key,
        [string]$Default = 'Unknown'
    )

    $k = Normalize-Key -Key $Key -Default $Default
    if ($Map.ContainsKey($k)) { return [int]$Map[$k] }
    return 0
}

function Get-MapDecimal {
    param(
        [Parameter(Mandatory)][hashtable]$Map,
        [string]$Key,
        [string]$Default = 'Unknown'
    )

    $k = Normalize-Key -Key $Key -Default $Default
    if ($Map.ContainsKey($k)) { return (Convert-ToDecimal $Map[$k]) }
    return [decimal]0
}

function New-Key2 {
    param(
        [string]$A,
        [string]$B,
        [string]$DefaultA = 'Unknown',
        [string]$DefaultB = 'Unknown'
    )

    return ((Normalize-Key -Key $A -Default $DefaultA) + '||' + (Normalize-Key -Key $B -Default $DefaultB))
}

function Split-Key2 {
    param([Parameter(Mandatory)][string]$Key)

    $parts = $Key -split '\|\|',2
    if ($parts.Count -lt 2) { return @($parts[0],'') }
    return $parts
}

function New-DbAggregateStats {
    return @{
        RowCount = 0
        TotalMb = [decimal]0
        TotalGb = [decimal]0
        ProjectCount = @{}
        ProjectMb = @{}
        ProjectGb = @{}
        RegionCount = @{}
        RegionMb = @{}
        RegionGb = @{}
    }
}

function Get-GkeEffectiveSizeFromSums {
    param(
        [decimal]$PvGb,
        [decimal]$NodeGb
    )

    if ($PvGb -gt 0) { return (Round-Decimal -Value $PvGb -Scale 6) }
    return (Round-Decimal -Value $NodeGb -Scale 6)
}

function Update-RunProgress {
    param(
        [Parameter(Mandatory)][int]$CompletedProjects,
        [Parameter(Mandatory)][int]$CurrentProjectNumber,
        [Parameter(Mandatory)][int]$TotalProjects,
        [Parameter(Mandatory)][string]$Project,
        [Parameter(Mandatory)][string]$Stage
    )

    $pct = 0
    if ($TotalProjects -gt 0) {
        $pct = [Math]::Min(100,[Math]::Max(0,[int](($CompletedProjects * 100) / $TotalProjects)))
    }
    $current = $CurrentProjectNumber
    if ($current -lt 1) { $current = 1 }
    if ($TotalProjects -gt 0 -and $current -gt $TotalProjects) { $current = $TotalProjects }
    $elapsed = (Get-Date) - $script:RunStart
    $elapsedTxt = $elapsed.ToString('hh\:mm\:ss')
    $statusText = ("{0}% | #Project {1}/{2} | Project={3} | Stage={4} | Elapsed={5}" -f $pct,$current,$TotalProjects,$Project,$Stage,$elapsedTxt)

    Write-Progress -Id 1 -Activity 'GCP API Inventory Progress' -Status $statusText -PercentComplete $pct
    Write-Log -Level INFO -Message ("[Progress] {0}% | Completed={1}/{2} | CurrentProject={3}/{2} | Project={4} | Stage={5} | Elapsed={6}" -f $pct,$CompletedProjects,$TotalProjects,$current,$Project,$Stage,$elapsedTxt)
}

function Get-ProjectWorkerCount {
    param([int]$TotalProjects)

    if ($TotalProjects -le 1) { return 1 }

    if ($ProjectWorkers -gt 0) {
        $count = $ProjectWorkers
    } else {
        $auto = [Math]::Max(2,[Math]::Min(4,[int]$script:RecommendedMaxThreads))
        if ($script:HostMemoryGb -le 16) { $auto = [Math]::Min($auto,3) }
        if ($script:HostMemoryGb -le 8) { $auto = 2 }
        $count = $auto
    }

    if ($count -gt $TotalProjects) { $count = $TotalProjects }
    if ($count -lt 1) { $count = 1 }
    return [int]$count
}

function Invoke-ProjectWorkerPool {
    param(
        [Parameter(Mandatory)][string[]]$TargetProjects,
        [Parameter(Mandatory)][string]$RootOutputDir
    )

    $workerCount = Get-ProjectWorkerCount -TotalProjects $TargetProjects.Count
    if ($workerCount -le 1) {
        return [PSCustomObject]@{ WorkerCount = 1; ManifestPath = ''; FailedCount = 0 }
    }

    $scriptPath = $PSCommandPath
    if (-not $scriptPath) { $scriptPath = $MyInvocation.MyCommand.Path }
    if (-not $scriptPath) { throw 'Unable to resolve current script path for worker pool launch.' }

    Write-Log -Level INFO -Message ("[WorkerPool] Enabled | Workers={0} | Projects={1}" -f $workerCount,$TargetProjects.Count)

    $workerRoot = Join-Path $RootOutputDir 'workers'
    $workerInputRoot = Join-Path $RootOutputDir 'worker-inputs'
    New-Item -Path $workerRoot -ItemType Directory -Force | Out-Null
    New-Item -Path $workerInputRoot -ItemType Directory -Force | Out-Null

    $buckets = @()
    for ($i = 0; $i -lt $workerCount; $i++) {
        $buckets += ,(New-Object System.Collections.Generic.List[string])
    }
    for ($i = 0; $i -lt $TargetProjects.Count; $i++) {
        $bucketIdx = $i % $workerCount
        $buckets[$bucketIdx].Add([string]$TargetProjects[$i]) | Out-Null
    }

    $workers = New-Object System.Collections.Generic.List[object]
    for ($w = 0; $w -lt $workerCount; $w++) {
        $bucketProjects = @($buckets[$w])
        if ($bucketProjects.Count -eq 0) { continue }

        $workerId = $w + 1
        $inputFile = Join-Path $workerInputRoot ("worker-{0:d2}-projects.txt" -f $workerId)
        Set-Content -Path $inputFile -Value $bucketProjects -Encoding UTF8

        $workerOutDir = Join-Path $workerRoot ("worker-{0:d2}" -f $workerId)
        New-Item -Path $workerOutDir -ItemType Directory -Force | Out-Null

        $argList = @('-NoProfile','-ExecutionPolicy','Bypass','-File',"`"$scriptPath`"")
        if ($ServiceAccount) { $argList += @('-ServiceAccount',"`"$ServiceAccount`"") }
        if ($AccessToken) { $argList += @('-AccessToken',"`"$AccessToken`"") }
        if ($Types -and $Types.Count -gt 0) {
            foreach ($t in $Types) { $argList += @('-Types',$t) }
        }
        $argList += @('-ProjectsFile',"`"$inputFile`"")
        $argList += @('-OutputDir',"`"$workerOutDir`"")
        $argList += @('-MaxThreads',[string]$MaxThreads)
        $argList += @('-ApiTimeoutSec',[string]$ApiTimeoutSec)
        $argList += @('-MetricLookbackHours',[string]$MetricLookbackHours)
        $argList += @('-CsvFlushThreshold',[string]$CsvFlushThreshold)
        if ($LargeScaleMode) { $argList += '-LargeScaleMode' }
        $argList += '-SkipZip'
        $argList += '-KeepOutputDir'
        $argList += '-DisableProjectWorkerPool'

        $proc = Start-Process -FilePath 'powershell.exe' -ArgumentList $argList -PassThru -WindowStyle Hidden
        $workers.Add([PSCustomObject]@{
            WorkerId = $workerId
            ProjectsCount = $bucketProjects.Count
            FirstProject = [string]$bucketProjects[0]
            LastProject = [string]$bucketProjects[$bucketProjects.Count - 1]
            ProjectsFile = $inputFile
            OutputDir = $workerOutDir
            Process = $proc
        }) | Out-Null
        Write-Log -Level INFO -Message ("[WorkerPool] Worker={0} Pid={1} Projects={2} First={3} Last={4}" -f $workerId,$proc.Id,$bucketProjects.Count,[string]$bucketProjects[0],[string]$bucketProjects[$bucketProjects.Count - 1])
    }

    $lastHeartbeat = Get-Date
    while ($true) {
        $running = @($workers | Where-Object { -not $_.Process.HasExited })
        $completed = $workers.Count - $running.Count
        $now = Get-Date
        if ($running.Count -eq 0 -or ($now - $lastHeartbeat).TotalSeconds -ge 15) {
            Write-Log -Level INFO -Message ("[WorkerPool] Running={0}/{1} | Completed={2}/{1}" -f $running.Count,$workers.Count,$completed)
            $lastHeartbeat = $now
        }
        if ($running.Count -eq 0) { break }
        Start-Sleep -Seconds 5
    }

    $manifestRows = New-Object System.Collections.Generic.List[object]
    foreach ($w in @($workers | Sort-Object WorkerId)) {
        $exitCode = 0
        try { $exitCode = [int]$w.Process.ExitCode } catch { $exitCode = -1 }
        $manifestRows.Add([PSCustomObject]@{
            WorkerId = [int]$w.WorkerId
            Pid = [int]$w.Process.Id
            ProjectsCount = [int]$w.ProjectsCount
            FirstProject = [string]$w.FirstProject
            LastProject = [string]$w.LastProject
            ProjectsFile = [string]$w.ProjectsFile
            OutputDir = [string]$w.OutputDir
            ExitCode = $exitCode
            Status = if ($exitCode -eq 0) { 'Success' } else { 'Failed' }
        }) | Out-Null
    }

    $manifestPath = Join-Path $RootOutputDir ("project_worker_manifest_{0}.csv" -f $dateStr)
    Write-PlainCsv -Data $manifestRows.ToArray() -Path $manifestPath
    $failedCount = @($manifestRows | Where-Object { $_.ExitCode -ne 0 }).Count
    Write-Log -Level INFO -Message ("[WorkerPool] Manifest written: {0} | FailedWorkers={1}" -f (Split-Path $manifestPath -Leaf),$failedCount)

    return [PSCustomObject]@{ WorkerCount = $workers.Count; ManifestPath = $manifestPath; FailedCount = $failedCount }
}

function Get-StatusCodeFromMessage {
    param([string]$Message)
    if (-not $Message) { return '' }
    $m = [regex]::Match($Message,'\((\d{3})\)')
    if ($m.Success) { return [string]$m.Groups[1].Value }
    return ''
}

function Get-KnownErrorCategory {
    param(
        [string]$Message,
        [string]$StatusCode = ''
    )

    $msg = if ($Message) { $Message.ToLowerInvariant() } else { '' }

    if ($msg -match 'service_disabled|api has not been used|is disabled|not enabled|enable it by visiting') { return 'ApiDisabled' }
    if ($StatusCode -eq '401' -or $msg -match 'unauthenticated|unauthorized|invalid credentials|invalid_grant') { return 'AuthError' }
    if ($StatusCode -eq '403' -or $msg -match 'permission_denied|insufficientpermissions|permission|forbidden|denied') { return 'PermissionDenied' }
    if ($StatusCode -eq '404' -or $msg -match 'not found') { return 'NotFound' }
    if ($StatusCode -eq '429' -or $msg -match 'quota|rate limit|too many requests') { return 'RateLimited' }
    if ($StatusCode -eq '400' -or $msg -match 'bad request|invalid argument|invalid') { return 'BadRequest' }
    if ($StatusCode -eq '408' -or $StatusCode -eq '504' -or $msg -match 'timeout|timed out') { return 'Timeout' }
    if ($StatusCode -match '^5\d\d$' -or $msg -match 'internal server error|server error|unavailable') { return 'ServerError' }

    return 'UnknownError'
}

function Get-WorkloadFromMonitoringFilter {
    param([string]$Filter)

    if (-not $Filter) { return 'Monitoring' }
    $f = $Filter.ToLowerInvariant()

    if ($f -match 'storage.googleapis.com/storage/total_bytes') { return 'StorageBucket' }
    if ($f -match 'kubernetes.io/persistentvolume/capacity_bytes') { return 'GKECluster' }
    if ($f -match 'spanner.googleapis.com/instance/storage/used_bytes') { return 'Database' }
    if ($f -match 'bigquery.googleapis.com/storage/logical_bytes|bigquery.googleapis.com/storage/stored_bytes') { return 'Database' }
    if ($f -match 'alloydb.googleapis.com/cluster/storage/usage') { return 'Database' }

    return 'Monitoring'
}

function Get-ComponentFromMonitoringFilter {
    param([string]$Filter)

    if (-not $Filter) { return 'Monitoring' }
    $f = $Filter.ToLowerInvariant()

    if ($f -match 'spanner.googleapis.com/instance/storage/used_bytes') { return 'Spanner' }
    if ($f -match 'bigquery.googleapis.com/storage/logical_bytes|bigquery.googleapis.com/storage/stored_bytes') { return 'BigQuery' }
    if ($f -match 'alloydb.googleapis.com/cluster/storage/usage') { return 'AlloyDB' }
    if ($f -match 'kubernetes.io/persistentvolume/capacity_bytes') { return 'GKE' }
    if ($f -match 'storage.googleapis.com/storage/total_bytes') { return 'Storage' }

    return 'Monitoring'
}

function Get-InventoryErrorRecordsFromLog {
    param([Parameter(Mandatory)][string]$LogPath)

    if (-not (Test-Path $LogPath)) { return @() }

    $rows = New-Object System.Collections.Generic.List[object]
    $enc = New-Object System.Text.UTF8Encoding($false)
    $fs = [System.IO.File]::Open($LogPath,[System.IO.FileMode]::Open,[System.IO.FileAccess]::Read,[System.IO.FileShare]::ReadWrite)
    $reader = New-Object System.IO.StreamReader($fs,$enc,$true)
    try {
        while ($null -ne ($line = $reader.ReadLine())) {
        if ($line -match '^\[(?<ts>[^\]]+)\]\s+\[WARN\]\s+\[(?<tag1>[^\]]+)\](?:\[(?<tag2>[^\]]+)\])?\s+Project=(?<project>[^\s]+)\s+(?<action>[^:]+):\s*(?<message>.*)$') {
            $tag1 = [string]$Matches.tag1
            $workload = $tag1
            switch ($tag1.ToUpperInvariant()) {
                'VM' { $workload = 'VM' }
                'STORAGE' { $workload = 'StorageBucket' }
                'FILESHARE' { $workload = 'Fileshare' }
                'FS' { $workload = 'Fileshare' }
                'DB' { $workload = 'Database' }
                'GKE' { $workload = 'GKECluster' }
                default { $workload = $tag1 }
            }

            $hasTag2 = ($Matches.ContainsKey('tag2') -and $null -ne $Matches['tag2'] -and [string]$Matches['tag2'] -ne '')
            $component = if ($hasTag2) { [string]$Matches['tag2'] } elseif ($tag1 -eq 'FS') { 'FileShareAPI' } else { '' }
            $message = [string]$Matches.message
            $statusCode = Get-StatusCodeFromMessage -Message $message
            $category = Get-KnownErrorCategory -Message $message -StatusCode $statusCode

            $rows.Add([PSCustomObject]@{
                Timestamp = [string]$Matches.ts
                Project = [string]$Matches.project
                Workload = $workload
                Component = $component
                Action = [string]$Matches.action
                ErrorCategory = $category
                HttpStatusCode = $statusCode
                KnownError = if ($category -eq 'UnknownError') { 'N' } else { 'Y' }
                Message = $message
            }) | Out-Null
            continue
        }

        if ($line -match '^\[(?<ts>[^\]]+)\]\s+\[WARN\]\s+Monitoring query failed for project=(?<project>[^\.]+)\.\s+filter=(?<filter>.*?)\.\s+(?<message>.*)$') {
            $message = [string]$Matches.message
            $statusCode = Get-StatusCodeFromMessage -Message $message
            $category = Get-KnownErrorCategory -Message $message -StatusCode $statusCode
            $filter = [string]$Matches.filter
            $workload = Get-WorkloadFromMonitoringFilter -Filter $filter
            $component = Get-ComponentFromMonitoringFilter -Filter $filter

            $rows.Add([PSCustomObject]@{
                Timestamp = [string]$Matches.ts
                Project = [string]$Matches.project
                Workload = $workload
                Component = $component
                Action = 'MetricQuery'
                ErrorCategory = $category
                HttpStatusCode = $statusCode
                KnownError = if ($category -eq 'UnknownError') { 'N' } else { 'Y' }
                Message = $message
            }) | Out-Null
            continue
        }
        }
    } finally {
        $reader.Dispose()
        $fs.Dispose()
    }

    return $rows.ToArray()
}

function Get-ErrorPriorityRank {
    param([string]$Category)

    switch ($Category) {
        'PermissionDenied' { return 1 }
        'ApiDisabled' { return 2 }
        'AuthError' { return 3 }
        'Timeout' { return 4 }
        'RateLimited' { return 5 }
        'NotFound' { return 6 }
        'BadRequest' { return 7 }
        'ServerError' { return 8 }
        default { return 99 }
    }
}

function Convert-ToCsvReasonText {
    param(
        [string]$Text,
        [int]$MaxLen = 220
    )

    if (-not $Text) { return '' }
    $out = $Text -replace '[\r\n]',' '
    if ($out.Length -gt $MaxLen) { $out = $out.Substring(0,$MaxLen-3) + '...' }
    return ($out -replace ',',';')
}

function Select-PreferredErrorRecord {
    param($Rows)

    $best = $null
    $bestRank = 999
    foreach ($r in @($Rows)) {
        $rank = Get-ErrorPriorityRank -Category ([string]$r.ErrorCategory)
        if ($rank -lt $bestRank) {
            $best = $r
            $bestRank = $rank
        }
    }
    return $best
}

function Write-DbWorkloadCsvWithSummary {
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string]$WorkloadType,
        [Parameter(Mandatory)][string]$DetailCsvPath,
        [Parameter(Mandatory)][hashtable]$WorkloadStats,
        [Parameter(Mandatory)][string[]]$AllProjects,
        [Parameter(Mandatory)]$DbErrors
    )

    $errors = @($DbErrors)

    $components = @(
        switch ($WorkloadType) {
        'CloudSQL' { @('CloudSQL') }
        'SpannerInstance' { @('Spanner') }
        'BigQueryDataset' { @('BigQuery') }
        'AlloyDBCluster' { @('AlloyDB') }
        default { @() }
        }
    )

    $failed = New-Object System.Collections.Generic.List[object]
    foreach ($proj in @($AllProjects)) {
        $specific = @($errors | Where-Object { $_.Project -eq $proj -and ($components -contains [string]$_.Component) })
        $generic = @($errors | Where-Object { $_.Project -eq $proj -and [string]$_.Component -eq '' -and [string]$_.Action -eq 'failed' })
        $candidates = if ((@($specific)).Count -gt 0) { $specific } else { $generic }
        if ((@($candidates)).Count -eq 0) { continue }

        $best = Select-PreferredErrorRecord -Rows $candidates
        if ($null -eq $best) { continue }
        $reason = [string]$best.ErrorCategory
        if ($best.HttpStatusCode) { $reason = ("{0}({1})" -f $reason,[string]$best.HttpStatusCode) }

        $failed.Add([PSCustomObject]@{
            Project = [string]$proj
            Reason = $reason
            Component = [string]$best.Component
            Message = Convert-ToCsvReasonText -Text ([string]$best.Message)
        }) | Out-Null
    }

    Set-Content -Path $Path -Value 'DetailedInventory' -Encoding UTF8
    if (Test-Path $DetailCsvPath) {
        $copiedLines = Copy-FileContentStreaming -SourcePath $DetailCsvPath -DestinationPath $Path
        Write-Log -Level INFO -Message ("[DB-Finalize] Workload={0} | DetailLinesCopied={1} | File={2}" -f $WorkloadType,$copiedLines,(Split-Path $Path -Leaf))
    } else {
        Add-ContentSafe -Path $Path -Value 'Type,Project,Name,Region,Zone,Engine,TierOrCapacity,StorageType,NodeCount,TableCount,State,StorageMB,StorageGB,Extra,Error'
        Write-Log -Level WARN -Message ("[DB-Finalize] Workload={0} detail file missing. Writing header-only detail section for {1}." -f $WorkloadType,(Split-Path $Path -Leaf))
    }

    Add-ContentSafe -Path $Path -Value ''
    Add-ContentSafe -Path $Path -Value 'Summary'
    $rowCount = [int]$WorkloadStats.RowCount
    $totalMb = Convert-ToDecimal $WorkloadStats.TotalMb
    $totalGb = Convert-ToDecimal $WorkloadStats.TotalGb
    Add-ContentSafe -Path $Path -Value ("Overall,{0},Rows,{1},TotalSizeMB,{2},TotalSizeGB,{3},TotalSizeTB,{4}" -f $WorkloadType,$rowCount,(Round-Decimal -Value $totalMb -Scale 6),(Round-Decimal -Value $totalGb -Scale 6),(Convert-GBToDecimalTB -GB $totalGb -Scale 6))

    Add-ContentSafe -Path $Path -Value 'ProjectSummary'
    Add-ContentSafe -Path $Path -Value 'Project,Rows,TotalSizeMB,TotalSizeGB,TotalSizeTB'
    foreach ($project in @($WorkloadStats.ProjectCount.Keys | Sort-Object)) {
        $pCount = Get-MapCount -Map $WorkloadStats.ProjectCount -Key $project
        $pMb = Get-MapDecimal -Map $WorkloadStats.ProjectMb -Key $project
        $pGb = Get-MapDecimal -Map $WorkloadStats.ProjectGb -Key $project
        Add-ContentSafe -Path $Path -Value ("{0},{1},{2},{3},{4}" -f $project,$pCount,(Round-Decimal -Value $pMb -Scale 6),(Round-Decimal -Value $pGb -Scale 6),(Convert-GBToDecimalTB -GB $pGb -Scale 6))
    }

    Add-ContentSafe -Path $Path -Value 'RegionSummary'
    Add-ContentSafe -Path $Path -Value 'Region,Rows,TotalSizeMB,TotalSizeGB,TotalSizeTB'
    foreach ($region in @($WorkloadStats.RegionCount.Keys | Sort-Object)) {
        $rCount = Get-MapCount -Map $WorkloadStats.RegionCount -Key $region
        $rMb = Get-MapDecimal -Map $WorkloadStats.RegionMb -Key $region
        $rGb = Get-MapDecimal -Map $WorkloadStats.RegionGb -Key $region
        Add-ContentSafe -Path $Path -Value ("{0},{1},{2},{3},{4}" -f $region,$rCount,(Round-Decimal -Value $rMb -Scale 6),(Round-Decimal -Value $rGb -Scale 6),(Convert-GBToDecimalTB -GB $rGb -Scale 6))
    }

    Add-ContentSafe -Path $Path -Value ''
    Add-ContentSafe -Path $Path -Value 'FailedProjects'
    Add-ContentSafe -Path $Path -Value 'Project,Reason,Component,Message'
    if ($failed.Count -eq 0) {
        Add-ContentSafe -Path $Path -Value 'None,NoFailure,,No DB workload failures detected'
    } else {
        foreach ($f in @($failed | Sort-Object Project)) {
            Add-ContentSafe -Path $Path -Value ("{0},{1},{2},{3}" -f $f.Project,$f.Reason,$f.Component,$f.Message)
        }
    }
}

# -------------------------
# Auth helpers (API only)
# -------------------------
function ConvertTo-Base64Url {
    param([byte[]]$Bytes)
    [Convert]::ToBase64String($Bytes).TrimEnd('=') -replace '\+','-' -replace '/','_'
}

function Get-OptionalPropertyValue {
    param(
        [Parameter(Mandatory)]$Object,
        [Parameter(Mandatory)][string]$Name,
        $Default = $null
    )

    if ($null -eq $Object) { return $Default }
    $prop = $Object.PSObject.Properties[$Name]
    if ($null -eq $prop -or $null -eq $prop.Value) { return $Default }
    return $prop.Value
}

function New-RsaFromPemPrivateKey {
    param([Parameter(Mandatory)][string]$PemPrivateKey)

    $rsa = [System.Security.Cryptography.RSA]::Create()
    if ($rsa | Get-Member -Name ImportFromPem -MemberType Method -ErrorAction SilentlyContinue) {
        $rsa.ImportFromPem($PemPrivateKey)
        return $rsa
    }
    $rsa.Dispose()

    # Windows PowerShell/.NET Framework fallback for PKCS#8 service account keys.
    if ($PemPrivateKey -notmatch 'BEGIN PRIVATE KEY') {
        throw 'Unsupported private key format for this runtime. Provide a PKCS#8 PEM key (BEGIN PRIVATE KEY) or run with PowerShell 7+.'
    }

    $pkcs8Body = ($PemPrivateKey -replace '-----BEGIN PRIVATE KEY-----','' -replace '-----END PRIVATE KEY-----','' -replace '\s','')
    $pkcs8Bytes = [Convert]::FromBase64String($pkcs8Body)

    try {
        $cngKey = [System.Security.Cryptography.CngKey]::Import(
            $pkcs8Bytes,
            [System.Security.Cryptography.CngKeyBlobFormat]::Pkcs8PrivateBlob,
            [System.Security.Cryptography.CngProvider]::MicrosoftSoftwareKeyStorageProvider
        )
    } catch {
        throw 'This PowerShell runtime cannot import the service account private key. Use PowerShell 7+ or pass -AccessToken.'
    }

    return (New-Object System.Security.Cryptography.RSACng($cngKey))
}

function Get-ServiceAccountAccessToken {
    param([Parameter(Mandatory)][string]$KeyFile)

    if (-not (Test-Path $KeyFile)) { throw "Service account key file not found: $KeyFile" }
    $keyJson = Get-Content -Path $KeyFile -Raw | ConvertFrom-Json
    if (-not $keyJson.client_email -or -not $keyJson.private_key) {
        throw 'Invalid service account JSON. Missing client_email/private_key.'
    }

    $now = [DateTimeOffset]::UtcNow.ToUnixTimeSeconds()
    $header = @{ alg='RS256'; typ='JWT' }
    $claim = @{
        iss   = [string]$keyJson.client_email
        scope = 'https://www.googleapis.com/auth/cloud-platform'
        aud   = 'https://oauth2.googleapis.com/token'
        iat   = $now
        exp   = $now + 3600
    }

    $headerJson = ($header | ConvertTo-Json -Compress)
    $claimJson = ($claim | ConvertTo-Json -Compress)

    $unsigned = (ConvertTo-Base64Url -Bytes ([System.Text.Encoding]::UTF8.GetBytes($headerJson))) + '.' +
                (ConvertTo-Base64Url -Bytes ([System.Text.Encoding]::UTF8.GetBytes($claimJson)))

    $rsa = New-RsaFromPemPrivateKey -PemPrivateKey ([string]$keyJson.private_key)
    try {
        $sig = $rsa.SignData([System.Text.Encoding]::UTF8.GetBytes($unsigned), [System.Security.Cryptography.HashAlgorithmName]::SHA256, [System.Security.Cryptography.RSASignaturePadding]::Pkcs1)
    } finally {
        $rsa.Dispose()
    }

    $jwt = $unsigned + '.' + (ConvertTo-Base64Url -Bytes $sig)

    $tokenResp = Invoke-RestMethod -Method Post -Uri 'https://oauth2.googleapis.com/token' -Body @{
        grant_type = 'urn:ietf:params:oauth:grant-type:jwt-bearer'
        assertion = $jwt
    } -ContentType 'application/x-www-form-urlencoded' -TimeoutSec $ApiTimeoutSec

    if (-not $tokenResp.access_token) { throw 'Failed to obtain service account access token.' }
    return [string]$tokenResp.access_token
}

function Get-MetadataAccessToken {
    try {
        $uri = 'http://metadata.google.internal/computeMetadata/v1/instance/service-accounts/default/token'
        $resp = Invoke-RestMethod -Method Get -Uri $uri -Headers @{ 'Metadata-Flavor'='Google' } -TimeoutSec 5
        if ($resp.access_token) { return [string]$resp.access_token }
    } catch {}
    return $null
}

function Get-GcpAccessToken {
    if ($AccessToken) {
        Write-Log -Level INFO -Message 'Using caller supplied access token.'
        return $AccessToken
    }
    if ($ServiceAccount) {
        Write-Log -Level INFO -Message ("Acquiring token from service account key: {0}" -f $ServiceAccount)
        return (Get-ServiceAccountAccessToken -KeyFile $ServiceAccount)
    }
    if ($env:GOOGLE_OAUTH_ACCESS_TOKEN) {
        Write-Log -Level INFO -Message 'Using GOOGLE_OAUTH_ACCESS_TOKEN from environment.'
        return [string]$env:GOOGLE_OAUTH_ACCESS_TOKEN
    }
    $metadataToken = Get-MetadataAccessToken
    if ($metadataToken) {
        Write-Log -Level INFO -Message 'Using metadata server access token.'
        return $metadataToken
    }
    throw 'No API auth source found. Provide -ServiceAccount or -AccessToken.'
}

$script:GcpAccessToken = Get-GcpAccessToken

# -------------------------
# API core helpers
# -------------------------
function Apply-ApiAdaptiveDelay {
    $delayMs = [int]$script:ApiAdaptiveDelayMs
    if ($delayMs -le 0) { return }

    $elapsedMs = 0
    if ($script:ApiLastCallUtc -ne [datetime]::MinValue) {
        $elapsedMs = [int]((Get-Date).ToUniversalTime().Subtract($script:ApiLastCallUtc).TotalMilliseconds)
        if ($elapsedMs -lt 0) { $elapsedMs = 0 }
    }

    $waitMs = $delayMs - $elapsedMs
    if ($waitMs -gt 0) {
        Start-Sleep -Milliseconds $waitMs
    }
}

function Update-ApiAdaptiveDelay {
    param(
        [bool]$Success,
        [int]$StatusCode = 0
    )

    if ($Success) {
        if ($script:ApiAdaptiveDelayMs -gt 0) {
            $script:ApiAdaptiveDelayMs = [int][Math]::Floor($script:ApiAdaptiveDelayMs * 0.85)
            if ($script:ApiAdaptiveDelayMs -lt 20) { $script:ApiAdaptiveDelayMs = 0 }
        }
        return
    }

    if ($StatusCode -eq 429) {
        $next = [int]([Math]::Max(120,($script:ApiAdaptiveDelayMs * 2) + 120))
        $script:ApiAdaptiveDelayMs = [Math]::Min($script:ApiAdaptiveDelayMaxMs,$next)
        return
    }
    if ($StatusCode -ge 500 -and $StatusCode -lt 600) {
        $next = [int]([Math]::Max(80,($script:ApiAdaptiveDelayMs * 13 / 10) + 60))
        $script:ApiAdaptiveDelayMs = [Math]::Min($script:ApiAdaptiveDelayMaxMs,$next)
    }
}

function Invoke-GcpApi {
    param(
        [Parameter(Mandatory)][string]$Uri,
        [ValidateSet('GET','POST','PATCH','PUT','DELETE')][string]$Method = 'GET',
        $Body = $null,
        [int]$Retries = 5
    )

    $attempt = 0
    while ($true) {
        $attempt++
        try {
            Apply-ApiAdaptiveDelay
            $headers = @{ Authorization = "Bearer $($script:GcpAccessToken)" }
            $script:ApiLastCallUtc = (Get-Date).ToUniversalTime()
            if ($null -eq $Body) {
                $resp = Invoke-RestMethod -Method $Method -Uri $Uri -Headers $headers -TimeoutSec $ApiTimeoutSec
                Update-ApiAdaptiveDelay -Success $true -StatusCode 200
                return $resp
            }
            $jsonBody = $Body | ConvertTo-Json -Depth 20 -Compress
            $resp = Invoke-RestMethod -Method $Method -Uri $Uri -Headers $headers -Body $jsonBody -ContentType 'application/json' -TimeoutSec $ApiTimeoutSec
            Update-ApiAdaptiveDelay -Success $true -StatusCode 200
            return $resp
        }
        catch {
            $statusCode = $null
            try { $statusCode = [int]$_.Exception.Response.StatusCode } catch {}
            $retryable = ($statusCode -eq 429 -or ($statusCode -ge 500 -and $statusCode -lt 600))
            $statusForDelay = 0
            if ($statusCode) { $statusForDelay = [int]$statusCode }
            Update-ApiAdaptiveDelay -Success $false -StatusCode $statusForDelay
            if ($attempt -ge $Retries -or -not $retryable) {
                throw
            }
            $delay = [Math]::Min(30,[Math]::Pow(2,$attempt))
            Write-Log -Level WARN -Message ("API retry {0}/{1} status={2} uri={3} adaptiveDelayMs={4}" -f $attempt,$Retries,$statusCode,$Uri,$script:ApiAdaptiveDelayMs)
            Start-Sleep -Seconds $delay
        }
    }
}

function Get-PagedItems {
    param(
        [Parameter(Mandatory)][string]$BaseUri,
        [string]$ItemsProperty = 'items'
    )

    $all = New-Object System.Collections.Generic.List[object]
    $pageToken = $null
    do {
        $uri = if ($pageToken) { "$BaseUri&pageToken=$([uri]::EscapeDataString($pageToken))" } else { $BaseUri }
        $resp = Invoke-GcpApi -Uri $uri
        $items = Get-OptionalPropertyValue -Object $resp -Name $ItemsProperty
        if ($items) { $all.AddRange([object[]]@($items)) }
        $pageToken = [string](Get-OptionalPropertyValue -Object $resp -Name 'nextPageToken')
    } while ($pageToken)
    return $all.ToArray()
}

function Get-MonitoringSeries {
    param(
        [Parameter(Mandatory)][string]$ProjectId,
        [Parameter(Mandatory)][string]$Filter,
        [int]$LookbackHours = 24
    )

    $end = (Get-Date).ToUniversalTime().ToString('o')
    $start = (Get-Date).ToUniversalTime().AddHours(-1 * $LookbackHours).ToString('o')
    $filterEsc = [uri]::EscapeDataString($Filter)
    $baseUri = "https://monitoring.googleapis.com/v3/projects/$ProjectId/timeSeries?filter=$filterEsc&interval.startTime=$([uri]::EscapeDataString($start))&interval.endTime=$([uri]::EscapeDataString($end))&aggregation.alignmentPeriod=3600s&aggregation.perSeriesAligner=ALIGN_MAX&pageSize=10000"

    $all = New-Object System.Collections.Generic.List[object]
    $pageToken = $null
    do {
        $uri = if ($pageToken) { "$baseUri&pageToken=$([uri]::EscapeDataString($pageToken))" } else { $baseUri }
        try {
            $resp = Invoke-GcpApi -Uri $uri
        } catch {
            Write-Log -Level WARN -Message ("Monitoring query failed for project={0}. filter={1}. {2}" -f $ProjectId,$Filter,$_.Exception.Message)
            break
        }
        $timeSeries = Get-OptionalPropertyValue -Object $resp -Name 'timeSeries'
        if ($timeSeries) { $all.AddRange([object[]]@($timeSeries)) }
        $pageToken = [string](Get-OptionalPropertyValue -Object $resp -Name 'nextPageToken')
    } while ($pageToken)

    return $all.ToArray()
}

function Get-TimeSeriesValue {
    param($Series)
    $points = Get-OptionalPropertyValue -Object $Series -Name 'points' -Default @()
    if (-not $Series -or @($points).Count -eq 0) { return [decimal]0 }
    $point = $points | Sort-Object { $_.interval.endTime } -Descending | Select-Object -First 1
    $valueObj = Get-OptionalPropertyValue -Object $point -Name 'value'
    if (-not $point -or -not $valueObj) { return [decimal]0 }

    $int64Value = Get-OptionalPropertyValue -Object $valueObj -Name 'int64Value'
    if ($int64Value) { return (Convert-ToDecimal $int64Value) }
    $doubleValue = Get-OptionalPropertyValue -Object $valueObj -Name 'doubleValue'
    if ($doubleValue) { return (Convert-ToDecimal $doubleValue) }
    return [decimal]0
}

# -------------------------
# Project discovery
# -------------------------
function Get-TargetProjects {
    $resolved = @()

    if ($Projects -and $Projects.Count -gt 0) {
        Write-Log -Level INFO -Message ("Using user supplied project list ({0} projects)." -f $Projects.Count)
        $resolved = @($Projects)
    } else {
        Write-Log -Level INFO -Message 'Discovering accessible projects via Cloud Resource Manager API...'
        $baseUri = 'https://cloudresourcemanager.googleapis.com/v1/projects?pageSize=200'
        $all = New-Object System.Collections.Generic.List[string]
        $pageToken = $null
        do {
            $uri = if ($pageToken) { "$baseUri&pageToken=$([uri]::EscapeDataString($pageToken))" } else { $baseUri }
            $resp = Invoke-GcpApi -Uri $uri
            foreach ($p in @((Get-OptionalPropertyValue -Object $resp -Name 'projects' -Default @()))) {
                if ($p.lifecycleState -eq 'ACTIVE') { [void]$all.Add([string]$p.projectId) }
            }
            $pageToken = [string](Get-OptionalPropertyValue -Object $resp -Name 'nextPageToken')
        } while ($pageToken)

        $resolved = @($all.ToArray())
    }

    $resolved = @($resolved | ForEach-Object { [string]$_ } | ForEach-Object { $_.Trim() } | Where-Object { $_ } | Select-Object -Unique)

    if ($SkipProjects -and $SkipProjects.Count -gt 0) {
        $skipSet = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
        foreach ($sp in @($SkipProjects)) {
            if ($sp) { [void]$skipSet.Add(([string]$sp).Trim()) }
        }

        if ($skipSet.Count -gt 0) {
            $beforeCount = $resolved.Count
            $resolved = @($resolved | Where-Object { -not $skipSet.Contains([string]$_) })
            $skippedCount = $beforeCount - $resolved.Count
            if ($skippedCount -gt 0) {
                Write-Log -Level INFO -Message ("Applied skip-project filter. Skipped={0} | Remaining={1}" -f $skippedCount,$resolved.Count)
            } else {
                Write-Log -Level INFO -Message ("Skip-project filter provided but no projects matched. SkipEntries={0}" -f $skipSet.Count)
            }
        }
    }

    return $resolved
}

# -------------------------
# Inventory collectors (API)
# -------------------------
function Get-ProjectVmInventoryApi {
    param([Parameter(Mandatory)][string]$ProjectId)

    $vmRows = New-Object System.Collections.Generic.List[object]
    $attachedRows = New-Object System.Collections.Generic.List[object]
    $unattachedRows = New-Object System.Collections.Generic.List[object]

    # instances
    $instanceMap = @{}
    $instBase = "https://compute.googleapis.com/compute/v1/projects/$ProjectId/aggregated/instances?maxResults=500"
    $instPageToken = $null
    do {
        $uri = if ($instPageToken) { "$instBase&pageToken=$([uri]::EscapeDataString($instPageToken))" } else { $instBase }
        $resp = Invoke-GcpApi -Uri $uri
        $items = Get-OptionalPropertyValue -Object $resp -Name 'items'
        if ($items) {
            foreach ($prop in $items.PSObject.Properties) {
                $bucket = $prop.Value
                $instances = Get-OptionalPropertyValue -Object $bucket -Name 'instances' -Default @()
                if (@($instances).Count -gt 0) {
                    foreach ($i in @($instances)) {
                        $zone = if ($i.zone) { ([string]$i.zone -split '/')[-1] } else { '' }
                        $region = if ($zone -match '^(.*)-[a-z]$') { $Matches[1] } else { 'Unknown' }
                        $osName = ''
                        try {
                            if ($i.disks -and $i.disks.Count -gt 0 -and $i.disks[0].licenses -and $i.disks[0].licenses.Count -gt 0) {
                                $osName = ([string]$i.disks[0].licenses[0] -split '/')[-1]
                            }
                        } catch {}
                        $diskCount = if ($i.disks) { @($i.disks).Count } else { 0 }
                        $vmRows.Add([PSCustomObject]@{
                            Project = $ProjectId
                            VMName = [string]$i.name
                            VMSize = [string]$i.machineType.Split('/')[-1]
                            OS = $osName
                            Region = $region
                            Zone = $zone
                            VMId = [string]$i.id
                            DiskCount = $diskCount
                            VMDiskSizeGB = [decimal]0
                        }) | Out-Null
                        $instanceMap[[string]$i.selfLink] = $i
                    }
                }
            }
        }
        $instPageToken = [string](Get-OptionalPropertyValue -Object $resp -Name 'nextPageToken')
    } while ($instPageToken)

    # disks
    $diskBySelf = @{}
    $diskBase = "https://compute.googleapis.com/compute/v1/projects/$ProjectId/aggregated/disks?maxResults=500"
    $diskToken = $null
    do {
        $uri = if ($diskToken) { "$diskBase&pageToken=$([uri]::EscapeDataString($diskToken))" } else { $diskBase }
        $resp = Invoke-GcpApi -Uri $uri
        $items = Get-OptionalPropertyValue -Object $resp -Name 'items'
        if ($items) {
            foreach ($prop in $items.PSObject.Properties) {
                $bucket = $prop.Value
                $disks = Get-OptionalPropertyValue -Object $bucket -Name 'disks' -Default @()
                if (@($disks).Count -gt 0) {
                    foreach ($d in @($disks)) {
                        $diskRegion = Get-OptionalPropertyValue -Object $d -Name 'region'
                        $diskZone = Get-OptionalPropertyValue -Object $d -Name 'zone'
                        $region = if ($diskRegion) { ([string]$diskRegion -split '/')[-1] } elseif ($diskZone) { (([string]$diskZone -split '/')[-1] -replace '-[a-z]$','') } else { 'Unknown' }
                        $zone = if ($diskRegion) { '' } elseif ($diskZone) { ([string]$diskZone -split '/')[-1] } else { '' }
                        $diskSelf = [string]$d.selfLink
                        $users = @(Get-OptionalPropertyValue -Object $d -Name 'users' -Default @())
                        $vmNames = @($users | ForEach-Object { ([string]$_ -split '/')[-1] })
                        $hasDiskEncKey = (($d.PSObject.Properties.Name -contains 'diskEncryptionKey') -and $null -ne $d.diskEncryptionKey)
                        $hasEncKey = (($d.PSObject.Properties.Name -contains 'encryptionKey') -and $null -ne $d.encryptionKey)
                        $diskRow = [PSCustomObject]@{
                            DiskName = [string]$d.name
                            VMName = ($vmNames -join '|')
                            Project = $ProjectId
                            Region = $region
                            Zone = $zone
                            IsRegional = [bool]($null -ne $diskRegion)
                            Encrypted = if ($hasDiskEncKey -or $hasEncKey) { 'Yes' } else { 'No' }
                            DiskType = if ($d.type) { ([string]$d.type -split '/')[-1] } else { '' }
                            SizeGB = [int64](Convert-ToDecimal $d.sizeGb)
                            DiskSelfLink = $diskSelf
                        }
                        $diskBySelf[$diskSelf] = $diskRow
                        if ($users.Count -gt 0) {
                            $attachedRows.Add($diskRow) | Out-Null
                        } else {
                            $unattachedRows.Add($diskRow) | Out-Null
                        }
                    }
                }
            }
        }
        $diskToken = [string](Get-OptionalPropertyValue -Object $resp -Name 'nextPageToken')
    } while ($diskToken)

    # Compute VM disk totals from instance attached disk sources
    foreach ($v in $vmRows) {
        $vmSelf = $null
        foreach ($kv in $instanceMap.GetEnumerator()) {
            if ($kv.Value.name -eq $v.VMName) { $vmSelf = $kv.Key; break }
        }
        if (-not $vmSelf) { continue }
        $inst = $instanceMap[$vmSelf]
        $totalGb = [decimal]0
        $cnt = 0
        foreach ($idisk in @($inst.disks)) {
            if (-not $idisk.source) { continue }
            $src = [string]$idisk.source
            if ($diskBySelf.ContainsKey($src)) {
                $cnt++
                $totalGb += Convert-ToDecimal $diskBySelf[$src].SizeGB
            }
        }
        $v.DiskCount = $cnt
        $v.VMDiskSizeGB = Round-Decimal -Value $totalGb -Scale 6
    }

    return [PSCustomObject]@{
        VMs = $vmRows.ToArray()
        AttachedDisks = $attachedRows.ToArray()
        UnattachedDisks = $unattachedRows.ToArray()
    }
}

function Get-GcsBucketSizeMapMonitoring {
    param([Parameter(Mandatory)][string]$ProjectId)

    $map = @{}
    $series = Get-MonitoringSeries -ProjectId $ProjectId -Filter 'metric.type="storage.googleapis.com/storage/total_bytes" AND resource.type="gcs_bucket"' -LookbackHours $MetricLookbackHours
    foreach ($s in @($series)) {
        $bucket = $null
        try { $bucket = [string]$s.resource.labels.bucket_name } catch {}
        if (-not $bucket) { continue }
        $v = Get-TimeSeriesValue -Series $s
        if (-not $map.ContainsKey($bucket)) { $map[$bucket] = [decimal]0 }
        $map[$bucket] = [decimal]$map[$bucket] + $v
    }
    return $map
}

function Get-GcsBucketObjectCountMapMonitoring {
    param([Parameter(Mandatory)][string]$ProjectId)

    $map = @{}
    $series = Get-MonitoringSeries -ProjectId $ProjectId -Filter 'metric.type="storage.googleapis.com/storage/object_count" AND resource.type="gcs_bucket"' -LookbackHours $MetricLookbackHours
    foreach ($s in @($series)) {
        $bucket = $null
        try { $bucket = [string]$s.resource.labels.bucket_name } catch {}
        if (-not $bucket) { continue }

        $v = Get-TimeSeriesValue -Series $s
        if (-not $map.ContainsKey($bucket)) { $map[$bucket] = [decimal]0 }
        if ($v -gt $map[$bucket]) { $map[$bucket] = $v }
    }
    return $map
}

function Get-ProjectStorageInventoryApi {
    param([Parameter(Mandatory)][string]$ProjectId)

    $sizeMap = Get-GcsBucketSizeMapMonitoring -ProjectId $ProjectId
    $objectCountMap = Get-GcsBucketObjectCountMapMonitoring -ProjectId $ProjectId
    $rows = New-Object System.Collections.Generic.List[object]

    $base = "https://storage.googleapis.com/storage/v1/b?project=$([uri]::EscapeDataString($ProjectId))&maxResults=1000"
    $token = $null
    do {
        $uri = if ($token) { "$base&pageToken=$([uri]::EscapeDataString($token))" } else { $base }
        $resp = Invoke-GcpApi -Uri $uri
        foreach ($b in @((Get-OptionalPropertyValue -Object $resp -Name 'items' -Default @()))) {
            $name = [string]$b.name
            $bytes = if ($sizeMap.ContainsKey($name)) { [decimal]$sizeMap[$name] } else { [decimal]0 }
            $objectCount = if ($objectCountMap.ContainsKey($name)) { [decimal]$objectCountMap[$name] } else { [decimal]0 }
            $location = [string]$b.location
            $region = $location
            $zone = ''
            if ($location -match '^[a-z]+-[a-z0-9]+[0-9]+-[a-z]$') {
                $zone = $location
                $region = $location.Substring(0,$location.LastIndexOf('-'))
            }
            $rows.Add([PSCustomObject]@{
                StorageBucket = $name
                Project = $ProjectId
                Location = $location
                Region = $region
                Zone = $zone
                StorageClass = [string]$b.storageClass
                ObjectCount = (Format-WholeNumber -Value $objectCount)
                UsedCapacityBytes = $bytes
                UsedCapacityMiB = Round-Decimal -Value ($bytes/[decimal]1048576) -Scale 6
                UsedCapacityGiB = Round-Decimal -Value ($bytes/[decimal]1073741824) -Scale 6
                UsedCapacityTiB = Round-Decimal -Value ($bytes/[decimal]1099511627776) -Scale 6
                UsedCapacityMBDec = Round-Decimal -Value ($bytes/[decimal]1000000) -Scale 6
                UsedCapacityGB = Round-Decimal -Value ($bytes/[decimal]1000000000) -Scale 6
                UsedCapacityTB = Round-Decimal -Value ($bytes/[decimal]1000000000000) -Scale 6
            }) | Out-Null
        }
        $token = [string](Get-OptionalPropertyValue -Object $resp -Name 'nextPageToken')
    } while ($token)

    return $rows.ToArray()
}

function Get-ProjectFileShareInventoryApi {
    param([Parameter(Mandatory)][string]$ProjectId)

    $usedBytesMap = @{}
    try {
        $series = @()
        $metricFilters = @(
            'metric.type="file.googleapis.com/instance/storage/bytes_used"',
            'metric.type="file.googleapis.com/instance/storage/used_bytes"',
            'metric.type="file.googleapis.com/storage/bytes_used"'
        )
        foreach ($filter in $metricFilters) {
            $series = @(Get-MonitoringSeries -ProjectId $ProjectId -Filter $filter -LookbackHours $MetricLookbackHours)
            if ($series.Count -gt 0) { break }
        }
        foreach ($s in @($series)) {
            $instKey = ''
            try { $instKey = [string]$s.resource.labels.instance_id } catch {}
            if (-not $instKey) { try { $instKey = [string]$s.resource.labels.instance_name } catch {} }
            if (-not $instKey) { try { $instKey = [string]$s.metric.labels.instance_id } catch {} }
            if (-not $instKey) { try { $instKey = [string]$s.metric.labels.instance_name } catch {} }
            if (-not $instKey) { continue }

            $v = Get-TimeSeriesValue -Series $s
            $keys = @($instKey)
            if ($instKey.Contains('/')) { $keys += (($instKey -split '/')[-1]) }
            foreach ($k in $keys) {
                if (-not $k) { continue }
                if (-not $usedBytesMap.ContainsKey($k)) { $usedBytesMap[$k] = [decimal]0 }
                if ($v -gt $usedBytesMap[$k]) { $usedBytesMap[$k] = $v }
            }
        }
    } catch {
        Write-Log -Level WARN -Message ("[FS] Project={0} used-bytes metric lookup failed: {1}" -f $ProjectId,$_.Exception.Message)
    }

    $rows = New-Object System.Collections.Generic.List[object]
    $base = "https://file.googleapis.com/v1/projects/$ProjectId/locations/-/instances?pageSize=1000"
    $token = $null
    do {
        $uri = if ($token) { "$base&pageToken=$([uri]::EscapeDataString($token))" } else { $base }
        $resp = $null
        try { $resp = Invoke-GcpApi -Uri $uri } catch {
            Write-Log -Level WARN -Message ("[FS] Project={0} list failed: {1}" -f $ProjectId,$_.Exception.Message)
            break
        }

        foreach ($inst in @((Get-OptionalPropertyValue -Object $resp -Name 'instances' -Default @()))) {
            $fullInstName = [string](Get-OptionalPropertyValue -Object $inst -Name 'name')
            $instName = ($fullInstName -split '/')[-1]
            $loc = [string](Get-OptionalPropertyValue -Object $inst -Name 'location')
            if (-not $loc -and $fullInstName -match '/locations/([^/]+)/') {
                $loc = [string]$Matches[1]
            }
            $region = 'Unknown'
            $zone = ''
            if ($loc) {
                if ($loc -match '^[a-z]+-[a-z0-9]+[0-9]+-[a-z]$') {
                    $zone = $loc
                    $region = $loc.Substring(0,$loc.LastIndexOf('-'))
                }
                elseif ($loc -match '^[a-z]+-[a-z0-9]+[0-9]+$') {
                    $region = $loc
                }
                else {
                    $region = $loc
                }
            }
            $instanceUsedBytes = if ($usedBytesMap.ContainsKey($instName)) { [decimal]$usedBytesMap[$instName] } else { [decimal]0 }
            $shares = @((Get-OptionalPropertyValue -Object $inst -Name 'fileShares' -Default @()))
            $networks = @((Get-OptionalPropertyValue -Object $inst -Name 'networks' -Default @()))
            $labelsText = ''
            $labelsObj = Get-OptionalPropertyValue -Object $inst -Name 'labels'
            if ($labelsObj -and $labelsObj.PSObject -and $labelsObj.PSObject.Properties) {
                $labelsText = ((@($labelsObj.PSObject.Properties) | ForEach-Object { "{0}={1}" -f $_.Name,$_.Value }) -join ';')
            }

            foreach ($share in $shares) {
                $gib = Convert-ToDecimal (Get-OptionalPropertyValue -Object $share -Name 'capacityGb')
                $gb = Round-Decimal -Value (($gib * [decimal]1073741824)/[decimal]1000000000) -Scale 6
                $tb = Convert-GBToDecimalTB -GB $gb -Scale 6
                $tib = Round-Decimal -Value ($gib/[decimal]1024) -Scale 6
                $shareUsedBytes = Convert-ToDecimal (Get-OptionalPropertyValue -Object $share -Name 'usedBytes')
                if ($shareUsedBytes -le 0) { $shareUsedBytes = $instanceUsedBytes }
                $usedGb = Convert-BytesToDecimalGB -Bytes $shareUsedBytes -Scale 6
                $rows.Add([PSCustomObject]@{
                    Project = $ProjectId
                    InstanceName = $instName
                    ShareName = [string](Get-OptionalPropertyValue -Object $share -Name 'name')
                    Tier = [string](Get-OptionalPropertyValue -Object $inst -Name 'tier')
                    Region = $region
                    Zone = $zone
                    provisionedGb = $gb
                    provisionedGib = $gib
                    provisionedTb = $tb
                    provisionedTib = $tib
                    CapacityGB = $gib
                    UsedCapacityGB = $usedGb
                    Networks = (($networks | ForEach-Object { [string](Get-OptionalPropertyValue -Object $_ -Name 'network') }) -join '|')
                    IPAddresses = (($networks | ForEach-Object { @((Get-OptionalPropertyValue -Object $_ -Name 'ipAddresses' -Default @())) }) -join '|')
                    State = [string](Get-OptionalPropertyValue -Object $inst -Name 'state')
                    CreateTime = [string](Get-OptionalPropertyValue -Object $inst -Name 'createTime')
                    Labels = $labelsText
                    Protocol = [string](Get-OptionalPropertyValue -Object $share -Name 'nfsExportOptions')
                    Error = ''
                }) | Out-Null
            }
        }

        $token = [string](Get-OptionalPropertyValue -Object $resp -Name 'nextPageToken')
    } while ($token)

    return $rows.ToArray()
}

function Get-SpannerStorageMapMonitoring {
    param([Parameter(Mandatory)][string]$ProjectId)

    $map = @{}
    $series = Get-MonitoringSeries -ProjectId $ProjectId -Filter 'metric.type="spanner.googleapis.com/instance/storage/used_bytes"' -LookbackHours $MetricLookbackHours
    foreach ($s in @($series)) {
        $inst = $null
        try { $inst = [string]$s.resource.labels.instance_id } catch {}
        if (-not $inst) { continue }
        $v = Get-TimeSeriesValue -Series $s
        if (-not $map.ContainsKey($inst)) { $map[$inst] = [decimal]0 }
        if ($v -gt $map[$inst]) { $map[$inst] = $v }
    }
    return $map
}

function Get-BigQueryStorageMapMonitoring {
    param([Parameter(Mandatory)][string]$ProjectId)

    $map = @{}
    # GCP metric descriptor documents stored_bytes as logical bytes for BigQuery dataset storage.
    $series = Get-MonitoringSeries -ProjectId $ProjectId -Filter 'metric.type="bigquery.googleapis.com/storage/stored_bytes"' -LookbackHours $MetricLookbackHours
    foreach ($s in @($series)) {
        $dataset = $null
        try { $dataset = [string]$s.resource.labels.dataset_id } catch {}
        if (-not $dataset) { continue }
        $v = Get-TimeSeriesValue -Series $s
        if (-not $map.ContainsKey($dataset)) { $map[$dataset] = [decimal]0 }
        if ($v -gt $map[$dataset]) { $map[$dataset] = $v }
    }
    return $map
}

function Get-BigQueryDatasetTableCount {
    param(
        [Parameter(Mandatory)][string]$ProjectId,
        [Parameter(Mandatory)][string]$DatasetId
    )

    try {
        $metaUri = "https://bigquery.googleapis.com/bigquery/v2/projects/$ProjectId/datasets/$DatasetId"
        $meta = Invoke-GcpApi -Uri $metaUri
        $numTables = Get-OptionalPropertyValue -Object $meta -Name 'numTables'
        if ($null -ne $numTables -and [string]$numTables -ne '') {
            return [int](Convert-ToDecimal $numTables)
        }
    } catch {
        Write-Log -Level DEBUG -Message ("[DB][BigQuery] dataset metadata failed for {0}:{1}: {2}" -f $ProjectId,$DatasetId,$_.Exception.Message)
    }

    $count = 0
    try {
        $base = "https://bigquery.googleapis.com/bigquery/v2/projects/$ProjectId/datasets/$DatasetId/tables?maxResults=1000"
        $token = $null
        do {
            $uri = if ($token) { "$base&pageToken=$([uri]::EscapeDataString($token))" } else { $base }
            $resp = Invoke-GcpApi -Uri $uri
            $count += @((Get-OptionalPropertyValue -Object $resp -Name 'tables' -Default @())).Count
            $token = [string](Get-OptionalPropertyValue -Object $resp -Name 'nextPageToken')
        } while ($token)
    } catch {
        Write-Log -Level WARN -Message ("[DB][BigQuery] table count failed for {0}:{1}: {2}" -f $ProjectId,$DatasetId,$_.Exception.Message)
    }

    return [int]$count
}

function Get-BigQueryDatasetTableCountMap {
    param(
        [Parameter(Mandatory)][string]$ProjectId,
        [Parameter(Mandatory)][string[]]$DatasetIds,
        [int]$Concurrency = 1
    )

    $map = @{}
    $ids = @($DatasetIds | Where-Object { $_ } | Select-Object -Unique)
    if ($ids.Count -eq 0) { return $map }

    $canParallel = ($Concurrency -gt 1 -and $PSVersionTable.PSVersion.Major -ge 7)
    if ($canParallel) {
        try {
            $token = [string]$script:GcpAccessToken
            $timeoutSec = [int]$ApiTimeoutSec
            $proj = [string]$ProjectId
            $results = @($ids | ForEach-Object -ThrottleLimit $Concurrency -Parallel {
                $datasetId = [string]$_
                $headers = @{ Authorization = ("Bearer {0}" -f $using:token) }
                $count = 0
                $err = ''
                try {
                    $datasetIdEsc = [uri]::EscapeDataString($datasetId)
                    $metaUri = "https://bigquery.googleapis.com/bigquery/v2/projects/$($using:proj)/datasets/$datasetIdEsc"
                    $meta = Invoke-RestMethod -Method Get -Uri $metaUri -Headers $headers -TimeoutSec $using:timeoutSec
                    $numTables = $null
                    try { $numTables = $meta.numTables } catch {}
                    if ($null -ne $numTables -and [string]$numTables -ne '') {
                        $count = [int]$numTables
                    } else {
                        $base = "https://bigquery.googleapis.com/bigquery/v2/projects/$($using:proj)/datasets/$datasetIdEsc/tables?maxResults=1000"
                        $tokenLocal = $null
                        do {
                            $uri = if ($tokenLocal) { "$base&pageToken=$([uri]::EscapeDataString([string]$tokenLocal))" } else { $base }
                            $resp = Invoke-RestMethod -Method Get -Uri $uri -Headers $headers -TimeoutSec $using:timeoutSec
                            if ($resp.tables) { $count += @($resp.tables).Count }
                            $tokenLocal = $null
                            try { $tokenLocal = [string]$resp.nextPageToken } catch {}
                        } while ($tokenLocal)
                    }
                } catch {
                    $err = $_.Exception.Message
                }

                [PSCustomObject]@{
                    DatasetId = $datasetId
                    TableCount = $count
                    Error = $err
                }
            })

            $errorCount = 0
            foreach ($r in $results) {
                if (-not $r) { continue }
                $map[[string]$r.DatasetId] = [int](Convert-ToDecimal $r.TableCount)
                if ([string]$r.Error) { $errorCount++ }
            }
            if ($errorCount -gt 0) {
                Write-Log -Level WARN -Message ("[DB][BigQuery] Project={0} table-count lookup had {1} dataset errors; missing values defaulted to 0." -f $ProjectId,$errorCount)
            }
            return $map
        } catch {
            Write-Log -Level WARN -Message ("[DB][BigQuery] Parallel table-count lookup fallback to sequential for Project={0}: {1}" -f $ProjectId,$_.Exception.Message)
        }
    }

    foreach ($datasetId in $ids) {
        $map[[string]$datasetId] = Get-BigQueryDatasetTableCount -ProjectId $ProjectId -DatasetId ([string]$datasetId)
    }
    return $map
}

function Get-AlloyClusterStorageBytes {
    param(
        [Parameter(Mandatory)][string]$ProjectId,
        [Parameter(Mandatory)][string]$ClusterId,
        [Parameter(Mandatory)][string]$Region
    )
    $metricType = 'alloydb.googleapis.com/cluster/storage/usage'
    $filter = 'metric.type="{0}" AND resource.label.cluster_id="{1}" AND resource.label.location="{2}"' -f $metricType,$ClusterId,$Region
    try {
        $series = Get-MonitoringSeries -ProjectId $ProjectId -Filter $filter -LookbackHours $MetricLookbackHours
        if ($series -and @($series).Count -gt 0) {
            $v = Get-TimeSeriesValue -Series $series[0]
            if ($v -gt 0) { return $v }
        }
    } catch {
        Write-Log -Level DEBUG -Message ("AlloyDB metrics lookup failed for cluster=$ClusterId region=$Region : {0}" -f $_.Exception.Message)
    }
    return [decimal]0
}

function Get-ProjectDatabaseInventoryApi {
    param([Parameter(Mandatory)][string]$ProjectId)

    $rows = New-Object System.Collections.Generic.List[object]

    # Cloud SQL
    try {
        $sqlUri = "https://sqladmin.googleapis.com/sql/v1beta4/projects/$ProjectId/instances"
        $sqlResp = Invoke-GcpApi -Uri $sqlUri
        foreach ($i in @((Get-OptionalPropertyValue -Object $sqlResp -Name 'items' -Default @()))) {
            $settings = Get-OptionalPropertyValue -Object $i -Name 'settings'
            $diskGb = Convert-ToDecimal (Get-OptionalPropertyValue -Object $settings -Name 'dataDiskSizeGb')
            $storageType = [string](Get-OptionalPropertyValue -Object $settings -Name 'dataDiskType')
            $pubIps = @($i.ipAddresses | Where-Object { $_.type -eq 'PRIMARY' } | ForEach-Object { $_.ipAddress })
            $privIps = @($i.ipAddresses | Where-Object { $_.type -eq 'PRIVATE' } | ForEach-Object { $_.ipAddress })
            $rows.Add([PSCustomObject]@{
                Type='CloudSQL'
                Project=$ProjectId
                Name=[string]$i.name
                Region=[string]$i.region
                Zone=[string](Get-OptionalPropertyValue -Object $i -Name 'gceZone')
                Engine=[string]$i.databaseVersion
                TierOrCapacity=[string](Get-OptionalPropertyValue -Object $settings -Name 'tier')
                StorageType=$storageType
                NodeCount=''
                TableCount=''
                State=[string]$i.state
                StorageMB=(Convert-GBToDecimalMB -GB $diskGb -Scale 6)
                StorageGB=(Round-Decimal -Value $diskGb -Scale 6)
                Extra=("PrivateIPs={0};PublicIPs={1}" -f ($privIps -join '|'),($pubIps -join '|'))
                Error=''
            }) | Out-Null
        }
    } catch {
        Write-Log -Level WARN -Message ("[DB][CloudSQL] Project={0} failed: {1}" -f $ProjectId,$_.Exception.Message)
    }

    # Spanner
    $spannerMap = Get-SpannerStorageMapMonitoring -ProjectId $ProjectId
    try {
        $spBase = "https://spanner.googleapis.com/v1/projects/$ProjectId/instances?pageSize=1000"
        $spToken = $null
        do {
            $uri = if ($spToken) { "$spBase&pageToken=$([uri]::EscapeDataString($spToken))" } else { $spBase }
            $spResp = Invoke-GcpApi -Uri $uri
            foreach ($s in @((Get-OptionalPropertyValue -Object $spResp -Name 'instances' -Default @()))) {
                $instName = ([string]$s.name -split '/')[-1]
                $bytes = if ($spannerMap.ContainsKey($instName)) { $spannerMap[$instName] } else { [decimal]0 }
                $spConfig = [string]$s.config
                $spRegion = $spConfig
                if ($spConfig -match 'instanceConfigs/(.+)$') { $spRegion = $Matches[1] }
                if ($spRegion -match '^regional-(.+)$') { $spRegion = $Matches[1] }
                $processingUnits = Convert-ToDecimal (Get-OptionalPropertyValue -Object $s -Name 'processingUnits')
                $spNodeCount = Convert-ToDecimal (Get-OptionalPropertyValue -Object $s -Name 'nodeCount')
                if ($spNodeCount -le 0 -and $processingUnits -gt 0) {
                    $spNodeCount = [decimal]([Math]::Ceiling([double]($processingUnits/[decimal]1000)))
                }
                $spNodeCountText = ''
                if ($spNodeCount -gt 0) { $spNodeCountText = Format-WholeNumber -Value $spNodeCount }
                $rows.Add([PSCustomObject]@{
                    Type='SpannerInstance'
                    Project=$ProjectId
                    Name=$instName
                    Region=$spRegion
                    Zone=''
                    Engine='Spanner'
                    TierOrCapacity=[string](Get-OptionalPropertyValue -Object $s -Name 'processingUnits')
                    StorageType=''
                    NodeCount=$spNodeCountText
                    TableCount=''
                    State=[string]$s.state
                    StorageMB=(Convert-BytesToDecimalMB -Bytes $bytes -Scale 6)
                    StorageGB=(Convert-BytesToDecimalGB -Bytes $bytes -Scale 6)
                    Extra='MonitoringMetric=spanner.googleapis.com/instance/storage/used_bytes'
                    Error=''
                }) | Out-Null
            }
            $spToken = [string](Get-OptionalPropertyValue -Object $spResp -Name 'nextPageToken')
        } while ($spToken)
    } catch {
        Write-Log -Level WARN -Message ("[DB][Spanner] Project={0} failed: {1}" -f $ProjectId,$_.Exception.Message)
    }

    # BigQuery datasets (size from monitoring metric when available)
    $bqMap = Get-BigQueryStorageMapMonitoring -ProjectId $ProjectId
    try {
        $bqBase = "https://bigquery.googleapis.com/bigquery/v2/projects/$ProjectId/datasets?maxResults=1000"
        $bqToken = $null
        do {
            $uri = if ($bqToken) { "$bqBase&pageToken=$([uri]::EscapeDataString($bqToken))" } else { $bqBase }
            $bqResp = Invoke-GcpApi -Uri $uri
            foreach ($d in @((Get-OptionalPropertyValue -Object $bqResp -Name 'datasets' -Default @()))) {
                $datasetId = [string]$d.datasetReference.datasetId
                if (-not $datasetId) { continue }
                $bytes = if ($bqMap.ContainsKey($datasetId)) { $bqMap[$datasetId] } else { [decimal]0 }
                $bqRegion = [string](Get-OptionalPropertyValue -Object $d -Name 'location' -Default '')
                $tableCount = Get-BigQueryDatasetTableCount -ProjectId $ProjectId -DatasetId $datasetId
                $rows.Add([PSCustomObject]@{
                    Type='BigQueryDataset'
                    Project=$ProjectId
                    Name=$datasetId
                    Region=$bqRegion
                    Zone=''
                    Engine='BigQuery'
                    TierOrCapacity=''
                    StorageType=''
                    NodeCount=''
                    TableCount=(Format-WholeNumber -Value $tableCount)
                    State=''
                    StorageMB=(Convert-BytesToDecimalMB -Bytes $bytes -Scale 6)
                    StorageGB=(Convert-BytesToDecimalGB -Bytes $bytes -Scale 6)
                    Extra='MonitoringMetric=bigquery.googleapis.com/storage/stored_bytes;BytesType=Logical'
                    Error=''
                }) | Out-Null
            }
            $bqToken = [string](Get-OptionalPropertyValue -Object $bqResp -Name 'nextPageToken')
        } while ($bqToken)
    } catch {
        Write-Log -Level WARN -Message ("[DB][BigQuery] Project={0} failed: {1}" -f $ProjectId,$_.Exception.Message)
    }

    # AlloyDB clusters
    try {
        $alBase = "https://alloydb.googleapis.com/v1/projects/$ProjectId/locations/-/clusters?pageSize=1000"
        $alToken = $null
        do {
            $uri = if ($alToken) { "$alBase&pageToken=$([uri]::EscapeDataString($alToken))" } else { $alBase }
            $alResp = Invoke-GcpApi -Uri $uri
            foreach ($c in @((Get-OptionalPropertyValue -Object $alResp -Name 'clusters' -Default @()))) {
                $clusterName = ([string]$c.name -split '/')[-1]
                $alRegion = ''
                if ([string]$c.name -match 'locations/([^/]+)/') { $alRegion = $Matches[1] }
                $bytes = Get-AlloyClusterStorageBytes -ProjectId $ProjectId -ClusterId $clusterName -Region $alRegion
                $rows.Add([PSCustomObject]@{
                    Type='AlloyDBCluster'
                    Project=$ProjectId
                    Name=$clusterName
                    Region=$alRegion
                    Zone=''
                    Engine='AlloyDB'
                    TierOrCapacity=[string]$c.clusterType
                    StorageType=''
                    NodeCount=''
                    TableCount=''
                    State=[string]$c.state
                    StorageMB=(Convert-BytesToDecimalMB -Bytes $bytes -Scale 6)
                    StorageGB=(Convert-BytesToDecimalGB -Bytes $bytes -Scale 6)
                    Extra='MonitoringMetric=alloydb.googleapis.com/cluster/storage/usage'
                    Error=''
                }) | Out-Null
            }
            $alToken = [string](Get-OptionalPropertyValue -Object $alResp -Name 'nextPageToken')
        } while ($alToken)
    } catch {
        Write-Log -Level WARN -Message ("[DB][AlloyDB] Project={0} failed: {1}" -f $ProjectId,$_.Exception.Message)
    }

    return $rows.ToArray()
}

function Get-GkePvCapacityMapMonitoring {
    param([Parameter(Mandatory)][string]$ProjectId)

    $capMap = @{}
    $countSet = @{}

    $series = Get-MonitoringSeries -ProjectId $ProjectId -Filter 'metric.type="kubernetes.io/persistentvolume/capacity_bytes"' -LookbackHours $MetricLookbackHours
    foreach ($s in @($series)) {
        $cluster = $null
        $pv = $null
        try { $cluster = [string]$s.resource.labels.cluster_name } catch {}
        try { $pv = [string]$s.resource.labels.persistentvolume_name } catch {}
        if (-not $cluster) { continue }

        $val = Get-TimeSeriesValue -Series $s
        if (-not $capMap.ContainsKey($cluster)) { $capMap[$cluster] = [decimal]0 }
        $capMap[$cluster] = [decimal]$capMap[$cluster] + $val

        if (-not $countSet.ContainsKey($cluster)) {
            $countSet[$cluster] = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
        }
        if ($pv) { [void]$countSet[$cluster].Add($pv) }
    }

    $countMap = @{}
    foreach ($k in $countSet.Keys) { $countMap[$k] = $countSet[$k].Count }

    return [PSCustomObject]@{ CapacityBytes = $capMap; CountMap = $countMap }
}

function Get-ProjectGkeInventoryApi {
    param([Parameter(Mandatory)][string]$ProjectId)

    $rows = New-Object System.Collections.Generic.List[object]
    $pvMapObj = Get-GkePvCapacityMapMonitoring -ProjectId $ProjectId

    $base = "https://container.googleapis.com/v1/projects/$ProjectId/locations/-/clusters"
    $resp = $null
    try { $resp = Invoke-GcpApi -Uri $base } catch {
        Write-Log -Level WARN -Message ("[GKE] Project={0} cluster list failed: {1}" -f $ProjectId,$_.Exception.Message)
        return $rows.ToArray()
    }

    foreach ($c in @((Get-OptionalPropertyValue -Object $resp -Name 'clusters' -Default @()))) {
        $clusterName = [string]$c.name
        $location = [string]$c.location
        $region = ''; $zone = ''
        if ($location) {
            if ($location.Split('-').Count -ge 3 -and $location -match '-[A-Za-z]$') { $zone = $location; $region = $location.Substring(0, $location.LastIndexOf('-')) }
            else { $region = $location }
        }

        $nodeCount = [int](Convert-ToDecimal $c.currentNodeCount)

        $estimatedNodeDiskGb = [decimal]0
        foreach ($np in @($c.nodePools)) {
            $diskGb = Convert-ToDecimal $np.config.diskSizeGb
            $poolNodes = Convert-ToDecimal $np.initialNodeCount
            if ($poolNodes -le 0 -and $nodeCount -gt 0) { $poolNodes = $nodeCount }
            $estimatedNodeDiskGb += ($diskGb * $poolNodes)
        }

        $pvBytes = if ($pvMapObj.CapacityBytes.ContainsKey($clusterName)) { $pvMapObj.CapacityBytes[$clusterName] } else { [decimal]0 }
        $pvCount = if ($pvMapObj.CountMap.ContainsKey($clusterName)) { [int]$pvMapObj.CountMap[$clusterName] } else { 0 }

        $rows.Add([PSCustomObject]@{
            Project = $ProjectId
            ClusterName = $clusterName
            Region = $region
            Zone = $zone
            Location = $location
            KubernetesVer = [string]$c.currentMasterVersion
            NodeCount = $nodeCount
            EstimatedvCPU = [decimal]0
            EstimatedMemGB = [decimal]0
            PersistentVolumeCapacityGB = (Convert-BytesToDecimalGB -Bytes $pvBytes -Scale 6)
            PersistentVolumeCount = $pvCount
            PersistentVolumeSource = if ($pvBytes -gt 0) { 'MonitoringAPI' } else { 'NotAvailable' }
            ReleaseChannel = if ($c.releaseChannel) { [string]$c.releaseChannel.channel } else { '' }
            Network = [string]$c.network
            Subnetwork = [string]$c.subnetwork
            Endpoint = [string]$c.endpoint
            EstimatedNodeDiskGB = (Round-Decimal -Value $estimatedNodeDiskGb -Scale 6)
        }) | Out-Null
    }

    return $rows.ToArray()
}

# -------------------------
# Main collection loop
# -------------------------
Write-Log -Level INFO -Message '=== GCP API Resource Inventory Started ==='
Write-Log -Level INFO -Message ("[ExecutionProfile] CPU={0} | RAM_GB={1} | RecommendedMaxThreads={2} | MaxThreads={3} | HardCap={4} | ParallelCapability={5}" -f $script:HostCpuCores,$script:HostMemoryGb,$script:RecommendedMaxThreads,$MaxThreads,$script:ResourceHardCap,$script:ParallelCapability)
if ($script:MaxThreadsClamped) {
    Write-Log -Level WARN -Message ("[ExecutionProfile] Requested MaxThreads exceeded safe host cap; clamped to {0}." -f $MaxThreads)
}
$targetProjects = @(Get-TargetProjects)
if ($targetProjects.Count -eq 0) { throw 'No target projects resolved.' }
Write-Log -Level INFO -Message ("Projects targeted: {0}" -f $targetProjects.Count)

$autoLargeScale = ($targetProjects.Count -ge 200)
if ($autoLargeScale -and -not $PSBoundParameters.ContainsKey('MetricLookbackHours') -and $MetricLookbackHours -lt 48) {
    $MetricLookbackHours = 48
}

$scaleReason = if ($autoLargeScale) { 'ProjectCount>=200' } else { 'NormalScale' }
Write-Log -Level INFO -Message ("[ScaleDetect] AutoLargeScale={0} | Reason={1} | LargeScaleFlag={2} | MetricLookbackHours={3}" -f $(if($autoLargeScale){'Y'}else{'N'}),$scaleReason,$LargeScaleMode,$MetricLookbackHours)

$useWorkerPool = ($EnableProjectWorkerPool -or ((-not $DisableProjectWorkerPool) -and ($LargeScaleMode -or $autoLargeScale)))
if ($useWorkerPool -and $targetProjects.Count -gt 1) {
    Write-Log -Level INFO -Message ("[ScaleDetect] WorkerPool=Enabled | EnableFlag={0} | DisableFlag={1} | ProjectWorkers={2}" -f $EnableProjectWorkerPool,$DisableProjectWorkerPool,$ProjectWorkers)
    $poolResult = Invoke-ProjectWorkerPool -TargetProjects $targetProjects -RootOutputDir $outDir
    if ($poolResult.FailedCount -gt 0) {
        throw ("Worker pool completed with {0} failed workers. Check manifest: {1}" -f $poolResult.FailedCount,$poolResult.ManifestPath)
    }
    Write-Log -Level INFO -Message ("Worker-pool inventory complete. Worker outputs are under: {0}" -f (Join-Path $outDir 'workers'))
    Write-Host "`nInventory complete. Results in $outDir`n" -ForegroundColor Green
    return
}
Write-Log -Level INFO -Message ("[ScaleDetect] WorkerPool=Disabled | EnableFlag={0} | DisableFlag={1} | AutoLargeScale={2}" -f $EnableProjectWorkerPool,$DisableProjectWorkerPool,$autoLargeScale)

$script:RunStart = Get-Date

$vmColumns = @('Project','VMName','VMSize','OS','Region','Zone','VMId','DiskCount','VMDiskSizeGB')
$diskColumns = @('DiskName','VMName','Project','Region','Zone','IsRegional','Encrypted','DiskType','SizeGB','DiskSelfLink')
$storageColumns = @('StorageBucket','Project','Location','Region','Zone','StorageClass','ObjectCount','UsedCapacityBytes','UsedCapacityMiB','UsedCapacityGiB','UsedCapacityTiB','UsedCapacityMBDec','UsedCapacityGB','UsedCapacityTB')
$fileshareColumns = @('Project','InstanceName','ShareName','Tier','Region','Zone','provisionedGb','provisionedGib','provisionedTb','provisionedTib','CapacityGB','UsedCapacityGB','Networks','IPAddresses','State','CreateTime','Labels','Protocol','Error')
$dbColumns = @('Type','Project','Name','Region','Zone','Engine','TierOrCapacity','StorageType','NodeCount','TableCount','State','StorageMB','StorageGB','Extra','Error')
$gkeColumns = @('Project','ClusterName','Region','Zone','Location','KubernetesVer','NodeCount','EstimatedvCPU','EstimatedMemGB','PersistentVolumeCapacityGB','PersistentVolumeCount','PersistentVolumeSource','ReleaseChannel','Network','Subnetwork','Endpoint','EstimatedNodeDiskGB')

$csvPaths = @{}
if ($selected.ContainsKey('VM')) {
    $csvPaths.Vm = Join-Path $outDir ("gcp_vm_instance_info_" + $dateStr + ".csv")
    $csvPaths.Attached = Join-Path $outDir ("gcp_disks_attached_to_vm_instances_" + $dateStr + ".csv")
    $csvPaths.Unattached = Join-Path $outDir ("gcp_disks_unattached_to_vm_instances_" + $dateStr + ".csv")
    Initialize-CsvFile -Path $csvPaths.Vm -Columns $vmColumns
    Initialize-CsvFile -Path $csvPaths.Attached -Columns $diskColumns
    Initialize-CsvFile -Path $csvPaths.Unattached -Columns $diskColumns
}
if ($selected.ContainsKey('STORAGE')) {
    $csvPaths.Storage = Join-Path $outDir ("gcp_storage_buckets_info_" + $dateStr + ".csv")
    Initialize-CsvFile -Path $csvPaths.Storage -Columns $storageColumns
}
if ($selected.ContainsKey('FILESHARE')) {
    $csvPaths.Fileshare = Join-Path $outDir ("gcp_fileshares_info_" + $dateStr + ".csv")
    Initialize-CsvFile -Path $csvPaths.Fileshare -Columns $fileshareColumns
}

$dbWorkloadTypes = @('CloudSQL','SpannerInstance','BigQueryDataset','AlloyDBCluster')
$dbWorkloadDetailFiles = @{}
$dbWorkloadStats = @{}
if ($selected.ContainsKey('DB')) {
    $csvPaths.Db = Join-Path $outDir ("gcp_database_inventory_" + $dateStr + ".csv")
    Initialize-CsvFile -Path $csvPaths.Db -Columns $dbColumns
    foreach ($w in $dbWorkloadTypes) {
        $tmp = Join-Path $outDir ("tmp_" + $w.ToLowerInvariant() + "_" + $dateStr + ".csv")
        Initialize-CsvFile -Path $tmp -Columns $dbColumns
        $dbWorkloadDetailFiles[$w] = $tmp
        $dbWorkloadStats[$w] = New-DbAggregateStats
    }
}

if ($selected.ContainsKey('GKE')) {
    $csvPaths.Gke = Join-Path $outDir ("gcp_gke_clusters_" + $dateStr + ".csv")
    Initialize-CsvFile -Path $csvPaths.Gke -Columns $gkeColumns
}

$stats = @{
    VM = @{
        TotalCount = 0
        ProjectCounts = @{}
        RegionCounts = @{}
        ProjectRegionCounts = @{}
        ProjectZoneCounts = @{}
        ProjectZoneRegion = @{}
        DiskGbTotal = [decimal]0
        DiskGbByProject = @{}
        DiskGbByRegion = @{}
        DiskGbByProjectRegion = @{}
        DiskGbByProjectZone = @{}
    }
    Attached = @{
        TotalCount = 0
        TotalGb = [decimal]0
        ProjectCounts = @{}
        ProjectGb = @{}
        RegionCounts = @{}
        RegionGb = @{}
    }
    Unattached = @{
        TotalCount = 0
        TotalGb = [decimal]0
        ProjectCounts = @{}
        ProjectGb = @{}
        RegionCounts = @{}
        RegionGb = @{}
    }
    Storage = @{
        TotalCount = 0
        TotalBytes = [decimal]0
        TotalObjectCount = [decimal]0
        ProjectCounts = @{}
        ProjectBytes = @{}
        ProjectObjectCount = @{}
        RegionCounts = @{}
        RegionBytes = @{}
        RegionObjectCount = @{}
    }
    Fileshare = @{
        TotalCount = 0
        TotalGb = [decimal]0
        TotalUsedGb = [decimal]0
        ProjectCounts = @{}
        ProjectGb = @{}
        ProjectUsedGb = @{}
        RegionCounts = @{}
        RegionGb = @{}
        RegionUsedGb = @{}
    }
    DB = (New-DbAggregateStats)
    GKE = @{
        TotalCount = 0
        TotalPvGb = [decimal]0
        TotalNodeGb = [decimal]0
        ProjectCounts = @{}
        ProjectPvGb = @{}
        ProjectNodeGb = @{}
        RegionCounts = @{}
        RegionPvGb = @{}
        RegionNodeGb = @{}
    }
}

for ($projIndex = 0; $projIndex -lt $targetProjects.Count; $projIndex++) {
    $proj = [string]$targetProjects[$projIndex]
    $currentProjectNumber = $projIndex + 1
    $completedBefore = $projIndex
    $completedAfter = $projIndex + 1

    Update-RunProgress -CompletedProjects $completedBefore -CurrentProjectNumber $currentProjectNumber -TotalProjects $targetProjects.Count -Project $proj -Stage 'ProjectStart'
    Write-Log -Level INFO -Message ("[Project-Start] {0}" -f $proj)

    if ($selected.ContainsKey('VM')) {
        Update-RunProgress -CompletedProjects $completedBefore -CurrentProjectNumber $currentProjectNumber -TotalProjects $targetProjects.Count -Project $proj -Stage 'VM'
        try {
            $vmRes = Get-ProjectVmInventoryApi -ProjectId $proj
            $vmRows = @($vmRes.VMs)
            $attachedOut = @($vmRes.AttachedDisks | Select-Object DiskName,VMName,Project,Region,Zone,IsRegional,Encrypted,DiskType,SizeGB,DiskSelfLink)
            $unattachedOut = @($vmRes.UnattachedDisks | Select-Object DiskName,VMName,Project,Region,Zone,IsRegional,Encrypted,DiskType,SizeGB,DiskSelfLink)

            Append-CsvRows -Path $csvPaths.Vm -Rows $vmRows -Columns $vmColumns -Context ("VM|{0}" -f $proj)
            Append-CsvRows -Path $csvPaths.Attached -Rows $attachedOut -Columns $diskColumns -Context ("AttachedDisks|{0}" -f $proj)
            Append-CsvRows -Path $csvPaths.Unattached -Rows $unattachedOut -Columns $diskColumns -Context ("UnattachedDisks|{0}" -f $proj)

            foreach ($v in $vmRows) {
                $region = Normalize-Key -Key ([string]$v.Region)
                $zone = Normalize-Key -Key ([string]$v.Zone)
                $stats.VM.TotalCount++
                Add-CountToMap -Map $stats.VM.ProjectCounts -Key $proj -Default 'Unknown'
                Add-CountToMap -Map $stats.VM.RegionCounts -Key $region -Default 'Unknown'
                $prKey = New-Key2 -A $proj -B $region
                Add-CountToMap -Map $stats.VM.ProjectRegionCounts -Key $prKey -Default ''
                $pzKey = New-Key2 -A $proj -B $zone
                Add-CountToMap -Map $stats.VM.ProjectZoneCounts -Key $pzKey -Default ''
                $stats.VM.ProjectZoneRegion[$pzKey] = $region
            }

            foreach ($d in $attachedOut) {
                $sizeGb = Convert-ToDecimal $d.SizeGB
                $region = Normalize-Key -Key ([string]$d.Region)
                $zone = Normalize-Key -Key ([string]$d.Zone)

                $stats.Attached.TotalCount++
                $stats.Attached.TotalGb += $sizeGb
                Add-CountToMap -Map $stats.Attached.ProjectCounts -Key $proj -Default 'Unknown'
                Add-DecimalToMap -Map $stats.Attached.ProjectGb -Key $proj -Increment $sizeGb -Default 'Unknown'
                Add-CountToMap -Map $stats.Attached.RegionCounts -Key $region -Default 'Unknown'
                Add-DecimalToMap -Map $stats.Attached.RegionGb -Key $region -Increment $sizeGb -Default 'Unknown'

                $stats.VM.DiskGbTotal += $sizeGb
                Add-DecimalToMap -Map $stats.VM.DiskGbByProject -Key $proj -Increment $sizeGb -Default 'Unknown'
                Add-DecimalToMap -Map $stats.VM.DiskGbByRegion -Key $region -Increment $sizeGb -Default 'Unknown'
                Add-DecimalToMap -Map $stats.VM.DiskGbByProjectRegion -Key (New-Key2 -A $proj -B $region) -Increment $sizeGb -Default ''
                Add-DecimalToMap -Map $stats.VM.DiskGbByProjectZone -Key (New-Key2 -A $proj -B $zone) -Increment $sizeGb -Default ''
            }

            foreach ($d in $unattachedOut) {
                $sizeGb = Convert-ToDecimal $d.SizeGB
                $region = Normalize-Key -Key ([string]$d.Region)
                $stats.Unattached.TotalCount++
                $stats.Unattached.TotalGb += $sizeGb
                Add-CountToMap -Map $stats.Unattached.ProjectCounts -Key $proj -Default 'Unknown'
                Add-DecimalToMap -Map $stats.Unattached.ProjectGb -Key $proj -Increment $sizeGb -Default 'Unknown'
                Add-CountToMap -Map $stats.Unattached.RegionCounts -Key $region -Default 'Unknown'
                Add-DecimalToMap -Map $stats.Unattached.RegionGb -Key $region -Increment $sizeGb -Default 'Unknown'
            }

            Write-Log -Level INFO -Message ("[VM] Project={0} VMs={1} Attached={2} Unattached={3}" -f $proj,$vmRows.Count,$attachedOut.Count,$unattachedOut.Count)
        } catch {
            Write-Log -Level WARN -Message ("[VM] Project={0} failed: {1}" -f $proj,$_.Exception.Message)
        }
    }

    if ($selected.ContainsKey('STORAGE')) {
        Update-RunProgress -CompletedProjects $completedBefore -CurrentProjectNumber $currentProjectNumber -TotalProjects $targetProjects.Count -Project $proj -Stage 'Storage'
        try {
            $bRows = @(Get-ProjectStorageInventoryApi -ProjectId $proj)
            Append-CsvRows -Path $csvPaths.Storage -Rows $bRows -Columns $storageColumns -Context ("Storage|{0}" -f $proj)

            foreach ($b in $bRows) {
                $bytes = Convert-ToDecimal $b.UsedCapacityBytes
                $objectCount = Convert-ToDecimal $b.ObjectCount
                $region = Normalize-Key -Key ([string]$b.Region)
                $stats.Storage.TotalCount++
                $stats.Storage.TotalBytes += $bytes
                $stats.Storage.TotalObjectCount += $objectCount
                Add-CountToMap -Map $stats.Storage.ProjectCounts -Key $proj -Default 'Unknown'
                Add-DecimalToMap -Map $stats.Storage.ProjectBytes -Key $proj -Increment $bytes -Default 'Unknown'
                Add-DecimalToMap -Map $stats.Storage.ProjectObjectCount -Key $proj -Increment $objectCount -Default 'Unknown'
                Add-CountToMap -Map $stats.Storage.RegionCounts -Key $region -Default 'Unknown'
                Add-DecimalToMap -Map $stats.Storage.RegionBytes -Key $region -Increment $bytes -Default 'Unknown'
                Add-DecimalToMap -Map $stats.Storage.RegionObjectCount -Key $region -Increment $objectCount -Default 'Unknown'
            }

            Write-Log -Level INFO -Message ("[Storage] Project={0} Buckets={1}" -f $proj,$bRows.Count)
        } catch {
            Write-Log -Level WARN -Message ("[Storage] Project={0} failed: {1}" -f $proj,$_.Exception.Message)
        }
    }

    if ($selected.ContainsKey('FILESHARE')) {
        Update-RunProgress -CompletedProjects $completedBefore -CurrentProjectNumber $currentProjectNumber -TotalProjects $targetProjects.Count -Project $proj -Stage 'Fileshare'
        try {
            $fRows = @(Get-ProjectFileShareInventoryApi -ProjectId $proj)
            Append-CsvRows -Path $csvPaths.Fileshare -Rows $fRows -Columns $fileshareColumns -Context ("Fileshare|{0}" -f $proj)

            foreach ($f in $fRows) {
                $gb = Convert-ToDecimal $f.provisionedGb
                $usedGb = Convert-ToDecimal $f.UsedCapacityGB
                $region = Normalize-Key -Key ([string]$f.Region)
                $stats.Fileshare.TotalCount++
                $stats.Fileshare.TotalGb += $gb
                $stats.Fileshare.TotalUsedGb += $usedGb
                Add-CountToMap -Map $stats.Fileshare.ProjectCounts -Key $proj -Default 'Unknown'
                Add-DecimalToMap -Map $stats.Fileshare.ProjectGb -Key $proj -Increment $gb -Default 'Unknown'
                Add-DecimalToMap -Map $stats.Fileshare.ProjectUsedGb -Key $proj -Increment $usedGb -Default 'Unknown'
                Add-CountToMap -Map $stats.Fileshare.RegionCounts -Key $region -Default 'Unknown'
                Add-DecimalToMap -Map $stats.Fileshare.RegionGb -Key $region -Increment $gb -Default 'Unknown'
                Add-DecimalToMap -Map $stats.Fileshare.RegionUsedGb -Key $region -Increment $usedGb -Default 'Unknown'
            }

            Write-Log -Level INFO -Message ("[Fileshare] Project={0} Shares={1}" -f $proj,$fRows.Count)
        } catch {
            Write-Log -Level WARN -Message ("[Fileshare] Project={0} failed: {1}" -f $proj,$_.Exception.Message)
        }
    }

    if ($selected.ContainsKey('DB')) {
        Update-RunProgress -CompletedProjects $completedBefore -CurrentProjectNumber $currentProjectNumber -TotalProjects $targetProjects.Count -Project $proj -Stage 'Database'
        try {
            $dRows = @(Get-ProjectDatabaseInventoryApi -ProjectId $proj)
            Append-CsvRows -Path $csvPaths.Db -Rows $dRows -Columns $dbColumns -Context ("Database|{0}" -f $proj)

            foreach ($wg in @($dRows | Group-Object Type)) {
                $wType = [string]$wg.Name
                if ($dbWorkloadDetailFiles.ContainsKey($wType)) {
                    Append-CsvRows -Path $dbWorkloadDetailFiles[$wType] -Rows $wg.Group -Columns $dbColumns -Context ("DBWorkload:{0}|{1}" -f $wType,$proj)
                }
            }

            foreach ($d in $dRows) {
                $mb = Convert-ToDecimal $d.StorageMB
                $gb = Convert-ToDecimal $d.StorageGB
                $region = Normalize-Key -Key ([string]$d.Region)
                $stats.DB.RowCount++
                $stats.DB.TotalMb += $mb
                $stats.DB.TotalGb += $gb
                Add-CountToMap -Map $stats.DB.ProjectCount -Key $proj -Default 'Unknown'
                Add-DecimalToMap -Map $stats.DB.ProjectMb -Key $proj -Increment $mb -Default 'Unknown'
                Add-DecimalToMap -Map $stats.DB.ProjectGb -Key $proj -Increment $gb -Default 'Unknown'
                Add-CountToMap -Map $stats.DB.RegionCount -Key $region -Default 'Unknown'
                Add-DecimalToMap -Map $stats.DB.RegionMb -Key $region -Increment $mb -Default 'Unknown'
                Add-DecimalToMap -Map $stats.DB.RegionGb -Key $region -Increment $gb -Default 'Unknown'

                $w = [string]$d.Type
                if ($dbWorkloadStats.ContainsKey($w)) {
                    $ws = $dbWorkloadStats[$w]
                    $ws.RowCount++
                    $ws.TotalMb += $mb
                    $ws.TotalGb += $gb
                    Add-CountToMap -Map $ws.ProjectCount -Key $proj -Default 'Unknown'
                    Add-DecimalToMap -Map $ws.ProjectMb -Key $proj -Increment $mb -Default 'Unknown'
                    Add-DecimalToMap -Map $ws.ProjectGb -Key $proj -Increment $gb -Default 'Unknown'
                    Add-CountToMap -Map $ws.RegionCount -Key $region -Default 'Unknown'
                    Add-DecimalToMap -Map $ws.RegionMb -Key $region -Increment $mb -Default 'Unknown'
                    Add-DecimalToMap -Map $ws.RegionGb -Key $region -Increment $gb -Default 'Unknown'
                }
            }

            Write-Log -Level INFO -Message ("[DB] Project={0} Rows={1}" -f $proj,$dRows.Count)
        } catch {
            Write-Log -Level WARN -Message ("[DB] Project={0} failed: {1}" -f $proj,$_.Exception.Message)
        }
    }

    if ($selected.ContainsKey('GKE')) {
        Update-RunProgress -CompletedProjects $completedBefore -CurrentProjectNumber $currentProjectNumber -TotalProjects $targetProjects.Count -Project $proj -Stage 'GKE'
        try {
            $gRows = @(Get-ProjectGkeInventoryApi -ProjectId $proj)
            Append-CsvRows -Path $csvPaths.Gke -Rows $gRows -Columns $gkeColumns -Context ("GKE|{0}" -f $proj)

            foreach ($g in $gRows) {
                $region = Normalize-Key -Key ([string]$g.Region)
                $pvGb = Convert-ToDecimal $g.PersistentVolumeCapacityGB
                $nodeGb = Convert-ToDecimal $g.EstimatedNodeDiskGB
                $stats.GKE.TotalCount++
                $stats.GKE.TotalPvGb += $pvGb
                $stats.GKE.TotalNodeGb += $nodeGb
                Add-CountToMap -Map $stats.GKE.ProjectCounts -Key $proj -Default 'Unknown'
                Add-DecimalToMap -Map $stats.GKE.ProjectPvGb -Key $proj -Increment $pvGb -Default 'Unknown'
                Add-DecimalToMap -Map $stats.GKE.ProjectNodeGb -Key $proj -Increment $nodeGb -Default 'Unknown'
                Add-CountToMap -Map $stats.GKE.RegionCounts -Key $region -Default 'Unknown'
                Add-DecimalToMap -Map $stats.GKE.RegionPvGb -Key $region -Increment $pvGb -Default 'Unknown'
                Add-DecimalToMap -Map $stats.GKE.RegionNodeGb -Key $region -Increment $nodeGb -Default 'Unknown'
            }

            Write-Log -Level INFO -Message ("[GKE] Project={0} Clusters={1}" -f $proj,$gRows.Count)
        } catch {
            Write-Log -Level WARN -Message ("[GKE] Project={0} failed: {1}" -f $proj,$_.Exception.Message)
        }
    }

    Write-Log -Level INFO -Message ("[Project-End] {0}" -f $proj)
    Update-RunProgress -CompletedProjects $completedAfter -CurrentProjectNumber $currentProjectNumber -TotalProjects $targetProjects.Count -Project $proj -Stage 'ProjectComplete'
}

Write-Progress -Id 1 -Activity 'GCP API Inventory Progress' -Completed

# -------------------------
# CSV output summaries / finalize
# -------------------------
if ($selected.ContainsKey('VM')) {
    $vmTotalGb = Round-Decimal -Value $stats.VM.DiskGbTotal -Scale 6
    Add-ContentSafe -Path $csvPaths.Vm -Value ''
    Add-ContentSafe -Path $csvPaths.Vm -Value '### VM Summary ###'
    Add-ContentSafe -Path $csvPaths.Vm -Value ("TotalVMs,{0},TotalDiskGB,{1},TotalDiskTB,{2}" -f $stats.VM.TotalCount,$vmTotalGb,(Convert-GBToDecimalTB -GB $vmTotalGb -Scale 6))

    Add-ContentSafe -Path $csvPaths.Vm -Value ''
    Add-ContentSafe -Path $csvPaths.Vm -Value '### VM Project Summary ###'
    Add-ContentSafe -Path $csvPaths.Vm -Value 'Project,VMCount,TotalDiskGB,TotalDiskTB'
    foreach ($project in @($stats.VM.ProjectCounts.Keys | Sort-Object)) {
        $pCount = Get-MapCount -Map $stats.VM.ProjectCounts -Key $project
        $pGb = Get-MapDecimal -Map $stats.VM.DiskGbByProject -Key $project
        Add-ContentSafe -Path $csvPaths.Vm -Value ("{0},{1},{2},{3}" -f $project,$pCount,(Round-Decimal -Value $pGb -Scale 6),(Convert-GBToDecimalTB -GB $pGb -Scale 6))
    }

    Add-ContentSafe -Path $csvPaths.Vm -Value ''
    Add-ContentSafe -Path $csvPaths.Vm -Value '### VM Region Summary ###'
    Add-ContentSafe -Path $csvPaths.Vm -Value 'Region,VMCount,TotalDiskGB,TotalDiskTB'
    foreach ($region in @($stats.VM.RegionCounts.Keys | Sort-Object)) {
        $rCount = Get-MapCount -Map $stats.VM.RegionCounts -Key $region
        $rGb = Get-MapDecimal -Map $stats.VM.DiskGbByRegion -Key $region
        Add-ContentSafe -Path $csvPaths.Vm -Value ("{0},{1},{2},{3}" -f $region,$rCount,(Round-Decimal -Value $rGb -Scale 6),(Convert-GBToDecimalTB -GB $rGb -Scale 6))
    }
    Write-Log -Level INFO -Message ("VM CSV written: {0}" -f (Split-Path $csvPaths.Vm -Leaf))

    $attachedTotalGb = Round-Decimal -Value $stats.Attached.TotalGb -Scale 6
    Add-ContentSafe -Path $csvPaths.Attached -Value ''
    Add-ContentSafe -Path $csvPaths.Attached -Value '### Attached Disk Summary ###'
    Add-ContentSafe -Path $csvPaths.Attached -Value ("TotalAttachedDisks,{0},TotalSizeGB,{1},TotalSizeTB,{2}" -f $stats.Attached.TotalCount,$attachedTotalGb,(Convert-GBToDecimalTB -GB $attachedTotalGb -Scale 6))

    Add-ContentSafe -Path $csvPaths.Attached -Value ''
    Add-ContentSafe -Path $csvPaths.Attached -Value '### Attached Disk Project Summary ###'
    Add-ContentSafe -Path $csvPaths.Attached -Value 'Project,DiskCount,TotalSizeGB,TotalSizeTB'
    foreach ($project in @($stats.Attached.ProjectCounts.Keys | Sort-Object)) {
        $pCount = Get-MapCount -Map $stats.Attached.ProjectCounts -Key $project
        $pGb = Get-MapDecimal -Map $stats.Attached.ProjectGb -Key $project
        Add-ContentSafe -Path $csvPaths.Attached -Value ("{0},{1},{2},{3}" -f $project,$pCount,(Round-Decimal -Value $pGb -Scale 6),(Convert-GBToDecimalTB -GB $pGb -Scale 6))
    }

    Add-ContentSafe -Path $csvPaths.Attached -Value ''
    Add-ContentSafe -Path $csvPaths.Attached -Value '### Attached Disk Region Summary ###'
    Add-ContentSafe -Path $csvPaths.Attached -Value 'Region,DiskCount,TotalSizeGB,TotalSizeTB'
    foreach ($region in @($stats.Attached.RegionCounts.Keys | Sort-Object)) {
        $rCount = Get-MapCount -Map $stats.Attached.RegionCounts -Key $region
        $rGb = Get-MapDecimal -Map $stats.Attached.RegionGb -Key $region
        Add-ContentSafe -Path $csvPaths.Attached -Value ("{0},{1},{2},{3}" -f $region,$rCount,(Round-Decimal -Value $rGb -Scale 6),(Convert-GBToDecimalTB -GB $rGb -Scale 6))
    }
    Write-Log -Level INFO -Message ("Attached disk CSV written: {0}" -f (Split-Path $csvPaths.Attached -Leaf))

    $unattachedTotalGb = Round-Decimal -Value $stats.Unattached.TotalGb -Scale 6
    Add-ContentSafe -Path $csvPaths.Unattached -Value ''
    Add-ContentSafe -Path $csvPaths.Unattached -Value '### Unattached Disk Summary ###'
    Add-ContentSafe -Path $csvPaths.Unattached -Value ("TotalUnattachedDisks,{0},TotalSizeGB,{1},TotalSizeTB,{2}" -f $stats.Unattached.TotalCount,$unattachedTotalGb,(Convert-GBToDecimalTB -GB $unattachedTotalGb -Scale 6))

    Add-ContentSafe -Path $csvPaths.Unattached -Value ''
    Add-ContentSafe -Path $csvPaths.Unattached -Value '### Unattached Disk Project Summary ###'
    Add-ContentSafe -Path $csvPaths.Unattached -Value 'Project,DiskCount,TotalSizeGB,TotalSizeTB'
    foreach ($project in @($stats.Unattached.ProjectCounts.Keys | Sort-Object)) {
        $pCount = Get-MapCount -Map $stats.Unattached.ProjectCounts -Key $project
        $pGb = Get-MapDecimal -Map $stats.Unattached.ProjectGb -Key $project
        Add-ContentSafe -Path $csvPaths.Unattached -Value ("{0},{1},{2},{3}" -f $project,$pCount,(Round-Decimal -Value $pGb -Scale 6),(Convert-GBToDecimalTB -GB $pGb -Scale 6))
    }

    Add-ContentSafe -Path $csvPaths.Unattached -Value ''
    Add-ContentSafe -Path $csvPaths.Unattached -Value '### Unattached Disk Region Summary ###'
    Add-ContentSafe -Path $csvPaths.Unattached -Value 'Region,DiskCount,TotalSizeGB,TotalSizeTB'
    foreach ($region in @($stats.Unattached.RegionCounts.Keys | Sort-Object)) {
        $rCount = Get-MapCount -Map $stats.Unattached.RegionCounts -Key $region
        $rGb = Get-MapDecimal -Map $stats.Unattached.RegionGb -Key $region
        Add-ContentSafe -Path $csvPaths.Unattached -Value ("{0},{1},{2},{3}" -f $region,$rCount,(Round-Decimal -Value $rGb -Scale 6),(Convert-GBToDecimalTB -GB $rGb -Scale 6))
    }
    Write-Log -Level INFO -Message ("Unattached disk CSV written: {0}" -f (Split-Path $csvPaths.Unattached -Leaf))
}

if ($selected.ContainsKey('STORAGE')) {
    $totalBytes = Convert-ToDecimal $stats.Storage.TotalBytes
    Add-ContentSafe -Path $csvPaths.Storage -Value ''
    Add-ContentSafe -Path $csvPaths.Storage -Value '### Storage Summary ###'
    Add-ContentSafe -Path $csvPaths.Storage -Value ("TotalBuckets,{0},TotalObjectCount,{1},TotalSizeGB,{2},TotalSizeTB,{3}" -f $stats.Storage.TotalCount,(Format-WholeNumber -Value $stats.Storage.TotalObjectCount),(Convert-BytesToDecimalGB -Bytes $totalBytes -Scale 6),(Convert-BytesToDecimalTB -Bytes $totalBytes -Scale 6))

    Add-ContentSafe -Path $csvPaths.Storage -Value ''
    Add-ContentSafe -Path $csvPaths.Storage -Value '### Storage Project Summary ###'
    Add-ContentSafe -Path $csvPaths.Storage -Value 'Project,BucketCount,ObjectCount,TotalSizeGB,TotalSizeTB'
    foreach ($project in @($stats.Storage.ProjectCounts.Keys | Sort-Object)) {
        $pCount = Get-MapCount -Map $stats.Storage.ProjectCounts -Key $project
        $pBytes = Get-MapDecimal -Map $stats.Storage.ProjectBytes -Key $project
        $pObj = Get-MapDecimal -Map $stats.Storage.ProjectObjectCount -Key $project
        Add-ContentSafe -Path $csvPaths.Storage -Value ("{0},{1},{2},{3},{4}" -f $project,$pCount,(Format-WholeNumber -Value $pObj),(Convert-BytesToDecimalGB -Bytes $pBytes -Scale 6),(Convert-BytesToDecimalTB -Bytes $pBytes -Scale 6))
    }

    Add-ContentSafe -Path $csvPaths.Storage -Value ''
    Add-ContentSafe -Path $csvPaths.Storage -Value '### Storage Region Summary ###'
    Add-ContentSafe -Path $csvPaths.Storage -Value 'Region,BucketCount,ObjectCount,TotalSizeGB,TotalSizeTB'
    foreach ($region in @($stats.Storage.RegionCounts.Keys | Sort-Object)) {
        $rCount = Get-MapCount -Map $stats.Storage.RegionCounts -Key $region
        $rBytes = Get-MapDecimal -Map $stats.Storage.RegionBytes -Key $region
        $rObj = Get-MapDecimal -Map $stats.Storage.RegionObjectCount -Key $region
        Add-ContentSafe -Path $csvPaths.Storage -Value ("{0},{1},{2},{3},{4}" -f $region,$rCount,(Format-WholeNumber -Value $rObj),(Convert-BytesToDecimalGB -Bytes $rBytes -Scale 6),(Convert-BytesToDecimalTB -Bytes $rBytes -Scale 6))
    }
    Write-Log -Level INFO -Message ("Storage CSV written: {0}" -f (Split-Path $csvPaths.Storage -Leaf))
}

if ($selected.ContainsKey('FILESHARE')) {
    $totalFsGb = Round-Decimal -Value $stats.Fileshare.TotalGb -Scale 6
    $totalFsUsedGb = Round-Decimal -Value $stats.Fileshare.TotalUsedGb -Scale 6
    Add-ContentSafe -Path $csvPaths.Fileshare -Value ''
    Add-ContentSafe -Path $csvPaths.Fileshare -Value '### Fileshare Summary ###'
    Add-ContentSafe -Path $csvPaths.Fileshare -Value ("TotalShares,{0},TotalSizeGB,{1},TotalSizeTB,{2},TotalUsedCapacityGB,{3},TotalUsedCapacityTB,{4}" -f $stats.Fileshare.TotalCount,$totalFsGb,(Convert-GBToDecimalTB -GB $totalFsGb -Scale 6),$totalFsUsedGb,(Convert-GBToDecimalTB -GB $totalFsUsedGb -Scale 6))
    Write-Log -Level INFO -Message ("Fileshare CSV written: {0}" -f (Split-Path $csvPaths.Fileshare -Leaf))
}

if ($selected.ContainsKey('DB')) {
    $dbErrors = @((Get-InventoryErrorRecordsFromLog -LogPath $script:LogFile) | Where-Object Workload -eq 'Database')

    Add-ContentSafe -Path $csvPaths.Db -Value ''
    Add-ContentSafe -Path $csvPaths.Db -Value '### Database Summary ###'
    Add-ContentSafe -Path $csvPaths.Db -Value ("TotalRows,{0},TotalSizeMB,{1},TotalSizeGB,{2},TotalSizeTB,{3}" -f $stats.DB.RowCount,(Round-Decimal -Value $stats.DB.TotalMb -Scale 6),(Round-Decimal -Value $stats.DB.TotalGb -Scale 6),(Convert-GBToDecimalTB -GB $stats.DB.TotalGb -Scale 6))

    Add-ContentSafe -Path $csvPaths.Db -Value ''
    Add-ContentSafe -Path $csvPaths.Db -Value '### Database Project Summary ###'
    Add-ContentSafe -Path $csvPaths.Db -Value 'Project,Rows,TotalSizeMB,TotalSizeGB,TotalSizeTB'
    foreach ($project in @($stats.DB.ProjectCount.Keys | Sort-Object)) {
        $pCount = Get-MapCount -Map $stats.DB.ProjectCount -Key $project
        $pMb = Get-MapDecimal -Map $stats.DB.ProjectMb -Key $project
        $pGb = Get-MapDecimal -Map $stats.DB.ProjectGb -Key $project
        Add-ContentSafe -Path $csvPaths.Db -Value ("{0},{1},{2},{3},{4}" -f $project,$pCount,(Round-Decimal -Value $pMb -Scale 6),(Round-Decimal -Value $pGb -Scale 6),(Convert-GBToDecimalTB -GB $pGb -Scale 6))
    }

    Add-ContentSafe -Path $csvPaths.Db -Value ''
    Add-ContentSafe -Path $csvPaths.Db -Value '### Database Region Summary ###'
    Add-ContentSafe -Path $csvPaths.Db -Value 'Region,Rows,TotalSizeMB,TotalSizeGB,TotalSizeTB'
    foreach ($region in @($stats.DB.RegionCount.Keys | Sort-Object)) {
        $rCount = Get-MapCount -Map $stats.DB.RegionCount -Key $region
        $rMb = Get-MapDecimal -Map $stats.DB.RegionMb -Key $region
        $rGb = Get-MapDecimal -Map $stats.DB.RegionGb -Key $region
        Add-ContentSafe -Path $csvPaths.Db -Value ("{0},{1},{2},{3},{4}" -f $region,$rCount,(Round-Decimal -Value $rMb -Scale 6),(Round-Decimal -Value $rGb -Scale 6),(Convert-GBToDecimalTB -GB $rGb -Scale 6))
    }
    Write-Log -Level INFO -Message ("Database CSV written: {0}" -f (Split-Path $csvPaths.Db -Leaf))

    foreach ($w in $dbWorkloadTypes) {
        $safe = $w.ToLowerInvariant()
        $wCsv = Join-Path $outDir ("gcp_" + $safe + "_" + $dateStr + ".csv")
        Write-DbWorkloadCsvWithSummary -Path $wCsv -WorkloadType $w -DetailCsvPath $dbWorkloadDetailFiles[$w] -WorkloadStats $dbWorkloadStats[$w] -AllProjects $targetProjects -DbErrors $dbErrors
        Remove-Item -Path $dbWorkloadDetailFiles[$w] -Force -ErrorAction SilentlyContinue
        Write-Log -Level INFO -Message ("DB workload CSV written: {0}" -f (Split-Path $wCsv -Leaf))
    }
}

if ($selected.ContainsKey('GKE')) {
    $pvGb = Round-Decimal -Value $stats.GKE.TotalPvGb -Scale 6
    $nodeGb = Round-Decimal -Value $stats.GKE.TotalNodeGb -Scale 6
    Add-ContentSafe -Path $csvPaths.Gke -Value ''
    Add-ContentSafe -Path $csvPaths.Gke -Value '### GKE Summary ###'
    Add-ContentSafe -Path $csvPaths.Gke -Value ("TotalClusters,{0},TotalPVCapGB,{1},TotalEstimatedNodeDiskGB,{2}" -f $stats.GKE.TotalCount,$pvGb,$nodeGb)
    Write-Log -Level INFO -Message ("GKE CSV written: {0}" -f (Split-Path $csvPaths.Gke -Leaf))
}

# -------------------------
# Known-error extraction (from runtime log)
# -------------------------
$inventoryErrors = @(Get-InventoryErrorRecordsFromLog -LogPath $script:LogFile)
if ($inventoryErrors.Count -gt 0) {
    $errCsv = Join-Path $outDir ("gcp_inventory_errors_" + $dateStr + ".csv")
    Write-PlainCsv -Data $inventoryErrors -Path $errCsv
    Write-Log -Level INFO -Message ("Inventory error CSV written: {0} (rows={1})" -f (Split-Path $errCsv -Leaf),$inventoryErrors.Count)
}

# -------------------------
# Global summary (old-style detailed rollups)
# -------------------------
function New-BlankSummaryRow {
    return [PSCustomObject]@{
        Level=''
        ResourceType=''
        Project=''
        Region=''
        Zone=''
        Count=''
        TotalSizeGB=''
        TotalSizeTB=''
        TotalSizeTiB=''
    }
}

$summaryRows = New-Object System.Collections.Generic.List[object]
$runElapsedTotal = (Get-Date) - $script:RunStart
$runElapsedSeconds = [int][Math]::Round($runElapsedTotal.TotalSeconds,0,[System.MidpointRounding]::AwayFromZero)
$runElapsedText = $runElapsedTotal.ToString('hh\:mm\:ss')
$summaryRows.Add([PSCustomObject]@{ Level='RunMeta'; ResourceType='Execution'; Project='All'; Region='All'; Zone='All'; Count=$runElapsedSeconds; TotalSizeGB=$runElapsedText; TotalSizeTB=''; TotalSizeTiB='' }) | Out-Null

$vmGb = Round-Decimal -Value $stats.VM.DiskGbTotal -Scale 6
$bucketBytes = Convert-ToDecimal $stats.Storage.TotalBytes
$bucketGb = Convert-BytesToDecimalGB -Bytes $bucketBytes -Scale 6
$fsGb = Round-Decimal -Value $stats.Fileshare.TotalGb -Scale 6
$dbGb = Round-Decimal -Value $stats.DB.TotalGb -Scale 6
$gkeGb = Get-GkeEffectiveSizeFromSums -PvGb $stats.GKE.TotalPvGb -NodeGb $stats.GKE.TotalNodeGb
$overallGb = Round-Decimal -Value ($vmGb + $bucketGb + $fsGb + $dbGb + $gkeGb) -Scale 6

$summaryRows.Add([PSCustomObject]@{ Level='Overall'; ResourceType='AllResources'; Project='All'; Region='All'; Zone='All'; Count='N/A'; TotalSizeGB=$overallGb; TotalSizeTB=(Convert-GBToDecimalTB -GB $overallGb -Scale 6); TotalSizeTiB=(Round-Decimal -Value ($overallGb/[decimal]1024) -Scale 6) }) | Out-Null
if ($selected.ContainsKey('VM')) {
    $summaryRows.Add([PSCustomObject]@{ Level='Overall'; ResourceType='VM'; Project='All'; Region='All'; Zone='All'; Count=$stats.VM.TotalCount; TotalSizeGB=(Round-Decimal -Value $vmGb -Scale 6); TotalSizeTB=(Convert-GBToDecimalTB -GB $vmGb -Scale 6); TotalSizeTiB=(Round-Decimal -Value ($vmGb/[decimal]1024) -Scale 6) }) | Out-Null
}
if ($selected.ContainsKey('STORAGE')) {
    $summaryRows.Add([PSCustomObject]@{ Level='Overall'; ResourceType='StorageBucket'; Project='All'; Region='All'; Zone='All'; Count=$stats.Storage.TotalCount; TotalSizeGB=$bucketGb; TotalSizeTB=(Convert-BytesToDecimalTB -Bytes $bucketBytes -Scale 6); TotalSizeTiB=(Convert-BytesToTiB -Bytes $bucketBytes -Scale 6) }) | Out-Null
}
if ($selected.ContainsKey('FILESHARE')) {
    $summaryRows.Add([PSCustomObject]@{ Level='Overall'; ResourceType='Fileshare'; Project='All'; Region='All'; Zone='All'; Count=$stats.Fileshare.TotalCount; TotalSizeGB=(Round-Decimal -Value $fsGb -Scale 6); TotalSizeTB=(Convert-GBToDecimalTB -GB $fsGb -Scale 6); TotalSizeTiB=(Round-Decimal -Value ($fsGb/[decimal]1024) -Scale 6) }) | Out-Null
}
if ($selected.ContainsKey('DB')) {
    $summaryRows.Add([PSCustomObject]@{ Level='Overall'; ResourceType='Database'; Project='All'; Region='All'; Zone='All'; Count=$stats.DB.RowCount; TotalSizeGB=(Round-Decimal -Value $dbGb -Scale 6); TotalSizeTB=(Convert-GBToDecimalTB -GB $dbGb -Scale 6); TotalSizeTiB=(Round-Decimal -Value ($dbGb/[decimal]1024) -Scale 6) }) | Out-Null
}
if ($selected.ContainsKey('GKE')) {
    $summaryRows.Add([PSCustomObject]@{ Level='Overall'; ResourceType='GKECluster'; Project='All'; Region='All'; Zone='All'; Count=$stats.GKE.TotalCount; TotalSizeGB=(Round-Decimal -Value $gkeGb -Scale 6); TotalSizeTB=(Convert-GBToDecimalTB -GB $gkeGb -Scale 6); TotalSizeTiB=(Round-Decimal -Value ($gkeGb/[decimal]1024) -Scale 6) }) | Out-Null
}

1..4 | ForEach-Object { $summaryRows.Add((New-BlankSummaryRow)) | Out-Null }

foreach ($proj in $targetProjects) {
    if ($selected.ContainsKey('VM')) {
        $projCount = Get-MapCount -Map $stats.VM.ProjectCounts -Key $proj
        if ($projCount -gt 0) {
            $projGb = Get-MapDecimal -Map $stats.VM.DiskGbByProject -Key $proj
            $summaryRows.Add([PSCustomObject]@{ Level='Project'; ResourceType='VM'; Project=$proj; Region='All'; Zone='All'; Count=$projCount; TotalSizeGB=(Round-Decimal -Value $projGb -Scale 6); TotalSizeTB=(Convert-GBToDecimalTB -GB $projGb -Scale 6); TotalSizeTiB=(Round-Decimal -Value ($projGb/[decimal]1024) -Scale 6) }) | Out-Null
        }
    }

    if ($selected.ContainsKey('STORAGE')) {
        $projCount = Get-MapCount -Map $stats.Storage.ProjectCounts -Key $proj
        if ($projCount -gt 0) {
            $projBytes = Get-MapDecimal -Map $stats.Storage.ProjectBytes -Key $proj
            $summaryRows.Add([PSCustomObject]@{ Level='Project'; ResourceType='StorageBucket'; Project=$proj; Region='All'; Zone='All'; Count=$projCount; TotalSizeGB=(Convert-BytesToDecimalGB -Bytes $projBytes -Scale 6); TotalSizeTB=(Convert-BytesToDecimalTB -Bytes $projBytes -Scale 6); TotalSizeTiB=(Convert-BytesToTiB -Bytes $projBytes -Scale 6) }) | Out-Null
        }
    }

    if ($selected.ContainsKey('FILESHARE')) {
        $projCount = Get-MapCount -Map $stats.Fileshare.ProjectCounts -Key $proj
        if ($projCount -gt 0) {
            $projFsGb = Get-MapDecimal -Map $stats.Fileshare.ProjectGb -Key $proj
            $summaryRows.Add([PSCustomObject]@{ Level='Project'; ResourceType='Fileshare'; Project=$proj; Region='All'; Zone='All'; Count=$projCount; TotalSizeGB=(Round-Decimal -Value $projFsGb -Scale 6); TotalSizeTB=(Convert-GBToDecimalTB -GB $projFsGb -Scale 6); TotalSizeTiB=(Round-Decimal -Value ($projFsGb/[decimal]1024) -Scale 6) }) | Out-Null
        }
    }

    if ($selected.ContainsKey('DB')) {
        $projCount = Get-MapCount -Map $stats.DB.ProjectCount -Key $proj
        if ($projCount -gt 0) {
            $projDbGb = Get-MapDecimal -Map $stats.DB.ProjectGb -Key $proj
            $summaryRows.Add([PSCustomObject]@{ Level='Project'; ResourceType='Database'; Project=$proj; Region='All'; Zone='All'; Count=$projCount; TotalSizeGB=(Round-Decimal -Value $projDbGb -Scale 6); TotalSizeTB=(Convert-GBToDecimalTB -GB $projDbGb -Scale 6); TotalSizeTiB=(Round-Decimal -Value ($projDbGb/[decimal]1024) -Scale 6) }) | Out-Null
        }
    }

    if ($selected.ContainsKey('GKE')) {
        $projCount = Get-MapCount -Map $stats.GKE.ProjectCounts -Key $proj
        if ($projCount -gt 0) {
            $projPv = Get-MapDecimal -Map $stats.GKE.ProjectPvGb -Key $proj
            $projNode = Get-MapDecimal -Map $stats.GKE.ProjectNodeGb -Key $proj
            $projGkeGb = Get-GkeEffectiveSizeFromSums -PvGb $projPv -NodeGb $projNode
            $summaryRows.Add([PSCustomObject]@{ Level='Project'; ResourceType='GKECluster'; Project=$proj; Region='All'; Zone='All'; Count=$projCount; TotalSizeGB=(Round-Decimal -Value $projGkeGb -Scale 6); TotalSizeTB=(Convert-GBToDecimalTB -GB $projGkeGb -Scale 6); TotalSizeTiB=(Round-Decimal -Value ($projGkeGb/[decimal]1024) -Scale 6) }) | Out-Null
        }
    }
}

1..4 | ForEach-Object { $summaryRows.Add((New-BlankSummaryRow)) | Out-Null }

if ($selected.ContainsKey('VM')) {
    foreach ($region in @($stats.VM.RegionCounts.Keys | Sort-Object)) {
        $count = Get-MapCount -Map $stats.VM.RegionCounts -Key $region
        $regionGb = Get-MapDecimal -Map $stats.VM.DiskGbByRegion -Key $region
        $summaryRows.Add([PSCustomObject]@{ Level='Region'; ResourceType='VM'; Project='All'; Region=$region; Zone='All'; Count=$count; TotalSizeGB=(Round-Decimal -Value $regionGb -Scale 6); TotalSizeTB=(Convert-GBToDecimalTB -GB $regionGb -Scale 6); TotalSizeTiB=(Round-Decimal -Value ($regionGb/[decimal]1024) -Scale 6) }) | Out-Null
    }
}

if ($selected.ContainsKey('STORAGE')) {
    foreach ($region in @($stats.Storage.RegionCounts.Keys | Sort-Object)) {
        $count = Get-MapCount -Map $stats.Storage.RegionCounts -Key $region
        $locBytes = Get-MapDecimal -Map $stats.Storage.RegionBytes -Key $region
        $summaryRows.Add([PSCustomObject]@{ Level='Region'; ResourceType='StorageBucket'; Project='All'; Region=$region; Zone='All'; Count=$count; TotalSizeGB=(Convert-BytesToDecimalGB -Bytes $locBytes -Scale 6); TotalSizeTB=(Convert-BytesToDecimalTB -Bytes $locBytes -Scale 6); TotalSizeTiB=(Convert-BytesToTiB -Bytes $locBytes -Scale 6) }) | Out-Null
    }
}

if ($selected.ContainsKey('FILESHARE')) {
    foreach ($region in @($stats.Fileshare.RegionCounts.Keys | Sort-Object)) {
        $count = Get-MapCount -Map $stats.Fileshare.RegionCounts -Key $region
        $rGb = Get-MapDecimal -Map $stats.Fileshare.RegionGb -Key $region
        $summaryRows.Add([PSCustomObject]@{ Level='Region'; ResourceType='Fileshare'; Project='All'; Region=$region; Zone='All'; Count=$count; TotalSizeGB=(Round-Decimal -Value $rGb -Scale 6); TotalSizeTB=(Convert-GBToDecimalTB -GB $rGb -Scale 6); TotalSizeTiB=(Round-Decimal -Value ($rGb/[decimal]1024) -Scale 6) }) | Out-Null
    }
}

if ($selected.ContainsKey('DB')) {
    foreach ($region in @($stats.DB.RegionCount.Keys | Sort-Object)) {
        $count = Get-MapCount -Map $stats.DB.RegionCount -Key $region
        $rDbGb = Get-MapDecimal -Map $stats.DB.RegionGb -Key $region
        $summaryRows.Add([PSCustomObject]@{ Level='Region'; ResourceType='Database'; Project='All'; Region=$region; Zone='All'; Count=$count; TotalSizeGB=(Round-Decimal -Value $rDbGb -Scale 6); TotalSizeTB=(Convert-GBToDecimalTB -GB $rDbGb -Scale 6); TotalSizeTiB=(Round-Decimal -Value ($rDbGb/[decimal]1024) -Scale 6) }) | Out-Null
    }
}

if ($selected.ContainsKey('GKE')) {
    foreach ($region in @($stats.GKE.RegionCounts.Keys | Sort-Object)) {
        $count = Get-MapCount -Map $stats.GKE.RegionCounts -Key $region
        $rPv = Get-MapDecimal -Map $stats.GKE.RegionPvGb -Key $region
        $rNode = Get-MapDecimal -Map $stats.GKE.RegionNodeGb -Key $region
        $rGkeGb = Get-GkeEffectiveSizeFromSums -PvGb $rPv -NodeGb $rNode
        $summaryRows.Add([PSCustomObject]@{ Level='Region'; ResourceType='GKECluster'; Project='All'; Region=$region; Zone='All'; Count=$count; TotalSizeGB=(Round-Decimal -Value $rGkeGb -Scale 6); TotalSizeTB=(Convert-GBToDecimalTB -GB $rGkeGb -Scale 6); TotalSizeTiB=(Round-Decimal -Value ($rGkeGb/[decimal]1024) -Scale 6) }) | Out-Null
    }
}

1..4 | ForEach-Object { $summaryRows.Add((New-BlankSummaryRow)) | Out-Null }

if ($selected.ContainsKey('VM')) {
    foreach ($proj in $targetProjects) {
        $projVmCount = Get-MapCount -Map $stats.VM.ProjectCounts -Key $proj
        if ($projVmCount -eq 0) { continue }

        $summaryRows.Add([PSCustomObject]@{ Level=("Per region/zone in project [{0}]" -f $proj); ResourceType=''; Project=''; Region=''; Zone=''; Count=''; TotalSizeGB=''; TotalSizeTB=''; TotalSizeTiB='' }) | Out-Null

        $prefix = $proj + '||'
        $projRegionKeys = @($stats.VM.ProjectRegionCounts.Keys | Where-Object { $_.StartsWith($prefix) } | Sort-Object)
        foreach ($k in $projRegionKeys) {
            $parts = Split-Key2 -Key $k
            $region = [string]$parts[1]
            $regionCount = [int]$stats.VM.ProjectRegionCounts[$k]
            $regionGb = if ($stats.VM.DiskGbByProjectRegion.ContainsKey($k)) { Convert-ToDecimal $stats.VM.DiskGbByProjectRegion[$k] } else { [decimal]0 }
            $summaryRows.Add([PSCustomObject]@{ Level='Region'; ResourceType='VM'; Project=$proj; Region=$region; Zone='All'; Count=$regionCount; TotalSizeGB=(Round-Decimal -Value $regionGb -Scale 6); TotalSizeTB=(Convert-GBToDecimalTB -GB $regionGb -Scale 6); TotalSizeTiB=(Round-Decimal -Value ($regionGb/[decimal]1024) -Scale 6) }) | Out-Null
        }

        1..2 | ForEach-Object { $summaryRows.Add((New-BlankSummaryRow)) | Out-Null }

        $projZoneKeys = @($stats.VM.ProjectZoneCounts.Keys | Where-Object { $_.StartsWith($prefix) } | Sort-Object)
        foreach ($k in $projZoneKeys) {
            $parts = Split-Key2 -Key $k
            $zoneName = [string]$parts[1]
            $zoneCount = [int]$stats.VM.ProjectZoneCounts[$k]
            $zoneGb = if ($stats.VM.DiskGbByProjectZone.ContainsKey($k)) { Convert-ToDecimal $stats.VM.DiskGbByProjectZone[$k] } else { [decimal]0 }
            $regionName = if ($stats.VM.ProjectZoneRegion.ContainsKey($k)) { [string]$stats.VM.ProjectZoneRegion[$k] } else { 'Unknown' }
            $summaryRows.Add([PSCustomObject]@{ Level='Zone'; ResourceType='VM'; Project=$proj; Region=$regionName; Zone=$zoneName; Count=$zoneCount; TotalSizeGB=(Round-Decimal -Value $zoneGb -Scale 6); TotalSizeTB=(Convert-GBToDecimalTB -GB $zoneGb -Scale 6); TotalSizeTiB=(Round-Decimal -Value ($zoneGb/[decimal]1024) -Scale 6) }) | Out-Null
        }

        $summaryRows.Add((New-BlankSummaryRow)) | Out-Null
    }
}

if ($inventoryErrors.Count -gt 0) {
    1..4 | ForEach-Object { $summaryRows.Add((New-BlankSummaryRow)) | Out-Null }

    $summaryRows.Add([PSCustomObject]@{ Level='KnownErrors'; ResourceType='AllKnownErrors'; Project='All'; Region='All'; Zone='All'; Count=$inventoryErrors.Count; TotalSizeGB=''; TotalSizeTB=''; TotalSizeTiB='' }) | Out-Null

    foreach ($eg in @($inventoryErrors | Group-Object ErrorCategory,Workload | Sort-Object Count -Descending)) {
        $first = $eg.Group | Select-Object -First 1
        $summaryRows.Add([PSCustomObject]@{ Level='KnownError'; ResourceType=[string]$first.ErrorCategory; Project=[string]$first.Workload; Region='All'; Zone='All'; Count=$eg.Count; TotalSizeGB=''; TotalSizeTB=''; TotalSizeTiB='' }) | Out-Null
    }

    $summaryRows.Add((New-BlankSummaryRow)) | Out-Null
    foreach ($pg in @($inventoryErrors | Group-Object Project,Workload | Sort-Object Count -Descending)) {
        $first = $pg.Group | Select-Object -First 1
        $topCategory = (($pg.Group | Group-Object ErrorCategory | Sort-Object Count -Descending | Select-Object -First 1).Name)
        $summaryRows.Add([PSCustomObject]@{ Level='ProjectError'; ResourceType=[string]$first.Workload; Project=[string]$first.Project; Region=[string]$topCategory; Zone='All'; Count=$pg.Count; TotalSizeGB=''; TotalSizeTB=''; TotalSizeTiB='' }) | Out-Null
    }
}

$summaryCsv = Join-Path $outDir ("gcp_inventory_summary_" + $dateStr + ".csv")
Write-PlainCsv -Data $summaryRows.ToArray() -Path $summaryCsv
Write-Log -Level INFO -Message ("Summary CSV written: {0}" -f (Split-Path $summaryCsv -Leaf))

# -------------------------
# Zip + finalize
# -------------------------
$zipFile = Join-Path $PWD ("gcp_sizing_api_" + $dateStr + ".zip")
if (-not $SkipZip) {
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    try {
        [IO.Compression.ZipFile]::CreateFromDirectory($outDir, $zipFile)
        Write-Log -Level INFO -Message ("ZIP archive created: {0}" -f $zipFile)
    } catch {
        Write-Log -Level WARN -Message ("ZIP creation failed: {0}" -f $_.Exception.Message)
    }
}

if (-not $KeepOutputDir -and -not $SkipZip -and (Test-Path $zipFile)) {
    try {
        Write-Log -Level INFO -Message ("Removing temporary directory: {0}" -f $outDir)
        Remove-Item -Path $outDir -Recurse -Force
        $script:LogFile = $null
    } catch {
        Write-Log -Level WARN -Message ("Cleanup failed for {0}: {1}" -f $outDir,$_.Exception.Message)
    }
}

$runElapsedFinal = ((Get-Date) - $script:RunStart).ToString('hh\:mm\:ss')
Write-Log -Level INFO -Message ("Inventory complete. TotalElapsed={0}" -f $runElapsedFinal)
if (Test-Path $zipFile) {
    Write-Host "`nInventory complete. Results in $zipFile`n" -ForegroundColor Green
} else {
    Write-Host "`nInventory complete. Results in $outDir`n" -ForegroundColor Green
}

