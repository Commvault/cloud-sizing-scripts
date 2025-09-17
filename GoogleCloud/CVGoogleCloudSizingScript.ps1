<#
.SYNOPSIS
    GCP Cloud Sizing Script - Fast VM and Storage inventory with correct summaries

.DESCRIPTION
    - Inventories GCP Compute Engine VMs and Cloud Storage Buckets across all or specified projects.
    - Speeds up disk lookups by using a single project-wide disk list.
    - Uses gsutil du -s for bucket sizing (fast).
    - Produces summary with VM counts only and disk sizes without double-counting regional disks.
    - Regional disks: Zone is blank, sizes roll up at Region level only.
    - Transcript is closed before zipping so the full log is included in the archive.

.PARAMETER Types
    Optional. Restrict inventory to specific resource types. 
    **Valid values**: [VM, Storage, Fileshare]
    If omitted, all resource types will be inventoried.
    Accepts any of the following forms:
        -Types VM,Storage              (unquoted comma-separated list)
        -Types "VM"                    (single type)
        -Types "VM","Storage"        (standard string array)
        -Types "VM,Storage"            (single quoted comma-separated string, case insensitive)
        
.PARAMETER Projects
        Optional. Target specific GCP projects by name or ID. If omitted, all accessible projects will be processed.
        Accepts any of the following forms:
            -Projects proj1,proj2              (unquoted comma-separated list)
            -Projects "proj1"                  (single project)
            -Projects "proj1","proj2"        (standard string array)
            -Projects "proj1,proj2"          (single quoted comma-separated string)

.OUTPUTS
        Runtime creates a timestamped working directory (gcp-inv-YYYY-MM-DD_HHMMSS) containing:
            - gcp_vm_instance_info_YYYY-MM-DD_HHMMSS.csv                (VM inventory)
            - gcp_disks_attached_to_vm_instances_YYYY-MM-DD_HHMMSS.csv  (attached disks)
            - gcp_disks_unattached_to_vm_instances_YYYY-MM-DD_HHMMSS.csv(unattached disks)
            - gcp_storage_buckets_info_YYYY-MM-DD_HHMMSS.csv   (bucket inventory)
            - gcp_inventory_summary_YYYY-MM-DD_HHMMSS.csv      (summary rollups)
            - gcp_sizing_script_output_YYYY-MM-DD_HHMMSS.log   (transcript/log)
        These files are zipped into:
            - gcp_sizing_YYYY-MM-DD_HHMMSS.zip
        After the ZIP is created the working directory is deleted; only the ZIP archive remains.

.NOTES
    Requires Google Cloud SDK (gcloud CLI and gsutil) installed and authenticated.
    Must be run by a user with appropriate GCP permissions.


    SETUP INSTRUCTIONS FOR GOOGLE CLOUD SHELL (Recommended):

    1. Learn about Google Cloud Shell:
        Visit: https://cloud.google.com/shell/docs

    2. Verify GCP permissions:
        Ensure your Google account has "Viewer" or higher role on target projects.

    3. Access Google Cloud Shell:
        - Login to Google Cloud Console with your account
        - Open Google Cloud Shell

    4. Upload this script:
        Use the Cloud Shell file upload feature to upload CVGoogleCloudSizingScript.ps1
        - Enter PowerShell mode, by executing the command:
            pwsh
        - run chmod +x CVGoogleCloudSizingScript.ps1 to allow the script execution permissions

    5. Run the script:
        # For all workload, all Projects
        ./CVGoogleCloudSizingScript.ps1

        # For specific workloads, all Projects
        ./CVGoogleCloudSizingScript.ps1 -Types VM,Storage

        # For all workload, specific Projects
        ./CVGoogleCloudSizingScript.ps1 -Projects my-gcp-project-1,my-gcp-project-2

        # For specific workloads, specific Projects
        ./CVGoogleCloudSizingScript.ps1 -Types VM -Projects my-gcp-project-1,my-gcp-project-2


    SETUP INSTRUCTIONS FOR LOCAL SYSTEM:

    1. Install PowerShell 7:
        Download from: https://github.com/PowerShell/PowerShell/releases

    2. Install Google Cloud SDK:
        Download from: https://cloud.google.com/sdk/docs/install

    3. Authenticate with GCP:
        gcloud auth login

    4. Verify permissions:
        Ensure your account has "Viewer" or higher role on target projects

    5. Run the script:
        # For all workload, all Projects
        ./CVGoogleCloudSizingScript.ps1

        # For specific workloads, all Projects
        ./CVGoogleCloudSizingScript.ps1 -Types VM,Storage

        # For all workload, specific Projects
        ./CVGoogleCloudSizingScript.ps1 -Projects my-gcp-project-1,my-gcp-project-2

        # For specific workloads, specific Projects
        ./CVGoogleCloudSizingScript.ps1 -Types VM -Projects my-gcp-project-1,my-gcp-project-2

    EXAMPLE USAGE
    -------------
        .\CVGoogleCloudSizingScript.ps1
        # Inventories VMs and Storage Buckets in all accessible projects

        .\CVGoogleCloudSizingScript.ps1 -Types VM,Storage
        # Explicitly inventories VMs and Storage Buckets in all projects (same as default)

        .\CVGoogleCloudSizingScript.ps1 -Types VM
        # Only inventories Compute Engine VMs in all projects

        .\CVGoogleCloudSizingScript.ps1 -Projects my-gcp-project-1,my-gcp-project-2
        # Inventories VMs and Storage Buckets in only the specified projects

        .\CVGoogleCloudSizingScript.ps1 -Types Storage -Projects my-gcp-project-1
        # Only inventories Storage Buckets in the specified project
#>


param(
    [ValidateSet('VM','Storage','Fileshare','DB','GKE', IgnoreCase = $true)]
    [string[]]$Types,
    [string[]]$Projects,
    [switch]$LightMode,               # Leaner concurrency model
    [switch]$ForceKubectl,            # Require kubectl to be present (fail if cannot provision)
    [int]$MaxVMThreads                = 10,
    [int]$MaxBucketThreads            = 20,
    [int]$VMProjectTimeoutSec         = 600,   # 10m
    [int]$BucketProjectListTimeoutSec = 300,   # 5m
    [int]$BucketSizingTimeoutSec      = 1200,  # 20m
    [int]$DbProjectTimeoutSec         = 900    # 15m per project for DB inventory (soft)
)

# Normalize -Projects if provided as a single comma-separated string inside quotes
if ($Projects -and $Projects.Count -eq 1 -and $Projects[0] -match ',') {
    $Projects = $Projects[0].Split(',') | ForEach-Object { $_.Trim() } | Where-Object { $_ }
}

# Normalize -Types if provided as a single comma-separated string inside quotes
if ($Types -and $Types.Count -eq 1 -and $Types[0] -match ',') {
    $Types = $Types[0].Split(',') | ForEach-Object { $_.Trim() } | Where-Object { $_ }
}

# Enforce non-interactive execution for all gcloud commands
$env:CLOUDSDK_CORE_DISABLE_PROMPTS = '1'

# Post-normalization validation for -Types (case-insensitive)
if ($Types) {
    $allowed = @('VM','STORAGE','FILESHARE','DB','GKE')
    $bad = $Types | Where-Object { $allowed -notcontains ($_.Trim().ToUpper()) }
    if ($bad.Count -gt 0) {
    Write-Error ("Invalid value(s) for -Types: {0}. Valid values: VM, Storage, Fileshare, DB" -f ($bad -join ', '))
        return
    }
}

# Try to remove this and always have more output in logs - $$$$$
$MinimalOutput = $false

# -------------------------
# Setup output + transcript
# -------------------------
$dateStr = (Get-Date).ToString("yyyy-MM-dd_HHmmss")
$outDir = Join-Path -Path $PWD -ChildPath ("gcp-inv-" + $dateStr)
New-Item -Path $outDir -ItemType Directory -Force | Out-Null

$transcriptFile = Join-Path $outDir ("gcp_sizing_script_output_" + $dateStr + ".log")
Start-Transcript -Path $transcriptFile -Append | Out-Null

# Expose primary log file path to child runspaces via environment variable so Write-Log can append directly
$env:GCP_INVENTORY_LOGFILE = $transcriptFile

# Global concurrent queue to collect log lines from child runspaces (VM & Bucket)
if (-not (Get-Variable -Name ChildLogQueue -Scope Global -ErrorAction SilentlyContinue)) {
    $Global:ChildLogQueue = [System.Collections.Concurrent.ConcurrentQueue[string]]::new()
}


Write-Host "=== GCP Resource Inventory Started ===" -ForegroundColor Green
if ($Types)    { Write-Host "  Types: $($Types -join ', ')" -ForegroundColor Green }
if ($Projects) { Write-Host "  Projects: $($Projects -join ', ')" -ForegroundColor Green }


# Resource type mapping
$ResourceTypeMap = @{
    "VM"        = "VMs"
    "STORAGE"   = "StorageBuckets"
    "FILESHARE" = "FileShares"
    "DB"        = "Databases"
    "GKE"       = "GKEClusters"
}

# Normalize types
if ($Types) {
    # Already validated; convert to uppercase normalized set
    $Types = $Types | ForEach-Object { $_.Trim().ToUpper() }
    $Selected = @{}
    foreach ($t in $Types) { if ($ResourceTypeMap.ContainsKey($t)) { $Selected[$t] = $true } }
    if ($Selected.Count -eq 0) { Write-Host "No valid -Types specified. Use: VM, Storage"; exit 1 }
} else { $Selected = @{}; $ResourceTypeMap.Keys | ForEach-Object { $Selected[$_] = $true } }

# -------------------------
# Helpers
# -------------------------
function Get-GcpProjects {
    try {
    # Added --quiet to suppress any interactive prompt
    $json = gcloud --quiet projects list --format=json | ConvertFrom-Json
        if (-not $json) { throw "No projects returned by gcloud." }
        return $json.projectId
    } catch {
        Write-Error "Failed to list GCP projects. Ensure gcloud SDK is installed and authenticated. Error: $_"
        Stop-Transcript | Out-Null
        exit 1
    }
}

function Get-RegionFromZone {
    param([string]$zone)
    if (-not $zone) { return "Unknown" }
    $z = $zone -replace '.*/',''
    return ($z -replace '-[a-z]$','')
}

function Ensure-Kubectl {
    param([switch]$Force)
    $env:CLOUDSDK_CORE_DISABLE_PROMPTS = '1'
    Write-Host '[Ensure-Kubectl] Ensuring kubectl presence.' -ForegroundColor Cyan

    # Detect platform
    $isWindows = $false
    try {
        if ($env:OS -eq 'Windows_NT' -or [System.Runtime.InteropServices.RuntimeInformation]::IsOSPlatform([System.Runtime.InteropServices.OSPlatform]::Windows)) { $isWindows = $true }
    } catch { if ($env:OS -match 'Windows') { $isWindows = $true } }

    function Add-PathSegment {
        param([string]$Dir,[switch]$Windows)
        if (-not $Dir) { return }
        if ($Windows) {
            if (-not (($env:PATH -split ';') -contains $Dir)) { $env:PATH = $Dir + ';' + $env:PATH }
        } else {
            if (-not (($env:PATH -split ':') -contains $Dir)) { $env:PATH = $Dir + ':' + $env:PATH }
        }
    }

    function Resolve-KubectlLocal {
        try { return (Get-Command kubectl -ErrorAction SilentlyContinue).Source } catch { return $null }
    }

    $resolved = Resolve-KubectlLocal
    if ($resolved) {
        Write-Host "[Ensure-Kubectl] Found existing kubectl: $resolved" -ForegroundColor Green
        # Don't return yet; still ensure auth plugin & component freshness.
    }

    # Try gcloud components (preferred official source)
    $gcloudCmd = Get-Command gcloud -ErrorAction SilentlyContinue
    if ($gcloudCmd) {
        Write-Host '[Ensure-Kubectl] Checking gcloud components...' -ForegroundColor Cyan
    try { $installedComponents = & gcloud components list --quiet --format="value(id)" 2>$null } catch { $installedComponents=@() }
    if (-not $installedComponents) { $installedComponents=@() } elseif ($installedComponents -is [string]) { $installedComponents=@($installedComponents) }

        # Helper: detect plugin executable presence
        $pluginExe = $null
        try {
            $pluginCmdTest = Get-Command gke-gcloud-auth-plugin -ErrorAction SilentlyContinue
            if ($pluginCmdTest) { $pluginExe = $pluginCmdTest.Source }
        } catch {}
        if (-not $pluginExe) {
            # Probe common locations relative to gcloud root
            $gcloudRoot = Split-Path -Parent (Split-Path -Parent $gcloudCmd.Source)  # .../google-cloud-sdk
            $candidatePaths = @(
                (Join-Path $gcloudRoot 'bin\gke-gcloud-auth-plugin.exe'),
                (Join-Path $gcloudRoot 'bin\gke-gcloud-auth-plugin'),
                (Join-Path (Split-Path -Parent $gcloudCmd.Source) 'gke-gcloud-auth-plugin.exe'),
                (Join-Path (Split-Path -Parent $gcloudCmd.Source) 'gke-gcloud-auth-plugin')
            )
            foreach ($cp in $candidatePaths) { if ($pluginExe) { break }; if (Test-Path $cp) { $pluginExe = $cp } }
        }

    # Always ensure both components each run (idempotent) so stale partial installs get corrected.
    $needInstall = @('kubectl','gke-gcloud-auth-plugin')
        if ($needInstall.Count -gt 0) {
            Write-Host "[Ensure-Kubectl] Installing: $($needInstall -join ', ')" -ForegroundColor Yellow
            # Windows non-interactive workaround: copy bundled python so component install can proceed headless
            if ($isWindows) {
                try {
                    Write-Host '[Ensure-Kubectl] (Windows) Preparing bundled Python for non-interactive component install...' -ForegroundColor Cyan
                    $copyOut = & gcloud components copy-bundled-python 2>&1
                    $pythonPath = ($copyOut | Select-Object -Last 1).Trim()
                    if ($pythonPath -and (Test-Path $pythonPath)) {
                        $env:CLOUDSDK_PYTHON = $pythonPath
                        Write-Host "[Ensure-Kubectl] Set CLOUDSDK_PYTHON=$pythonPath" -ForegroundColor DarkGray
                    } else {
                        Write-Warning '[Ensure-Kubectl] Could not resolve copied bundled python path; proceeding anyway.'
                    }
                } catch { Write-Warning "[Ensure-Kubectl] Bundled python copy failed: $($_.Exception.Message)" }
            }
            $componentInstallOutput = @()
            try {
                $componentInstallOutput = & gcloud components install @needInstall -q 2>&1
                $componentInstallOutput | ForEach-Object { Write-Host "[Ensure-Kubectl][gcloud] $_" -ForegroundColor DarkGray }
            } catch {
                Write-Warning "[Ensure-Kubectl] Component install error: $($_.Exception.Message)"
            }
            # Bundled python limitation handling (simpler): if failure, run copy-bundled-python and retry install once
            if ($isWindows -and ($componentInstallOutput -match 'Cannot use bundled Python installation')) {
                Write-Warning '[Ensure-Kubectl] Bundled Python restriction detected; copying bundled python and retrying.'
                try {
                    $cpy = & gcloud components copy-bundled-python 2>&1
                    $candidate = ($cpy | Select-Object -Last 1).Trim()
                    if ($candidate -and (Test-Path $candidate)) {
                        $env:CLOUDSDK_PYTHON = $candidate
                        Write-Host "[Ensure-Kubectl] (Retry) Set CLOUDSDK_PYTHON=$candidate" -ForegroundColor DarkGray
                        $componentInstallOutput = & gcloud components install @needInstall -q 2>&1
                        $componentInstallOutput | ForEach-Object { Write-Host "[Ensure-Kubectl][gcloud-rtry] $_" -ForegroundColor DarkGray }
                    } else { Write-Warning '[Ensure-Kubectl] (Retry) Could not determine copied python path.' }
                } catch { Write-Warning "[Ensure-Kubectl] (Retry) copy-bundled-python failed: $($_.Exception.Message)" }
            }
            # Re-evaluate component list after attempted install
            try { $installedComponents = & gcloud components list --quiet --format="value(id)" 2>$null } catch {}
            if ($needInstall | Where-Object { $installedComponents -notmatch $_ }) {
                Write-Warning '[Ensure-Kubectl] One or more requested components did not install successfully.'
            }
        } else { Write-Host '[Ensure-Kubectl] Required components already present.' -ForegroundColor Green }
        # Re-resolve
        $gcloudBin = Split-Path -Parent $gcloudCmd.Source
        foreach ($candidate in @('kubectl','kubectl.exe')) {
            $full = Join-Path $gcloudBin $candidate
            if (Test-Path $full) { $resolved = $full; break }
        }
        if (-not $resolved) { $resolved = Resolve-KubectlLocal }
        # Always add gcloud bin path so plugin executable is on PATH for downloaded kubectl too
        $gcloudBin = Split-Path -Parent $gcloudCmd.Source
        Add-PathSegment -Dir $gcloudBin -Windows:$isWindows
    if ($resolved) { Add-PathSegment -Dir (Split-Path -Parent $resolved) -Windows:$isWindows; Write-Host "[Ensure-Kubectl] Using kubectl from gcloud dir: $resolved" -ForegroundColor Green }
        # Ensure plugin exists / install retry if missing
        $pluginCmd = Get-Command gke-gcloud-auth-plugin -ErrorAction SilentlyContinue
        if (-not $pluginCmd) {
            Write-Warning '[Ensure-Kubectl] gke-gcloud-auth-plugin not found on PATH; attempting install (retry).'    
            try { & gcloud components install gke-gcloud-auth-plugin -q 2>&1 | ForEach-Object { Write-Host "[Ensure-Kubectl][gcloud] $_" -ForegroundColor DarkGray } } catch { Write-Warning "[Ensure-Kubectl] Plugin install retry error: $($_.Exception.Message)" }
            $pluginCmd = Get-Command gke-gcloud-auth-plugin -ErrorAction SilentlyContinue
        }
    if ($pluginCmd) { Write-Host "[Ensure-Kubectl] Auth plugin: $($pluginCmd.Source)" -ForegroundColor Green } else { Write-Warning '[Ensure-Kubectl] Auth plugin still missing.' }
    $env:USE_GKE_GCLOUD_AUTH_PLUGIN = 'True'
    if ($resolved) { $script:__KubectlResolvedPath=$resolved; return $resolved }
    }

    # Linux apt path (when gcloud not sufficient or absent)
    if (-not $isWindows) {
        $apt = Get-Command apt-get -ErrorAction SilentlyContinue
        if ($apt) {
            Write-Host '[Ensure-Kubectl] Attempting apt-get install.' -ForegroundColor Cyan
            $aptCmds = @('sudo apt-get update -y','sudo DEBIAN_FRONTEND=noninteractive apt-get install -y kubectl google-cloud-sdk-gke-gcloud-auth-plugin')
            foreach ($c in $aptCmds) {
                try { & /bin/bash -lc $c 2>&1 | ForEach-Object { Write-Host "[Ensure-Kubectl][apt] $_" -ForegroundColor DarkGray } } catch { Write-Warning "[Ensure-Kubectl] apt step failed: $($_.Exception.Message)"; break }
            }
            $resolved = Resolve-KubectlLocal
            if ($resolved) { Write-Host "[Ensure-Kubectl] Using kubectl from apt: $resolved" -ForegroundColor Green; $script:__KubectlResolvedPath=$resolved; return $resolved }
        }
    }

    # Direct download fallback (works for both platforms)
    Write-Host '[Ensure-Kubectl] Falling back to direct download.' -ForegroundColor Cyan
    try { [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor 3072 -bor 12288 } catch {}
    $stable = ''
    try { $stable = (Invoke-WebRequest -UseBasicParsing -Uri 'https://dl.k8s.io/release/stable.txt' -TimeoutSec 25).Content.Trim() } catch { $stable='v1.30.0' }
    if (-not $stable) { $stable='v1.30.0' }
    $arch='amd64'
    if ($isWindows) {
        $destDir = Join-Path $env:TEMP 'kubectl-bin'
        New-Item -ItemType Directory -Path $destDir -Force | Out-Null
        $destFile = Join-Path $destDir 'kubectl.exe'
        $url = "https://dl.k8s.io/release/$stable/bin/windows/$arch/kubectl.exe"
    } else {
        $destDir = '/tmp/kubectl-bin'
        if (-not (Test-Path $destDir)) { New-Item -ItemType Directory -Path $destDir -Force | Out-Null }
        $destFile = Join-Path $destDir 'kubectl'
        $url = "https://dl.k8s.io/release/$stable/bin/linux/$arch/kubectl"
    }
    try {
        Invoke-WebRequest -UseBasicParsing -Uri $url -OutFile $destFile -TimeoutSec 90
        if (-not $isWindows) { try { chmod +x $destFile 2>$null } catch {} }
        if (Test-Path $destFile) {
            Add-PathSegment -Dir $destDir -Windows:$isWindows
            $resolved = $destFile
            Write-Host ("[Ensure-Kubectl] Downloaded kubectl {0} -> {1}" -f $stable,$destFile) -ForegroundColor Green
        }
    } catch { Write-Warning "[Ensure-Kubectl] Download failed: $($_.Exception.Message)" }

    if ($resolved) {
        # Add gcloud bin path (if available) for plugin resolution in downloaded kubectl scenario
        if ($gcloudCmd) { Add-PathSegment -Dir (Split-Path -Parent $gcloudCmd.Source) -Windows:$isWindows }
        $env:USE_GKE_GCLOUD_AUTH_PLUGIN = 'True'
        $pluginCmd = Get-Command gke-gcloud-auth-plugin -ErrorAction SilentlyContinue
        if (-not $pluginCmd -and $gcloudCmd) {
            Write-Warning '[Ensure-Kubectl] Auth plugin not detected after download path; attempting component install.'
            try { & gcloud components install gke-gcloud-auth-plugin -q 2>&1 | ForEach-Object { Write-Host "[Ensure-Kubectl][gcloud] $_" -ForegroundColor DarkGray } } catch { Write-Warning "[Ensure-Kubectl] Plugin install attempt failed: $($_.Exception.Message)" }
            $pluginCmd = Get-Command gke-gcloud-auth-plugin -ErrorAction SilentlyContinue
        }
        if ($pluginCmd) { Write-Host "[Ensure-Kubectl] Auth plugin: $($pluginCmd.Source)" -ForegroundColor Green } else { Write-Warning '[Ensure-Kubectl] Auth plugin unavailable; kubectl GKE auth may fail.' }
        $script:__KubectlResolvedPath = $resolved
        return $resolved
    }

    if ($Force) {
        Write-Error '[Ensure-Kubectl] kubectl could not be provisioned (Force requested).'
        throw 'kubectl not available'
    } else {
        Write-Warning '[Ensure-Kubectl] kubectl not available; continuing without it.'
        return $null
    }
}


# -------------------------
# Lightweight logger 
# -------------------------
function Write-Log {
    param(
        [string]$Message,
        [ValidateSet('INFO','WARN','ERROR','DEBUG')][string]$Level = 'INFO'
    )
    $ts = (Get-Date).ToString('s')
    $line = "[$ts] [$Level] $Message"
    # Always enqueue so parent runspace can output once (avoids duplicate lines in transcript)
    if (-not (Get-Variable -Name ChildLogQueue -Scope Global -ErrorAction SilentlyContinue)) {
        $Global:ChildLogQueue = [System.Collections.Concurrent.ConcurrentQueue[string]]::new()
    }
    $Global:ChildLogQueue.Enqueue($line) | Out-Null
    # Also emit immediately so user sees progress; transcript will still capture (possible duplicate suppression not critical).
    switch ($Level) {
        'ERROR' { Write-Host $line -ForegroundColor Red }
        'WARN'  { Write-Host $line -ForegroundColor Yellow }
        'INFO'  { Write-Host $line -ForegroundColor Gray }
        'DEBUG' { if (-not $MinimalOutput) { Write-Host $line -ForegroundColor DarkGray } }
    }
    # Best-effort direct append to master log file (covers runspace output not captured by transcript)
    try {
        if ($env:GCP_INVENTORY_LOGFILE) {
            Add-Content -Path $env:GCP_INVENTORY_LOGFILE -Value $line -Encoding UTF8 -ErrorAction SilentlyContinue
        }
    } catch {}
}



# -------------------------
# VM + Disk Inventory (fast)
# -------------------------
function Get-GcpVMInventory {
    param(
        [string[]]$ProjectIds,
        [switch]$LightMode,
        [int]$MaxThreads = 10,
        [int]$TaskTimeoutSec = 600
    )
    # ScriptBlock executed per project (returns inventory + log lines)
    $vmProjectScriptBlock = {
        param($proj,$minimalFlag)
        $log = New-Object System.Collections.Generic.List[string]
        $startUtc = [DateTime]::UtcNow
        $log.Add("[VM-Project-Start] $proj") | Out-Null

        # Helper to extract region from zone - inner version to avoid cross-runspace issues
        function Get-RegionFromZoneInner {
            param([string]$zone)
            if (-not $zone) { return 'Unknown' }
            $z = $zone -replace '.*/',''
            return ($z -replace '-[a-z]$','')
        }

        # States Initialization to ensure error handling is done correctly
        $apiDisabled = $false
        $permIssue   = $false
        $env:CLOUDSDK_CORE_DISABLE_PROMPTS = '1'

        # Instances List project level - Try block
        try {
            $vmRaw = & gcloud --quiet compute instances list --project $proj --format=json 2>&1
            if ($LASTEXITCODE -ne 0) {
                $msg = ($vmRaw | Out-String).Trim()
                if     ($msg -match '(?i)not enabled|has not been used|is disabled|API .* not enabled') { $apiDisabled = $true }
                elseif ($msg -match '(?i)permission|denied|forbidden|403|PERMISSION_DENIED|insufficientPermissions') { $permIssue  = $true }
                $log.Add("[VM-Project-Warn] $proj instances list failed exit=$LASTEXITCODE msg=$msg") | Out-Null
                $vmList = @()
            } else {
                if ([string]::IsNullOrWhiteSpace(($vmRaw | Out-String))) {
                    $vmList = @()
                } else {
                    try {
                        $vmList = $vmRaw | ConvertFrom-Json
                    } catch {
                        $log.Add("[VM-Project-Error] $proj instances JSON parse failed: $($_.Exception.Message)") | Out-Null
                        # $$$$ check if this empty list causes data falseness or loss later
                        $vmList=@()
                    }
                }
            }
        } 
        
        # Instances List Project level - Exception handling
        catch {
            $log.Add("[VM-Project-Error] $proj instances command threw: $($_.Exception.Message)") | Out-Null
            # $$$$ check if this empty list causes data falseness or loss later
            $vmList=@()
        }

        # Early exit if API disabled or permission issue
        if ($apiDisabled) {
            $log.Add("[VM-Project-Skip] $proj Compute API disabled - skipping VMs & disks") | Out-Null
            return [PSCustomObject]@{ Project=$proj; VMs=@(); AttachedDisks=@(); AllDisks=@(); UnattachedDisks=@(); Logs=$log; DurationSec=[math]::Round(([DateTime]::UtcNow - $startUtc).TotalSeconds,2) }
        }
        if ($permIssue) {
            $log.Add("[VM-Project-Skip] $proj Compute API permission issue - skipping VMs & disks") | Out-Null
            return [PSCustomObject]@{ Project=$proj; VMs=@(); AttachedDisks=@(); AllDisks=@(); UnattachedDisks=@(); Logs=$log; DurationSec=[math]::Round(([DateTime]::UtcNow - $startUtc).TotalSeconds,2) }
        }

        # Disks List project level - Try block
        try {
            $diskRaw = & gcloud --quiet compute disks list --project $proj --format=json 2>&1
            if ($LASTEXITCODE -ne 0) {
                $dmsg = ($diskRaw | Out-String).Trim()
                $log.Add("[VM-Project-Warn] $proj disks list failed exit=$LASTEXITCODE msg=$dmsg") | Out-Null
                $diskListAll = @()
            } else {
                if ([string]::IsNullOrWhiteSpace(($diskRaw | Out-String))) { 
                    $diskListAll=@() 
                }
                else {
                     try { 
                        $diskListAll = $diskRaw | ConvertFrom-Json 
                    } 
                    catch { 
                        $log.Add("[VM-Project-Error] $proj disks JSON parse failed: $($_.Exception.Message)") | Out-Null; $diskListAll=@() 
                    } 
                }
            }
        } 
        # Disks List Project level - Exception handling
        catch {
            $log.Add("[VM-Project-Error] $proj disks command threw: $($_.Exception.Message)") | Out-Null
            # $$$$ check if this empty list causes data falseness or loss later
            $diskListAll=@()
        }

        if (-not $vmList)      { $vmList      = @() }
        if (-not $diskListAll) { $diskListAll = @() }

        # Build disk lookup map. 
        # We now key by (a) selfLink (preferred), and (b) composite "<zoneOrRegion>|<name>". We still keep first-seen entries for composite keys.
        $diskMap = @{}
        foreach ($d in $diskListAll) {
            if (-not $d.name) { continue }
            $selfKey = $null
            
            # Assign Self Link key if available - From disk properties
            if ($d.PSObject.Properties.Name -contains 'selfLink' -and $d.selfLink) {
                $selfKey = $d.selfLink.ToLower() 
            }

            # Fallback to id if selfLink not present - From disk properties
            elseif ($d.PSObject.Properties.Name -contains 'id' -and $d.id) { 
                $selfKey = ("id://" + $d.id) 
            }

            # Fallback - Synthetic key from zone/region + name if no selfLink or id
            else {
                # Fallback synthetic key
                $zr = if ($d.region) { ($d.region -split '/')[-1] } elseif ($d.zone) { ($d.zone -split '/')[-1] } else { '' }
                $selfKey = ("//disk/" + $zr + "/" + $d.name).ToLower()
            }
            
            # Add to map by selfLink/id key (ensures duplicate disks are handled and no data loss)
            $diskMap[$selfKey] = $d
            
            $zr2 = if ($d.region) { 
                ($d.region -split '/')[-1] 
            } elseif ($d.zone) { 
                ($d.zone -split '/')[-1] 
            } else { '' }

            # Add to map by composite key (zone|name or region|name) if not already present (first-seen wins)
            $composite = ($zr2 + '|' + $d.name).ToLower()
            if (-not $diskMap.ContainsKey($composite)) { $diskMap[$composite] = $d }
        }

        # Process VMs and attached disks - Variables to hold results
        $projectVMs       = @()
        $projectAttached  = @()
        $projectAllDisks  = @()
        $projectUnattached= @()
        $vmIndex = 0

        foreach ($vm in $vmList) {
            $vmIndex++
            $zone  = ($vm.zone -replace '.*/','')
            $region = Get-RegionFromZoneInner -zone $vm.zone
            $osType = 'Linux'

            # OS Type detection - from disks - Try block
            try {
                if ($vm.disks) {
                    foreach ($vd in $vm.disks) {
                        if ($vd.licenses) {
                            foreach ($lic in $vd.licenses) { 
                                if ($lic -match 'windows') {
                                    $osType='Windows';
                                    break 
                                } 
                            }
                            if ($osType -eq 'Windows') 
                            { 
                                break 
                            }
                        }
                    }
                }
            } 
            # OS Type detection - fallback - Try block
            catch { 
                $osType='Linux' 
            }

            $vmDiskGB = 0
            if ($vm.disks) {
                if (-not $script:__AllDiskKeys) { $script:__AllDiskKeys = @{} }
                foreach ($disk in $vm.disks) {
                    $diskName = ($disk.source -split '/')[-1]
                    $primaryKey = if ($disk.source) { $disk.source.ToLower() } else { $null }
                    $d = $null
                    $fromMap = $false
                    if ($primaryKey -and $diskMap.ContainsKey($primaryKey)) { 
                        $d = $diskMap[$primaryKey];
                        $fromMap=$true 
                    }
                    else {
                        # Derive composite key from URL for lookup
                        $zrMatch = ''
                        if ($disk.source -match '/zones/([^/]+)/') { 
                            $zrMatch = $Matches[1] 
                        }
                        elseif ($disk.source -match '/regions/([^/]+)/') {
                            $zrMatch = $Matches[1] 
                        }
                        $altKey = ($zrMatch + '|' + $diskName).ToLower()
                        if ($diskMap.ContainsKey($altKey)) {
                            $d = $diskMap[$altKey]; $fromMap=$true 
                        }
                    }
                    # Choose metadata source (disk list object when available, otherwise instance attachment data)
                    $sizeGbVal = 0
                    $regionVal = ''
                    $zoneVal = ''
                    $isRegional = $false
                    $enc = 'No'
                    $typeVal = ''
                    $selfLinkVal = $null
                    if ($d) {
                        $sizeGbVal = [int64]$d.sizeGb
                        if ($d.region) 
                        { 
                            $regionVal = ($d.region -split '/')[-1];
                            $isRegional = $true 
                        } 
                        else { 
                            $zoneVal = ($d.zone -split '/')[-1]; 
                            $regionVal = Get-RegionFromZoneInner -zone $d.zone 
                        }
                        if ($d.diskEncryptionKey -or $d.encryptionKey) {
                            $enc='Yes' 
                        }
                        $typeVal = ($d.type -replace '.*/','')
                        
                        if ($d.PSObject.Properties.Name -contains 'selfLink') {
                            $selfLinkVal = $d.selfLink 
                        }
                    } 
                    else {
                        # Use attachment data only
                        if ($disk.PSObject.Properties.Name -contains 'diskSizeGb' -and $disk.diskSizeGb) {
                            $sizeGbVal = [int64]$disk.diskSizeGb 
                        }
                        if ($disk.source -match '/zones/([^/]+)/') {
                            $zoneVal = $Matches[1]; $regionVal = Get-RegionFromZoneInner -zone $zoneVal 
                        }
                        elseif ($disk.source -match '/regions/([^/]+)/') {
                            $regionVal = $Matches[1]; $isRegional = $true 
                        }
                        if ($disk.PSObject.Properties.Name -contains 'type' -and $disk.type) {
                            $typeVal = ($disk.type -replace '.*/','') 
                        }
                        $selfLinkVal = $disk.source
                    }

                    # Accumulate disk size (regional disks counted once at region level)
                    $vmDiskGB += $sizeGbVal
                    $attachedObj = [PSCustomObject]@{
                        DiskName     = $diskName
                        VMName       = $vm.name
                        Project      = $proj
                        Region       = $regionVal
                        Zone         = if ($isRegional) { '' } else { $zoneVal }
                        IsRegional   = [bool]$isRegional
                        Encrypted    = $enc
                        DiskType     = $typeVal
                        SizeGB       = $sizeGbVal
                        DiskSelfLink = $selfLinkVal
                        DiskKey      = $primaryKey
                        Source       = if ($fromMap) { 'DiskList' } else { 'InstanceAttachment' }
                    }
                    $projectAttached += $attachedObj
                    # Ensure disk appears once in AllDisks (prefer enriched version if later disk list supplies it)
                    $diskKeyForAll = if ($selfLinkVal) { $selfLinkVal.ToLower() } elseif ($primaryKey) { $primaryKey } else { ($regionVal + '|' + $diskName).ToLower() }
                    if (-not $script:__AllDiskKeys.ContainsKey($diskKeyForAll)) {
                        $script:__AllDiskKeys[$diskKeyForAll] = $true
                        $projectAllDisks += [PSCustomObject]@{
                            DiskName     = $diskName
                            VMName       = $vm.name
                            Project      = $proj
                            Region       = $regionVal
                            Zone         = if ($isRegional) { '' } else { $zoneVal }
                            IsRegional   = [bool]$isRegional
                            Encrypted    = $enc
                            DiskType     = $typeVal
                            SizeGB       = $sizeGbVal
                            DiskSelfLink = $selfLinkVal
                            DiskKey      = $diskKeyForAll
                        }
                    }
                }
            }
            $diskCountLocal = if ($vm.disks) { $vm.disks.Count } else { 0 }
            $log.Add("[VM] Project=$proj $vmIndex/$($vmList.Count) Name=$($vm.name) Type=$(($vm.machineType -replace '.*/','')) Region=$region Zone=$zone Disks=$diskCountLocal DiskGB=$vmDiskGB") | Out-Null
            $projectVMs += [PSCustomObject]@{
                Project      = $proj
                VMName       = $vm.name
                VMSize       = ($vm.machineType -replace '.*/','')
                OS           = $osType
                Region       = $region
                Zone         = $zone
                VMId         = $vm.id
                DiskCount    = $diskCountLocal
                VMDiskSizeGB = [int64]$vmDiskGB
            }
        }

        foreach ($disk in $diskListAll) {
            if (-not $script:__AllDiskKeys) { $script:__AllDiskKeys = @{} }
            $isRegional = ($null -ne $disk.region)
            $diskSelf = (if ($disk.PSObject.Properties.Name -contains 'selfLink') { $disk.selfLink } else { $null })
            $diskKeyForAll = if ($diskSelf) { $diskSelf.ToLower() } else { ((if ($disk.region) { ($disk.region -split '/')[-1] } elseif ($disk.zone) { ($disk.zone -split '/')[-1] } else { '' }) + '|' + $disk.name).ToLower() }
            # Skip if already added during VM attachment processing
            if ($script:__AllDiskKeys.ContainsKey($diskKeyForAll)) { continue }
            $script:__AllDiskKeys[$diskKeyForAll] = $true
            $diskObj = [PSCustomObject]@{
                DiskName     = $disk.name
                VMName       = if ($disk.users -and $disk.users.Count -gt 0) { ($disk.users | ForEach-Object { ($_ -split '/')[-1] }) -join ',' } else { $null }
                Project      = $proj
                Region       = if ($disk.region) { ($disk.region -split '/')[-1] } else { (Get-RegionFromZoneInner -zone $disk.zone) }
                Zone         = if ($disk.region) { '' } else { ($disk.zone -split '/')[-1] }
                IsRegional   = [bool]$isRegional
                Encrypted    = if ($disk.diskEncryptionKey -or $disk.encryptionKey) { 'Yes' } else { 'No' }
                DiskType     = ($disk.type -replace '.*/','')
                SizeGB       = [int64]$disk.sizeGb
                DiskSelfLink = $diskSelf
                DiskKey      = $diskKeyForAll
            }
            $projectAllDisks += $diskObj
            if (-not $disk.users -or $disk.users.Count -eq 0) { $projectUnattached += $diskObj }
        }

        $log.Add("[VM-Project-Debug] $proj RawDisksListed=$($diskListAll.Count) AllDisksCaptured=$($projectAllDisks.Count) VMs=$($vmList.Count)") | Out-Null

    # Reconstruction block removed: we now always capture attachment disks during VM iteration.

        $log.Add("[VM-Project-End] $proj VMs=$($projectVMs.Count) Disks=$($projectAllDisks.Count)") | Out-Null
        $durationSec = [math]::Round(([DateTime]::UtcNow - $startUtc).TotalSeconds,2)
        return [PSCustomObject]@{
            Project         = $proj
            VMs             = $projectVMs
            AttachedDisks   = $projectAttached
            AllDisks        = $projectAllDisks
            UnattachedDisks = $projectUnattached
            Logs            = $log
            DurationSec     = $durationSec
            ApiDisabled     = $apiDisabled
            PermissionIssue = $permIssue
        }
    }
    # End - VM Project ScriptBlock

    $effectiveMax = [Math]::Min($MaxThreads, [Math]::Max(1,$ProjectIds.Count))
    $vmStatuses = New-Object System.Collections.Generic.List[object]
    if (-not $LightMode) {
        # Existing runspace-pool implementation (default)
        $iss = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()
        $pool = [RunspaceFactory]::CreateRunspacePool(1,$effectiveMax,$iss,$Host); $pool.Open()
        $runspaces=@()
        foreach ($p in $ProjectIds) {
            $ps = [PowerShell]::Create().AddScript($vmProjectScriptBlock).AddArgument($p).AddArgument($MinimalOutput); $ps.RunspacePool=$pool
            $handle = $ps.BeginInvoke(); $runspaces += [PSCustomObject]@{ PS=$ps; Handle=$handle; Project=$p; Submitted=[DateTime]::UtcNow }
        }
        $allVMs=@(); $allAttached=@(); $allAllDisks=@(); $allUnattached=@(); $completed=0; $total=$ProjectIds.Count
        $durations = New-Object System.Collections.Generic.List[double]; $overallStart=[DateTime]::UtcNow
        Write-Progress -Id 101 -Activity 'VM Projects' -Status 'Queued...' -PercentComplete 0
        while ($runspaces.Count -gt 0) {
            $still=@()
            foreach ($rs in $runspaces) {
                if ($rs.Handle.IsCompleted) {
                    try { $result = $rs.PS.EndInvoke($rs.Handle) } catch { Write-Host "[VM-ERROR] Project=$($rs.Project) $_" -ForegroundColor Red; $result=$null }
                    $completed++
                    if ($result) {
                        $allVMs += $result.VMs; $allAttached += $result.AttachedDisks; $allAllDisks += $result.AllDisks; $allUnattached += $result.UnattachedDisks
                        $durations.Add($result.DurationSec/60) | Out-Null
                        foreach ($l in $result.Logs) { Write-Host $l -ForegroundColor DarkGray }
                        Write-Host ("[Project-Done] {0} VMs={1} Disks={2} ElapsedSec={3}" -f $result.Project,$result.VMs.Count,$result.AllDisks.Count,$result.DurationSec) -ForegroundColor Green
                        Write-Host '--------------------------------------------------' -ForegroundColor DarkGray
                        $statusObj = [PSCustomObject]@{
                            Project=$result.Project
                            Success= -not ($result.ApiDisabled -or $result.PermissionIssue)
                            VMCount=$result.VMs.Count
                            DiskCount=$result.AllDisks.Count
                            ApiDisabled=$result.ApiDisabled
                            PermissionIssue=$result.PermissionIssue
                            Timeout=$false
                        }
                        $vmStatuses.Add($statusObj) | Out-Null
                    }
                } else {
                    $elapsedSec = ([DateTime]::UtcNow - $rs.Submitted).TotalSeconds
                    if ($elapsedSec -ge $TaskTimeoutSec) {
                        try { $rs.PS.Stop() } catch {}
                        try { $rs.PS.Dispose() } catch {}
                        $completed++
                        Write-Host ("[VM-Project-Timeout] Project={0} TimeoutSec={1}" -f $rs.Project,$TaskTimeoutSec) -ForegroundColor Yellow
                        $vmStatuses.Add([PSCustomObject]@{ Project=$rs.Project; Success=$false; VMCount=0; DiskCount=0; ApiDisabled=$false; PermissionIssue=$false; Timeout=$true }) | Out-Null
                    } else { $still += $rs }
                }
            }
            $runspaces = $still
            $pct = if ($total -gt 0) { [math]::Round(($completed / $total)*100,1) } else { 100 }
            $elapsedMin = [math]::Round(([DateTime]::UtcNow - $overallStart).TotalMinutes,3)
            $avgMin = if ($durations.Count -gt 0) { [math]::Round(($durations | Measure-Object -Average | Select -Expand Average),3) } else { 0 }
            $remaining = $total - $completed; $etaMin = if ($avgMin -gt 0 -and $remaining -gt 0) { [math]::Round($avgMin * $remaining,2) } else { 0 }
            $rate = if ($elapsedMin -gt 0) { [math]::Round($allVMs.Count / ($elapsedMin*60),2) } else { 0 }
            $status = "Projects {0}/{1} ({2}%) | Discovered VMs={3} | Rate={4}/s | ElapsedMin={5} | ETA_Min={6}" -f $completed,$total,$pct,$allVMs.Count,$rate,$elapsedMin,$etaMin
            Write-Progress -Id 101 -Activity 'VM Projects' -Status $status -PercentComplete $pct
            if ($runspaces.Count -gt 0) { Start-Sleep -Milliseconds 200 }
        }
        Write-Progress -Id 101 -Activity 'VM Projects' -Completed
        $pool.Close(); $pool.Dispose()
        $script:VmProjectStatuses = $vmStatuses
        return @{ VMs=$allVMs; AttachedDisks=$allAttached; AllDisks=$allAllDisks; UnattachedDisks=$allUnattached; Statuses=$vmStatuses }
    } else {
        # Lightweight on-demand runspace approach (creates only active runspaces; disposes immediately)
        $queue = New-Object System.Collections.Queue
        foreach ($p in $ProjectIds) { $queue.Enqueue($p) }
        $active=@()
        $allVMs=@(); $allAttached=@(); $allAllDisks=@(); $allUnattached=@(); $completed=0; $total=$ProjectIds.Count
        $durations = New-Object System.Collections.Generic.List[double]; $overallStart=[DateTime]::UtcNow
        Write-Progress -Id 101 -Activity 'VM Projects' -Status 'Starting (LightMode)...' -PercentComplete 0
        while ($active.Count -gt 0 -or $queue.Count -gt 0) {
            while ($active.Count -lt $effectiveMax -and $queue.Count -gt 0) {
                $proj = $queue.Dequeue()
                $ps = [PowerShell]::Create().AddScript($vmProjectScriptBlock).AddArgument($proj).AddArgument($MinimalOutput)
                $handle = $ps.BeginInvoke()
                $active += [PSCustomObject]@{ PS=$ps; Handle=$handle; Project=$proj; Started=[DateTime]::UtcNow }
            }
            $remaining=@()
            foreach ($a in $active) {
                if ($a.Handle.IsCompleted) {
                    try { $result = $a.PS.EndInvoke($a.Handle) } catch { Write-Host "[VM-ERROR] Project=$($a.Project) $_" -ForegroundColor Red; $result=$null }
                    $a.PS.Dispose()
                    $completed++
                    if ($result) {
                        $allVMs += $result.VMs; $allAttached += $result.AttachedDisks; $allAllDisks += $result.AllDisks; $allUnattached += $result.UnattachedDisks
                        $durations.Add($result.DurationSec/60) | Out-Null
                        foreach ($l in $result.Logs) { Write-Host $l -ForegroundColor DarkGray }
                        Write-Host ("[Project-Done] {0} VMs={1} Disks={2} ElapsedSec={3}" -f $result.Project,$result.VMs.Count,$result.AllDisks.Count,$result.DurationSec) -ForegroundColor Green
                        Write-Host '--------------------------------------------------' -ForegroundColor DarkGray
                        $vmStatuses.Add([PSCustomObject]@{ Project=$result.Project; Success= -not ($result.ApiDisabled -or $result.PermissionIssue); VMCount=$result.VMs.Count; DiskCount=$result.AllDisks.Count; ApiDisabled=$result.ApiDisabled; PermissionIssue=$result.PermissionIssue; Timeout=$false }) | Out-Null
                    }
                } else {
                    $elapsedSec = ([DateTime]::UtcNow - $a.Started).TotalSeconds
                    if ($elapsedSec -ge $TaskTimeoutSec) {
                        try { $a.PS.Stop() } catch {}
                        try { $a.PS.Dispose() } catch {}
                        $completed++
                        Write-Host ("[VM-Project-Timeout] Project={0} TimeoutSec={1}" -f $a.Project,$TaskTimeoutSec) -ForegroundColor Yellow
                        $vmStatuses.Add([PSCustomObject]@{ Project=$a.Project; Success=$false; VMCount=0; DiskCount=0; ApiDisabled=$false; PermissionIssue=$false; Timeout=$true }) | Out-Null
                    } else { $remaining += $a }
                }
            }
            $active=$remaining
            $pct = if ($total -gt 0) { [math]::Round(($completed / $total)*100,1) } else { 100 }
            $elapsedMin = [math]::Round(([DateTime]::UtcNow - $overallStart).TotalMinutes,3)
            $avgMin = if ($durations.Count -gt 0) { [math]::Round(($durations | Measure-Object -Average | Select -Expand Average),3) } else { 0 }
            $remainingCount = $total - $completed; $etaMin = if ($avgMin -gt 0 -and $remainingCount -gt 0) { [math]::Round($avgMin * $remainingCount,2) } else { 0 }
            $rate = if ($elapsedMin -gt 0) { [math]::Round($allVMs.Count / ($elapsedMin*60),2) } else { 0 }
            $status = "(Light) Projects {0}/{1} ({2}%) | Discovered VMs={3} | Rate={4}/s | ElapsedMin={5} | ETA_Min={6}" -f $completed,$total,$pct,$allVMs.Count,$rate,$elapsedMin,$etaMin
            Write-Progress -Id 101 -Activity 'VM Projects' -Status $status -PercentComplete $pct
            if ($active.Count -gt 0 -or $queue.Count -gt 0) { Start-Sleep -Milliseconds 150 }
        }
        Write-Progress -Id 101 -Activity 'VM Projects' -Completed
        $script:VmProjectStatuses = $vmStatuses
        return @{ VMs=$allVMs; AttachedDisks=$allAttached; AllDisks=$allAllDisks; UnattachedDisks=$allUnattached; Statuses=$vmStatuses }
    }
}


# Helper - Update bucket progress
function Update-BucketWorkProgress {
    param([string]$Phase,[int]$Current,[int]$Total)
    if ($Total -le 0) { $Total = 1 }
    $pctRaw = ($Current / $Total) * 100
    if ($pctRaw -gt 100) { $pctRaw = 100 }
    $pct = [math]::Round($pctRaw,0)
    Write-Progress -Id 41 -ParentId 4 -Activity "Bucket Workload" -Status ("{0} ({1}/{2})" -f $Phase,$Current,$Total) -PercentComplete $pct
}


# MultiThreading - Script Blocks for Bucket sizing
$bucketSizingScriptBlock = {
    param($projectName, $bucket, $minimalFlag)
        # -------------------------
        # Bucket sizing helper
        # Strategy:
        # 1. Fast path: gsutil du -s (handles large buckets, avoids enumerating every object)
        # 2. Fallback: gcloud storage objects list + sum sizes (only if du fails or returns nothing)
        # 3. If both fail -> 0 with warning
        # -------------------------
        function Get-BucketSizeBytes {
            param(
                [Parameter(Mandatory)][string]$BucketName,
                [Parameter(Mandatory)][string]$Project
            )
            
            # Attempt gsutil du -s
            try {
                $du = gsutil du -s "gs://$BucketName" 2>$null
                if ($LASTEXITCODE -eq 0 -and $du) {
                    $firstField = ($du -split '\s+')[0]
                    if ($firstField -match '^[0-9]+$') {
                        return [int64]$firstField
                    }
                }
            } catch {
                Write-Log -Level WARN -Message ("gsutil du failed for {0} in {1}: {2}" -f $BucketName, $Project, ($_.Exception.Message))
            }
            # Fallback (can be slow for very large buckets): enumerate objects
            $sizeBytes = 0
            try {
                $sizes = gcloud --quiet storage objects list "gs://$BucketName" --project $Project --format="value(size)" 2>$null
                foreach ($s in $sizes) { if ($s -match '^[0-9]+$') { $sizeBytes += [int64]$s } }
                return [int64]$sizeBytes
            } catch {
                Write-Log -Level WARN -Message ("Fallback object enumeration failed for {0} in {1}: {2}" -f $BucketName, $Project, ($_.Exception.Message))
            }
            Write-Log -Level WARN -Message ("Unable to determine size for bucket {0} in {1} (returning 0)." -f $BucketName, $Project)
            return 0
        }

        $bucketName = $bucket.name
        
        # Get Bucket size in Bytes
        $sizeBytes = Get-BucketSizeBytes -BucketName $bucketName -Project $projectName
    if (-not $minimalFlag) { Write-Host ("[Sizing] Bucket: {0} | Project={1} | SizeBytes={2}" -f $bucketName, $projectName, $sizeBytes) -ForegroundColor DarkGray }
        
        # Precise size conversions (binary vs decimal) with more precision
        $bytes = [int64]$sizeBytes
        $GiBBytes = 1GB              # 1,073,741,824
        $MiBBytes = 1MB              # 1,048,576
        $TiBBytes = $GiBBytes * 1024 # 1,099,511,627,776
        $GBDecimalDivisor = 1e9
        $MBDecimalDivisor = 1e6
        $TBDecimalDivisor = 1e12
        $sizeMiB     = if ($bytes -gt 0) { [math]::Round($bytes / $MiBBytes, 3) } else { 0 }
        $sizeGiB     = if ($bytes -gt 0) { [math]::Round($bytes / $GiBBytes, 4) } else { 0 }
        $sizeTiB     = if ($bytes -gt 0) { [math]::Round($bytes / $TiBBytes, 6) } else { 0 }
        $sizeMBDec   = if ($bytes -gt 0) { [math]::Round($bytes / $MBDecimalDivisor, 3) } else { 0 }
        $sizeGBDec   = if ($bytes -gt 0) { [math]::Round($bytes / $GBDecimalDivisor, 4) } else { 0 }
        $sizeTBDec   = if ($bytes -gt 0) { [math]::Round($bytes / $TBDecimalDivisor, 6) } else { 0 }
        
        return [PSCustomObject]@{
            StorageBucket       = $bucket.name
            Project             = $projectName
            Location            = $bucket.location
            StorageClass        = $bucket.storageClass
            UsedCapacityBytes   = $bytes
            UsedCapacityMiB     = $sizeMiB
            UsedCapacityGiB     = $sizeGiB
            UsedCapacityTiB     = $sizeTiB
            UsedCapacityMBDec   = $sizeMBDec
            UsedCapacityGB      = $sizeGBDec
            UsedCapacityTB      = $sizeTBDec
        }
    }


# -------------------------
# Storage Inventory (fast, gcloud-only)
# -------------------------
function Get-GcpStorageInventory {
    param([string[]]$ProjectIds)

    # Phase 1: Concurrent bucket listing per project
    $listingScript = {
        param($project,$minimalFlag)
        $perm=$false; $buckets=@(); $err=$null
        try {
            # Added --quiet to ensure non-interactive bucket listing
            $raw = & gcloud --quiet storage buckets list --project $project --format=json 2>&1
            if ($LASTEXITCODE -ne 0) {
                $txt = ($raw | Out-String)
                if ($txt -match '(?i)permission|denied|forbidden|403') { $perm=$true }
            } else {
                if (-not [string]::IsNullOrWhiteSpace(($raw|Out-String))) { try { $buckets = $raw | ConvertFrom-Json } catch { $err=$_.Exception.Message; $buckets=@() } }
            }
        } catch { $err=$_.Exception.Message }
        $count = if ($perm) { -1 } else { if ($buckets) { $buckets.Count } else { 0 } }
        return [PSCustomObject]@{ Project=$project; Buckets=$buckets; BucketCount=$count; PermissionIssue=$perm; Error=$err }
    }

    $maxProjThreads = [Math]::Min(20,[Math]::Max(1,$ProjectIds.Count))
    Write-Log -Level INFO -Message ("[Buckets-Phase1] Listing {0} projects with maxThreads={1}" -f $ProjectIds.Count,$maxProjThreads)
    $iss1 = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()
    $pool1 = [RunspaceFactory]::CreateRunspacePool(1,$maxProjThreads,$iss1,$Host); $pool1.Open()
    $rsList=@()
    foreach ($p in $ProjectIds) {
        $ps=[PowerShell]::Create().AddScript($listingScript).AddArgument($p).AddArgument($MinimalOutput); $ps.RunspacePool=$pool1
        $rsList += [PSCustomObject]@{ PS=$ps; Handle=$ps.BeginInvoke(); Project=$p }
    }
    $projResults=@(); $completed=0; $total=$rsList.Count; $listStart=[DateTime]::UtcNow
    Write-Progress -Id 400 -Activity 'Bucket Listing' -Status 'Starting...' -PercentComplete 0
    while ($rsList.Count -gt 0) {
        $next=@()
        foreach ($r in $rsList) {
            if ($r.Handle.IsCompleted) {
                try { $res=$r.PS.EndInvoke($r.Handle) } catch { $res=$null; Write-Host "[Bucket-List-Error] Project=$($r.Project) $_" -ForegroundColor Red }
                if ($res) { $projResults += $res }
                $completed++
            } else { $next += $r }
        }
        $rsList=$next
        $pct = if ($total -gt 0) { [math]::Round(($completed/$total)*100,1) } else { 100 }
        $elapsed = [math]::Round(([DateTime]::UtcNow - $listStart).TotalSeconds,1)
        Write-Progress -Id 400 -Activity 'Bucket Listing' -Status ("Projects {0}/{1} ({2}%) ElapsedSec={3}" -f $completed,$total,$pct,$elapsed) -PercentComplete $pct
        if ($rsList.Count -gt 0) { Start-Sleep -Milliseconds 150 }
    }
    Write-Progress -Id 400 -Activity 'Bucket Listing' -Completed
    $pool1.Close(); $pool1.Dispose()

    # Build project status list
    $projectStatuses=@()
    foreach ($pr in $projResults) {
        $display = if ($pr.PermissionIssue) { "$($pr.Project)*" } else { $pr.Project }
        $bucketCountCsv = if ($pr.BucketCount -lt 0) { '' } else { $pr.BucketCount }
        $projectStatuses += [PSCustomObject]@{
            Project=$display
            BucketCount=$bucketCountCsv
            PermissionIssue= if ($pr.PermissionIssue){'Y'} else {''}
            Success= if ($pr.PermissionIssue -or $pr.Error) { 'N' } else { 'Y' }
            Error = if ($pr.Error){ $pr.Error } else { '' }
        }
    }

    # Consolidate all bucket descriptors
    $allBucketDescriptors = @()
    foreach ($pr in $projResults) {
        if ($pr.Buckets) {
            foreach ($b in $pr.Buckets) {
                $allBucketDescriptors += [PSCustomObject]@{ Project=$pr.Project; Name=$b.name; Location=$b.location; StorageClass=$b.storageClass; Raw=$b }
            }
        }
    }
    Write-Log -Level INFO -Message ("[Buckets-Phase1] Total discoverable buckets={0}" -f $allBucketDescriptors.Count)
    if ($allBucketDescriptors.Count -eq 0) {
        $script:StorageProjectStatuses = $projectStatuses
        return @()
    }

    # Phase 2: Concurrent sizing of all buckets globally
    $maxBucketThreads = [Math]::Min(20,[Math]::Max(1,$allBucketDescriptors.Count))
    Write-Log -Level INFO -Message ("[Buckets-Phase2] Sizing {0} buckets with maxThreads={1}" -f $allBucketDescriptors.Count,$maxBucketThreads)
    $iss2 = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()
    $pool2 = [RunspaceFactory]::CreateRunspacePool(1,$maxBucketThreads,$iss2,$Host); $pool2.Open()
    $bucketRunspaces=@(); $index=0
    foreach ($bd in $allBucketDescriptors) {
        $index++
        $ps=[PowerShell]::Create().AddScript($bucketSizingScriptBlock).AddArgument($bd.Project).AddArgument([PSCustomObject]@{ name=$bd.Name; location=$bd.Location; storageClass=$bd.StorageClass }).AddArgument($MinimalOutput)
        $ps.RunspacePool=$pool2
        $bucketRunspaces += [PSCustomObject]@{ PS=$ps; Handle=$ps.BeginInvoke(); Project=$bd.Project; Bucket=$bd.Name }
    }
    $sized=@(); $done=0; $totalBuckets=$bucketRunspaces.Count; $sizeStart=[DateTime]::UtcNow
    Write-Progress -Id 401 -Activity 'Bucket Sizing' -Status 'Queued...' -PercentComplete 0
    while ($bucketRunspaces.Count -gt 0) {
        $next=@()
        foreach ($br in $bucketRunspaces) {
            if ($br.Handle.IsCompleted) {
                try { $res=$br.PS.EndInvoke($br.Handle) } catch { $res=$null; Write-Host "[Bucket-Size-Error] Bucket=$($br.Bucket) Project=$($br.Project) $_" -ForegroundColor Red }
                if ($res) {
                    $sized += $res
                    Write-Host ("Bucket={0} Project={1} Location={2} SizeGB={3}" -f $res.StorageBucket,$res.Project,$res.Location,$res.UsedCapacityGB) -ForegroundColor Cyan
                }
                $done++
            } else { $next += $br }
        }
        $bucketRunspaces=$next
        $pct = if ($totalBuckets -gt 0) { [math]::Round(($done/$totalBuckets)*100,1) } else { 100 }
        $elapsed = [math]::Round(([DateTime]::UtcNow - $sizeStart).TotalSeconds,1)
        $elapsedMin = [math]::Round($elapsed/60,2)
        $totalBytes = ($sized | Measure-Object UsedCapacityBytes -Sum).Sum
        $totalGB = if ($totalBytes) { [math]::Round($totalBytes/1e9,3) } else { 0 }
        Write-Progress -Id 401 -Activity 'Bucket Sizing' -Status ("Buckets {0}/{1} ({2}%) SizedGB={3} ElapsedSec={4} ElapsedMin={5}" -f $done,$totalBuckets,$pct,$totalGB,$elapsed,$elapsedMin) -PercentComplete $pct
        if ($bucketRunspaces.Count -gt 0) { Start-Sleep -Milliseconds 200 }
    }
    Write-Progress -Id 401 -Activity 'Bucket Sizing' -Completed
    $pool2.Close(); $pool2.Dispose()

    $script:StorageProjectStatuses = $projectStatuses
    return $sized
}

# -------------------------
# Filestore (File Share) Inventory
# -------------------------
function Get-GcpFileShareInventory {
    param(
        [string[]]$ProjectIds,
        [switch]$LightMode,
        [int]$MaxThreads = 10,
        [int]$TaskTimeoutSec = 600
    )
    # Defensive: ensure MaxThreads is at least 1 (guard against caller passing null/0)
    if (-not $MaxThreads -or $MaxThreads -lt 1) { $MaxThreads = 1 }
    # Collect all filestore log lines for possible global reconstruction later
    $allLogs = New-Object System.Collections.Generic.List[string]
    # Script per project: list filestore instances
    $fsProjectScript = {
        param($proj,$minimalFlag)
        $log = New-Object System.Collections.Generic.List[string]
        $startUtc=[DateTime]::UtcNow
        $log.Add("[FS-Project-Start] $proj")|Out-Null
        $env:CLOUDSDK_CORE_DISABLE_PROMPTS='1'
        $instances=@()
        try {
            $raw = & gcloud --quiet filestore instances list --project $proj --format=json 2>&1
            if ($LASTEXITCODE -ne 0) {
                $msg=($raw|Out-String).Trim()
                if ($msg -match '(?i)permission|denied|forbidden|403|not enabled|is disabled') {
                    $log.Add("[FS-Project-Skip] $proj Filestore access issue: $msg")|Out-Null
                    $errShare = [PSCustomObject]@{
                        Project=$proj; InstanceName=''; ShareName='(error)'; Tier=''; Region=''; Zone=''; provisionedGb=0; provisionedGib=0; provisionedTb=0; provisionedTib=0; CapacityGB=0; Networks=''; IPAddresses=''; State='ERROR'; CreateTime=''; Labels=''; Protocol=''; Error=$msg
                    }
                    return [PSCustomObject]@{ Project=$proj; Shares=@($errShare); Logs=$log; DurationSec=[math]::Round(([DateTime]::UtcNow - $startUtc).TotalSeconds,2) }
                }
                $log.Add("[FS-Project-Warn] $proj list failed exit=$LASTEXITCODE msg=$msg")|Out-Null
            } else {
                if (-not [string]::IsNullOrWhiteSpace(($raw|Out-String))) { try { $instances = $raw | ConvertFrom-Json } catch { $log.Add("[FS-Project-Error] $proj JSON parse failed: $($_.Exception.Message)")|Out-Null } }
            }
        } catch { $log.Add("[FS-Project-Error] $proj command threw: $($_.Exception.Message)")|Out-Null }
        if (-not $instances) { $instances=@() }
        $shares=@(); $idx=0
        foreach ($inst in $instances) {
            $idx++
            $fullName=$inst.name
            # Extract location (region or zone) from full resource name
            $locationRaw=''
            if ($fullName -match '/locations/([^/]+)/instances/') { $locationRaw=$Matches[1] }
            $region=''; $zone=''
            if ($locationRaw -match '^[a-z0-9-]+-[a-z]$') { $zone=$locationRaw; $region=($locationRaw -replace '-[a-z]$','') } else { $region=$locationRaw }
            $tier=$inst.tier
            $instanceShort = if ($fullName) { ($fullName -split '/')[-1] } else { '' }
            $state = $inst.state
            $createTime = $inst.createTime
            $labels=''
            try { if ($inst.labels) { $labels = ($inst.labels.GetEnumerator() | ForEach-Object { "{0}={1}" -f $_.Key,$_.Value }) -join ';' } } catch {}
            $networksNames = @(); $ipsList=@()
            if ($inst.networks) {
                foreach ($n in $inst.networks) { if ($n.network){ $networksNames += ($n.network -replace '.*/','') }; if ($n.ipAddresses){ $ipsList += $n.ipAddresses } }
            }
            $net = ($networksNames | Sort-Object -Unique) -join ';'
            $ip = ($ipsList | Sort-Object -Unique) -join ';'
            $protocol='NFS'
            if ($inst.fileShares -and $inst.fileShares.Count -gt 0) {
                foreach ($share in $inst.fileShares) {
                    $shareName = $share.name
                    $capGiB = 0; try { if ($share.capacityGb) { $capGiB=[int64]$share.capacityGb } } catch {}
                    # capacityGb from API is GiB; derive decimal GB/TB and binary TiB
                    $provGib = [double]$capGiB
                    $provGb  = [math]::Round(($provGib * 1024 * 1024 * 1024)/1e9,3)
                    $provTib = [math]::Round($provGib/1024,4)
                    $provTb  = [math]::Round($provGb/1000,4)
                    $log.Add("[FS] Project=$proj $idx/$($instances.Count) Instance=$instanceShort Share=$shareName Tier=$tier Region=$region Zone=$zone provisionedGib=$provGib provisionedGb=$provGb")|Out-Null
                    $shares += [PSCustomObject]@{
                        Project=$proj
                        InstanceName=$instanceShort
                        ShareName=$shareName
                        Tier=$tier
                        Region=$region
                        Zone=$zone
                        provisionedGb=$provGb
                        provisionedGib=$provGib
                        provisionedTb=$provTb
                        provisionedTib=$provTib
                        CapacityGB=$capGiB   # legacy field retained
                        Networks=$net
                        IPAddresses=$ip
                        State=$state
                        CreateTime=$createTime
                        Labels=$labels
                        Protocol=$protocol
                        Error=''
                    }
                }
            } else {
                # No shares array; still emit instance row
                $log.Add("[FS] Project=$proj $idx/$($instances.Count) Instance=$instanceShort Tier=$tier Region=$region Zone=$zone provisionedGib=0 provisionedGb=0 (NoShares)")|Out-Null
                $shares += [PSCustomObject]@{
                    Project=$proj
                    InstanceName=$instanceShort
                    ShareName=''
                    Tier=$tier
                    Region=$region
                    Zone=$zone
                    provisionedGb=0
                    provisionedGib=0
                    provisionedTb=0
                    provisionedTib=0
                    CapacityGB=0
                    Networks=$net
                    IPAddresses=$ip
                    State=$state
                    CreateTime=$createTime
                    Labels=$labels
                    Protocol=$protocol
                    Error=''
                }
            }
        }
        $log.Add("[FS-Project-End] $proj ShareRows=$($shares.Count)")|Out-Null
    # Ensure Shares is always an array (even single object)
    $shares = @($shares)
    return [PSCustomObject]@{ Project=$proj; Shares=$shares; Logs=$log; DurationSec=[math]::Round(([DateTime]::UtcNow - $startUtc).TotalSeconds,2) }
    }

    $effectiveMax=[Math]::Min($MaxThreads,[Math]::Max(1,$ProjectIds.Count))
    if (-not $LightMode) {
        $iss=[System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()
        $pool=[RunspaceFactory]::CreateRunspacePool(1,$effectiveMax,$iss,$Host); $pool.Open()
        $runspaces=@(); foreach ($p in $ProjectIds){ $ps=[PowerShell]::Create().AddScript($fsProjectScript).AddArgument($p).AddArgument($MinimalOutput); $ps.RunspacePool=$pool; $runspaces+= [PSCustomObject]@{ PS=$ps; Handle=$ps.BeginInvoke(); Project=$p; Submitted=[DateTime]::UtcNow } }
        $all=@(); $completed=0; $total=$ProjectIds.Count; $overall=[DateTime]::UtcNow
        Write-Progress -Id 501 -Activity 'FileShare Projects' -Status 'Queued...' -PercentComplete 0
        while ($runspaces.Count -gt 0) {
            $next=@()
            foreach ($rs in $runspaces) {
                if ($rs.Handle.IsCompleted) {
                    try { $res=$rs.PS.EndInvoke($rs.Handle) } catch { $res=$null; Write-Host "[FS-ERROR] Project=$($rs.Project) $_" -ForegroundColor Red }
                    $completed++
                    if ($res) {
                        # Force array semantics to avoid single-object collapsing (which breaks .Count checks)
                        $projShares = @($res.Shares)
                        if (-not $projShares -or (@($projShares)).Count -eq 0) {
                            # Fallback: reconstruct from log lines if any successful [FS] entries
                            $recovered=@()
                            foreach ($line in $res.Logs) {
                                if ($line -match '^\[FS\] Project=([^ ]+) .* Instance=([^ ]+) Share=([^ ]+) Tier=([^ ]+) Region=([^ ]+) Zone=([^ ]+) (?:CapacityGB|provisionedGib)=([0-9]+)') {
                                    $gib=[int64]$Matches[7]
                                    $gb=[math]::Round(($gib * 1024 * 1024 * 1024)/1e9,3)
                                    $tib=[math]::Round($gib/1024,4)
                                    $tb=[math]::Round($gb/1000,4)
                                    $recovered += [PSCustomObject]@{
                                        Project=$Matches[1]; InstanceName=$Matches[2]; ShareName=$Matches[3]; Tier=$Matches[4]; Region=$Matches[5]; Zone=$Matches[6]; provisionedGb=$gb; provisionedGib=$gib; provisionedTb=$tb; provisionedTib=$tib; CapacityGB=$gib; Networks=''; IPAddresses=''; State='READY'; CreateTime=''; Labels=''; Protocol='NFS'; Error=''
                                    }
                                }
                            }
                            if ($recovered.Count -gt 0) { Write-Host ("[FS-Recover] Project={0} ReconstructedShares={1}" -f $rs.Project,$recovered.Count) -ForegroundColor Yellow; $projShares=$recovered }
                        }
                        if ($projShares) { $all += $projShares }
                        foreach ($l in $res.Logs){ Write-Host $l -ForegroundColor DarkGray; $allLogs.Add($l) | Out-Null }
                        $shareCt = if ($projShares) { (@($projShares)).Count } else { 0 }
                        Write-Host ("[FS-Project-Done] {0} Shares={1} ElapsedSec={2}" -f $res.Project,$shareCt,$res.DurationSec) -ForegroundColor Green
                    }
                } else {
                    $elapsed = ([DateTime]::UtcNow - $rs.Submitted).TotalSeconds
                    if ($elapsed -ge $TaskTimeoutSec) { try { $rs.PS.Stop() } catch {}; try { $rs.PS.Dispose() } catch {}; $completed++; Write-Host ("[FS-Project-Timeout] Project={0} TimeoutSec={1}" -f $rs.Project,$TaskTimeoutSec) -ForegroundColor Yellow } else { $next += $rs }
                }
            }
            $runspaces=$next
            $pct = if ($total -gt 0){ [math]::Round(($completed/$total)*100,1)} else {100}
            Write-Progress -Id 501 -Activity 'FileShare Projects' -Status ("Projects {0}/{1} ({2}%)" -f $completed,$total,$pct) -PercentComplete $pct
            if ($runspaces.Count -gt 0) { Start-Sleep -Milliseconds 200 }
        }
        Write-Progress -Id 501 -Activity 'FileShare Projects' -Completed
        $pool.Close(); $pool.Dispose()
    $script:FileShareLogLines = $allLogs
        return $all
    } else {
        $queue=New-Object System.Collections.Queue; foreach ($p in $ProjectIds){ $queue.Enqueue($p) }
        $active=@(); $all=@(); $completed=0; $total=$ProjectIds.Count
        Write-Progress -Id 501 -Activity 'FileShare Projects' -Status 'Starting (LightMode)...' -PercentComplete 0
        while ($active.Count -gt 0 -or $queue.Count -gt 0) {
            while ($active.Count -lt $effectiveMax -and $queue.Count -gt 0) {
                $proj=$queue.Dequeue(); $ps=[PowerShell]::Create().AddScript($fsProjectScript).AddArgument($proj).AddArgument($MinimalOutput); $handle=$ps.BeginInvoke(); $active += [PSCustomObject]@{ PS=$ps; Handle=$handle; Project=$proj; Started=[DateTime]::UtcNow }
            }
            $remain=@()
            foreach ($a in $active) {
                if ($a.Handle.IsCompleted) {
                    try { $res=$a.PS.EndInvoke($a.Handle) } catch { $res=$null; Write-Host "[FS-ERROR] Project=$($a.Project) $_" -ForegroundColor Red }
                    $a.PS.Dispose(); $completed++
                    if ($res) {
                        $projShares = @($res.Shares)
                        if (-not $projShares -or (@($projShares)).Count -eq 0) {
                            $recovered=@()
                            foreach ($line in $res.Logs) {
                                if ($line -match '^\[FS\] Project=([^ ]+) .* Instance=([^ ]+) Share=([^ ]+) Tier=([^ ]+) Region=([^ ]+) Zone=([^ ]+) (?:CapacityGB|provisionedGib)=([0-9]+)') {
                                    $gib=[int64]$Matches[7]
                                    $gb=[math]::Round(($gib * 1024 * 1024 * 1024)/1e9,3)
                                    $tib=[math]::Round($gib/1024,4)
                                    $tb=[math]::Round($gb/1000,4)
                                    $recovered += [PSCustomObject]@{
                                        Project=$Matches[1]; InstanceName=$Matches[2]; ShareName=$Matches[3]; Tier=$Matches[4]; Region=$Matches[5]; Zone=$Matches[6]; provisionedGb=$gb; provisionedGib=$gib; provisionedTb=$tb; provisionedTib=$tib; CapacityGB=$gib; Networks=''; IPAddresses=''; State='READY'; CreateTime=''; Labels=''; Protocol='NFS'; Error=''
                                    }
                                }
                            }
                            if ($recovered.Count -gt 0) { Write-Host ("[FS-Recover] Project={0} ReconstructedShares={1}" -f $a.Project,$recovered.Count) -ForegroundColor Yellow; $projShares=$recovered }
                        }
                        if ($projShares) { $all += $projShares }
                        foreach ($l in $res.Logs){ Write-Host $l -ForegroundColor DarkGray; $allLogs.Add($l) | Out-Null }
                        $shareCt = if ($projShares) { (@($projShares)).Count } else { 0 }
                        Write-Host ("[FS-Project-Done] {0} Shares={1} ElapsedSec={2}" -f $res.Project,$shareCt,$res.DurationSec) -ForegroundColor Green
                    }
                } else {
                    $elapsed=([DateTime]::UtcNow - $a.Started).TotalSeconds
                    if ($elapsed -ge $TaskTimeoutSec) { try { $a.PS.Stop() } catch {}; try { $a.PS.Dispose() } catch {}; $completed++; Write-Host ("[FS-Project-Timeout] Project={0} TimeoutSec={1}" -f $a.Project,$TaskTimeoutSec) -ForegroundColor Yellow } else { $remain += $a }
                }
            }
            $active=$remain
            $pct = if ($total -gt 0){ [math]::Round(($completed/$total)*100,1)} else {100}
            Write-Progress -Id 501 -Activity 'FileShare Projects' -Status ("(Light) Projects {0}/{1} ({2}%)" -f $completed,$total,$pct) -PercentComplete $pct
            if ($active.Count -gt 0 -or $queue.Count -gt 0) { Start-Sleep -Milliseconds 160 }
        }
        Write-Progress -Id 501 -Activity 'FileShare Projects' -Completed
    $script:FileShareLogLines = $allLogs
        return $all
    }
}

<# =============================
   Modular DB Inventory Helpers
   (Refactor of original monolithic logic; behavior preserved)
   Each helper returns an array of PSCustomObject items OR $null.
   They update status & error refs just like original inline logic.
============================= #>

function Get-AlloyPrimaryClusterStorageBytes {
    param(
        [string]$ProjectId,
        [string]$ClusterId,
        [string]$Region,
        [string]$AccessToken
    )
    $ProjectId = ($ProjectId|ForEach-Object{$_}).Trim(); $ClusterId = ($ClusterId|ForEach-Object{$_}).Trim(); $Region = ($Region|ForEach-Object{$_}).Trim()
    if (-not $ProjectId -or -not $ClusterId -or -not $Region) { return -4 }
    try { if (-not $AccessToken) { $AccessToken = (gcloud auth print-access-token 2>$null).Trim() } } catch { return -1 }
    if (-not $AccessToken) { return -1 }
    $metricType = 'alloydb.googleapis.com/cluster/storage/usage'
    $now=[DateTime]::UtcNow; $endTime=$now.ToString('yyyy-MM-ddTHH:mm:ss.fffZ'); $startTime=$now.AddMinutes(-10).ToString('yyyy-MM-ddTHH:mm:ss.fffZ')
    $filter='metric.type="{0}" AND resource.label.cluster_id="{1}" AND resource.label.location="{2}"' -f $metricType,$ClusterId,$Region
    try { $enc=[System.Uri]::EscapeDataString($filter) } catch { return -3 }
    $url = 'https://monitoring.googleapis.com/v3/projects/{0}/timeSeries?filter={1}&interval.startTime={2}&interval.endTime={3}' -f $ProjectId,$enc,$startTime,$endTime
    try { $resp=Invoke-RestMethod -Uri $url -Headers @{ Authorization = "Bearer $AccessToken" } -Method GET -ErrorAction Stop } catch { return -3 }
    if ($resp.timeSeries -and $resp.timeSeries.Count -gt 0) {
        $series=$resp.timeSeries[0]
        $latest=($series.points | Sort-Object { $_.interval.endTime } | Select-Object -Last 1)
        if ($latest -and $latest.value) {
            $v=0; if ($latest.value.int64Value){$v=[double]$latest.value.int64Value} elseif ($latest.value.doubleValue){$v=[double][math]::Round($latest.value.doubleValue,0)}
            return $v
        }
        return -2
    }
    return -2
}

function Invoke-DbAlloyListInternal {
    param(
        [string]$ProjectId,
        [string]$Sub
    )
    $ga = & gcloud alloydb $Sub list --project $ProjectId --quiet --format=json 2>&1
    if ($LASTEXITCODE -eq 0 -and $ga -and -not ($ga -match 'not have this command group')) { return @{Ok=$true;Json=$ga;Msg=''} }
    $beta = & gcloud beta alloydb $Sub list --project $ProjectId --quiet --format=json 2>&1
    if ($LASTEXITCODE -eq 0 -and $beta -and -not ($beta -match 'not have this command group')) { return @{Ok=$true;Json=$beta;Msg=''} }
    return @{Ok=$false;Json='';Msg=($ga|Out-String)+($beta|Out-String)}
}

function Get-GcpCloudSqlInstances {
    param(
        [string]$ProjectId,
        [ref]$Errors,
        [ref]$Status
    )
    Write-Log "[DB] ($ProjectId) Start CloudSQL" -Level DEBUG
    $items=@()
    $rawSql = & gcloud sql instances list --project $ProjectId --quiet --format=json 2>&1
    if ($LASTEXITCODE -eq 0 -and $rawSql) {
        $sqlObjs = $null; try { $sqlObjs = $rawSql | ConvertFrom-Json } catch {}
        $cloudSqlInstCount=0
        foreach ($si in ($sqlObjs | Where-Object { $_ })) {
            $cloudSqlInstCount++
            $tier = $si.settings.tier; if (-not $tier) { $tier='' }
            $state = $si.state
            $region = $si.region
            $zone = if ($si.gceZone) { ($si.gceZone -split '/')[-1] } else { 'Regional' }
            $diskSizeGb = $si.settings.dataDiskSizeGb; if (-not $diskSizeGb) { $diskSizeGb = 0 }
            $engine = $si.databaseVersion
            $privateIPs=@(); $publicIPs=@(); $outgoingIPs=@()
            if ($si.ipAddresses) {
                foreach ($ipa in $si.ipAddresses) {
                    $ipVal = $ipa.ipAddress; if (-not $ipVal) { continue }
                    if ($ipa.type -eq 'PRIMARY') { $publicIPs += $ipVal } elseif ($ipa.type -eq 'OUTGOING') { $outgoingIPs += $ipVal } else { $privateIPs += $ipVal }
                }
            }
            $items += [PSCustomObject]@{Type='CloudSQL';Project=$ProjectId;Name=$si.name;Region=$region;Zone=$zone;Engine=$engine;TierOrCapacity=$tier;State=$state;StorageGB=[double]$diskSizeGb;PrivateIPs=($privateIPs -join ';');PublicIPs=($publicIPs -join ';');OutgoingIPs=($outgoingIPs -join ';');Extra='';Error=''}
            if ($sqlObjs -and $sqlObjs.Count -gt 0 -and $cloudSqlInstCount -le 200) {
                $pctInst=[int](($cloudSqlInstCount/[double]$sqlObjs.Count)*100); if ($pctInst -gt 100){$pctInst=100}
                Write-Progress -Id 611 -Activity ("CloudSQL $ProjectId") -Status ("Instances ($cloudSqlInstCount/$($sqlObjs.Count)) - $pctInst%") -PercentComplete $pctInst -ErrorAction SilentlyContinue
            }
        }
        Write-Log "[DB] ($ProjectId) End CloudSQL Count=$cloudSqlInstCount" -Level DEBUG
    } else {
        $errText = ($rawSql | Out-String)
        if ($errText -match '(?i)permission|denied|forbidden|403') { $Errors.Value.Add("CloudSQL permission issue for $ProjectId"); $Status.Value.PermissionIssue='Y' }
        elseif ($errText -match '(?i)not enabled|has not been used|api .* enable') { $Errors.Value.Add("CloudSQL API disabled for $ProjectId"); $Status.Value.ApiDisabled='Y' }
        elseif ($errText -and $errText.Trim()) { $Errors.Value.Add("CloudSQL list failed for $ProjectId"); if (-not $Status.Value.Error){ $Status.Value.Error='CloudSQL list failed' } }
    }
    return $items
}

function Get-GcpAlloyDbResources {
    param(
        [string]$ProjectId,
        [ref]$Errors,
        [ref]$Status
    )
    $items=@()
    Write-Log "[DB] ($ProjectId) Start AlloyDB Clusters" -Level DEBUG
    $alloyClustersCall = Invoke-DbAlloyListInternal -ProjectId $ProjectId -Sub 'clusters'
    $alloyComponentMissing = $false
    if (-not $alloyClustersCall.Ok -and $alloyClustersCall.Msg -match 'requires the installation of components') { $alloyComponentMissing = $true }
    if (-not $alloyComponentMissing -and $alloyClustersCall.Ok) {
        $clusters = $null; try { $clusters = $alloyClustersCall.Json | ConvertFrom-Json } catch {}
        $__primaryClusters = @($clusters | Where-Object { $_ -and $_.clusterType -and ($_.clusterType.ToString().ToUpper() -eq 'PRIMARY') })
        $___primTotal = if ($__primaryClusters) { $__primaryClusters.Count } else { 0 }
        $___primProcessed = 0
        foreach ($cl in ($clusters | Where-Object { $_ })) {
            $clusterRegion = $cl.locationId
            if ($cl.name -and $cl.name -match '/locations/([^/]+)/') { $clusterRegion = $Matches[1] }
            $clusterType = if ($cl.clusterType) { $cl.clusterType } elseif ($cl.instanceConfig) { $cl.instanceConfig } else { '' }
            $dbVersion = if ($cl.databaseVersion) { $cl.databaseVersion } else { '' }
            $shortName = if ($cl.name) { ($cl.name -split '/')[-1] } else { '' }
            $engineVal = if ($dbVersion) { "PostgreSQL-$dbVersion" } else { 'PostgreSQL' }
            if ($clusterType -and ($clusterType.ToString().ToUpper() -ne 'PRIMARY')) { continue }
            $___primProcessed++
            if ($___primTotal -gt 0 -and $___primProcessed -le 500) {
                $pctPrim=[int](($___primProcessed/[double]$___primTotal)*100); if ($pctPrim -gt 100) { $pctPrim=100 }
                Write-Progress -Id 613 -Activity ("AlloyDB Clusters $ProjectId") -Status ("PrimaryClusters ($___primProcessed/$___primTotal) - $pctPrim%") -PercentComplete $pctPrim -ErrorAction SilentlyContinue
            }
            $clusterStorageBytes = 0; $clusterStorageGB=0; $clusterStorageMB=0; $extraInfo=''
            if ($clusterType -eq 'PRIMARY') {
                if (-not $script:__AlloyAccessToken) { try { $script:__AlloyAccessToken = (gcloud auth print-access-token 2>$null).Trim() } catch { $script:__AlloyAccessToken=$null } }
                Write-Log -Level INFO -Message ("[DB][AlloyDB] ({0}) Cluster={1} Region={2} StartingMonitoringLookup" -f $ProjectId,$shortName,$clusterRegion)
                $bytes = Get-AlloyPrimaryClusterStorageBytes -ProjectId $ProjectId -ClusterId $shortName -Region $clusterRegion -AccessToken $script:__AlloyAccessToken
                $monitorStatus='OK'
                if ($bytes -and $bytes -gt 0) { $clusterStorageBytes=$bytes; $clusterStorageGB=[double][math]::Round($bytes/1GB,0); $clusterStorageMB=[double][math]::Round($bytes/1MB,0); $extraInfo=''; $monitorStatus='OK' }
                elseif ($bytes -eq -2) { $extraInfo='Monitoring=NoData';        $monitorStatus='NoData' }
                elseif ($bytes -eq -3) { $extraInfo='Monitoring=Error';         $monitorStatus='Error' }
                elseif ($bytes -eq -1) { $extraInfo='Monitoring=AuthError';     $monitorStatus='AuthError' }
                elseif ($bytes -eq -4) { $extraInfo='Monitoring=BadParams';     $monitorStatus='BadParams' }
                elseif ($bytes -eq -5) { $extraInfo='Monitoring=URLBuildError'; $monitorStatus='URLBuildError' }
                else { $extraInfo='Monitoring=Unknown';                         $monitorStatus='Unknown' }
                Write-Log -Level INFO -Message ("[DB][AlloyDB] ({0}) Cluster={1} Region={2} MonitoringResult Status={3} StorageBytes={4}" -f $ProjectId,$shortName,$clusterRegion,$monitorStatus,$clusterStorageBytes)
            }
            $items += [PSCustomObject]@{Type='alloydbcluster';Project=$ProjectId;Name=$shortName;Region=$clusterRegion;Zone='N/A';Engine=$engineVal;TierOrCapacity=$clusterType;State=$cl.state;StorageBytes=[int64]$clusterStorageBytes;StorageMB=[double]$clusterStorageMB;StorageGB=[double]$clusterStorageGB;MonitoringStatus=$monitorStatus;PrivateIPs='';PublicIPs='';Extra=$extraInfo;Error=''}
        }
    Write-Log "[DB] ($ProjectId) AlloyDB instance enumeration skipped (primary clusters only)" -Level INFO
    } elseif (-not $alloyComponentMissing -and $alloyClustersCall.Msg -match 'permission|denied|forbidden|403') {
        $Errors.Value.Add("AlloyDB clusters permission issue for $ProjectId")
        if (-not $Status.Value.PermissionIssue) { $Status.Value.PermissionIssue='Y' }
    } elseif ($alloyComponentMissing) {
        # silent skip
    }
    Write-Log "[DB] ($ProjectId) End AlloyDB Clusters" -Level DEBUG
    return $items
}

function Get-SpannerDatabaseStorageBytes {
    <#
        Returns an object with database used bytes plus (optionally) instance used bytes & derived limit.
        Backwards compatibility: if only a numeric value is expected by callers, they can still treat
        .DatabaseUsedBytes (or cast the return to [double] if they adapt). Current script only used the
        numeric return for database size lookups; we extend without breaking existing logic.
        Error Codes (numeric legacy mode):
          -1 Auth error, -2 No data, -3 API/URL build error, -4 Bad params
    #>
    param(
        [string]$ProjectId,
        [string]$InstanceId,
        [string]$DatabaseId,
        [string]$AccessToken,
        [switch]$IncludeInstanceMetrics,
        [int]$InstanceAlignSeconds = 3600,
        [int]$InstanceLookbackHours = 1
    )

    $ProjectId  = ($ProjectId  | ForEach-Object { $_ }).Trim()
    $InstanceId = ($InstanceId | ForEach-Object { $_ }).Trim()
    $DatabaseId = ($DatabaseId | ForEach-Object { $_ }).Trim()
    if (-not $ProjectId -or -not $InstanceId -or -not $DatabaseId) { return -4 }

    try { if (-not $AccessToken) { $AccessToken = (gcloud auth print-access-token 2>$null).Trim() } } catch { return -1 }
    if (-not $AccessToken) { return -1 }

    $now = [DateTime]::UtcNow
    $endTime = $now.ToString('yyyy-MM-ddTHH:mm:ss.fffZ')
    $startTime = $now.AddMinutes(-10).ToString('yyyy-MM-ddTHH:mm:ss.fffZ')
    $metricType = 'spanner.googleapis.com/database/storage/total_bytes'
    $filter = 'metric.type="{0}" AND resource.label.instance_id="{1}" AND resource.label.database_id="{2}"' -f $metricType,$InstanceId,$DatabaseId
    try { $enc = [System.Uri]::EscapeDataString($filter) } catch { return -3 }
    $url = 'https://monitoring.googleapis.com/v3/projects/{0}/timeSeries?filter={1}&interval.startTime={2}&interval.endTime={3}' -f $ProjectId,$enc,$startTime,$endTime
    $dbUsedBytes = -2
    try {
        $resp = Invoke-RestMethod -Uri $url -Headers @{ Authorization = "Bearer $AccessToken" } -Method GET -ErrorAction Stop
        if ($resp.timeSeries -and $resp.timeSeries.Count -gt 0) {
            $series = $resp.timeSeries[0]
            $latest = ($series.points | Sort-Object { $_.interval.endTime } | Select-Object -Last 1)
            if ($latest -and $latest.value) {
                if ($latest.value.int64Value) { $dbUsedBytes = [int64]$latest.value.int64Value }
                elseif ($latest.value.doubleValue) { $dbUsedBytes = [int64][math]::Round($latest.value.doubleValue,0) }
            }
        }
    } catch { return -3 }

    $instanceUsedBytes = $null
    $derivedLimitBytes = $null
    $utilPct = $null
    if ($IncludeInstanceMetrics) {
        if ($InstanceAlignSeconds -le 0) { $InstanceAlignSeconds = 300 }
        $instStart = $now.AddHours(-[double]$InstanceLookbackHours)
        $instFilter = 'metric.type="spanner.googleapis.com/instance/storage/used_bytes" AND resource.type="spanner_instance" AND resource.label.instance_id="{0}"' -f $InstanceId
        try { $instEnc = [System.Uri]::EscapeDataString($instFilter) } catch { $instEnc = $null }
        if ($instEnc) {
            $instUrl = "https://monitoring.googleapis.com/v3/projects/$ProjectId/timeSeries?filter=$instEnc&interval.endTime=$endTime&interval.startTime=$($instStart.ToString('yyyy-MM-ddTHH:mm:ss.fffZ'))&view=FULL&aggregation.alignmentPeriod=${InstanceAlignSeconds}s&aggregation.perSeriesAligner=ALIGN_MAX"
            try {
                $instResp = Invoke-RestMethod -Uri $instUrl -Headers @{ Authorization = "Bearer $AccessToken" } -Method GET -ErrorAction Stop
                if ($instResp.timeSeries -and $instResp.timeSeries.Count -gt 0) {
                    $latestSeries = $instResp.timeSeries | ForEach-Object {
                        $p = ($_.points | Sort-Object { $_.interval.endTime } | Select-Object -Last 1)
                        [pscustomobject]@{ Series = $_; LatestPoint = $p }
                    } | Sort-Object { $_.LatestPoint.interval.endTime } | Select-Object -Last 1
                    if ($latestSeries -and $latestSeries.LatestPoint -and $latestSeries.LatestPoint.value) {
                        $val = $latestSeries.LatestPoint.value
                        if ($val.int64Value) { $instanceUsedBytes = [int64]$val.int64Value }
                        elseif ($val.doubleValue) { $instanceUsedBytes = [int64][math]::Round($val.doubleValue,0) }
                    }
                }
            } catch { }
        }
        # Derive limit via processing units or nodeCount
        try {
            $instApi = Invoke-RestMethod -Uri "https://spanner.googleapis.com/v1/projects/$ProjectId/instances/$InstanceId" -Headers @{ Authorization = "Bearer $AccessToken" } -ErrorAction Stop
            $processingUnits = 0
            if ($instApi.processingUnits) { $processingUnits = [int]$instApi.processingUnits }
            elseif ($instApi.nodeCount) { $processingUnits = [int]$instApi.nodeCount * 100 }
            if ($processingUnits -gt 0) {
                $limitPer100PU = [int64]([Math]::Pow(1024,4)) # 1 TiB per 100 PU
                $derivedLimitBytes = [int64]($limitPer100PU * ($processingUnits / 100.0))
            }
            if ($derivedLimitBytes -and $instanceUsedBytes -ne $null -and $derivedLimitBytes -gt 0) {
                $utilPct = [math]::Round(($instanceUsedBytes / $derivedLimitBytes) * 100,4)
            }
        } catch { }
    }

    return [pscustomobject]@{
        DatabaseUsedBytes = $dbUsedBytes
        InstanceUsedBytes = $instanceUsedBytes
        DerivedLimitBytes = $derivedLimitBytes
        UtilizationPercent = $utilPct
    }
}

function Get-GcpSpannerInstances {
    param([string]$ProjectId)
    Write-Log "[DB] ($ProjectId) Start Spanner" -Level DEBUG
    $items=@()
    # Capture both stdout and stderr so we can emit diagnostics on failures
    $rawOutput = $null
    try {
        $rawOutput = & gcloud spanner instances list --project $ProjectId --quiet --format=json 2>&1
    } catch {
        Write-Log "[DB] ($ProjectId) Spanner command threw exception: $($_.Exception.Message)" -Level WARN
    }
    $exit = $LASTEXITCODE
    if ($exit -ne 0) {
        Write-Log "[DB] ($ProjectId) Spanner list failed ExitCode=$exit OutputSnippet='$(($rawOutput | Out-String).Trim() | Select-Object -First 1)'" -Level WARN
    } elseif (-not $rawOutput -or ($rawOutput.Trim() -eq '') -or ($rawOutput.Trim() -eq '[]')) {
        Write-Log "[DB] ($ProjectId) Spanner returned no instances (empty list)." -Level INFO
    } else {
        $spInstances = $null; $convertOk=$true
        try { $spInstances = $rawOutput | ConvertFrom-Json } catch { $convertOk=$false; Write-Log "[DB] ($ProjectId) Spanner JSON parse failure: $($_.Exception.Message)" -Level WARN }
        if ($convertOk -and $spInstances) {
            $count = (@($spInstances)).Count
            Write-Log "[DB] ($ProjectId) Spanner instances retrieved: $count" -Level INFO
            foreach ($spi in ($spInstances | Where-Object { $_ })) {
                $config = $spi.config
                $regionParsed = $config; if ($config -match 'instanceConfigs/(.+)$') { $regionParsed = $Matches[1] }
                $nodes = $spi.nodeCount; $processingUnits=$spi.processingUnits
                $tier = if ($nodes) { "$nodes nodes" } elseif ($processingUnits) { "$processingUnits PU" } else { '' }
                $derivedLimitBytes = $null
                if ($processingUnits) {
                    $derivedLimitBytes = [int64]([Math]::Pow(1024,4) * ($processingUnits / 100.0))
                } elseif ($nodes) {
                    $derivedLimitBytes = [int64]([Math]::Pow(1024,4) * $nodes) # assume 1TiB per node
                }
                $derivedLimitGiB = if ($derivedLimitBytes) { [math]::Round($derivedLimitBytes / 1GB,4) } else { 0 }
                $instMetrics = $null
                try { $instMetrics = Get-SpannerDatabaseStorageBytes -ProjectId $ProjectId -InstanceId $spi.name -DatabaseId 'dummy' -IncludeInstanceMetrics -ErrorAction SilentlyContinue } catch { }
                $usedGiB = 0; $utilPct = $null
                if ($instMetrics -and $instMetrics.InstanceUsedBytes) { $usedGiB = [math]::Round($instMetrics.InstanceUsedBytes / 1GB,4) }
                if ($instMetrics -and $instMetrics.UtilizationPercent -ne $null) { $utilPct = $instMetrics.UtilizationPercent }
                $extra = ''
                if ($usedGiB -gt 0 -or $utilPct -ne $null) {
                    $extra = "InstanceUsedGiB=$usedGiB;DerivedLimitGiB=$derivedLimitGiB;UtilPercent=$utilPct"
                }
                $items += [PSCustomObject]@{Type='SpannerInstance';Project=$ProjectId;Parent='';Name=$spi.name;Region=$regionParsed;Zone='';Engine='Spanner';TierOrCapacity=$tier;State=$spi.state;StorageGB=[double]$derivedLimitGiB;PrivateIPs='';PublicIPs='';Error='';Extra=$extra}
            }
        }
    }
    Write-Log "[DB] ($ProjectId) End Spanner" -Level DEBUG
    return $items
}

function Get-GcpBigQueryDatasets {
    param(
        [string]$ProjectId,
        [ref]$Errors,
        [ref]$Status
    )

    Write-Log "[DB] ($ProjectId) Start BigQuery" -Level DEBUG
    $items = @()
    $bqCmdPresent = $true

    if (-not (Get-Command bq -ErrorAction SilentlyContinue)) {
        $bqCmdPresent = $false
    }

    if (-not $bqCmdPresent) {
        $Errors.Value.Add("bq CLI not found for $ProjectId (skipping BigQuery)")
        if (-not $Status.Value.Error) { $Status.Value.Error = 'bq-cli-missing' }
        return $items
    }

    $bqJson = & bq --project_id=$ProjectId ls --format=json 2>$null

    if ($LASTEXITCODE -ne 0 -or -not $bqJson) {
        $Errors.Value.Add("Failed to list BigQuery datasets for $ProjectId")
        return $items
    }

    $bqDatasets = $null
    try { $bqDatasets = $bqJson | ConvertFrom-Json } catch {}

    $datasetList = @($bqDatasets | Where-Object { $_ })
    $datasetTotal = $datasetList.Count
    Write-Log ("[DB][BigQuery] ({0}) DatasetCount={1}" -f $ProjectId,$datasetTotal) -Level DEBUG
    $dsIndex = 0

    foreach ($ds in $datasetList) {
        $dsIndex++
        $datasetId = $ds.datasetReference.datasetId
        $region = $ds.location
        $totalLogicalBytes = 0
        $totalPhysicalBytes = 0
        $largestPhysicalBytes = 0
        $largestLogicalBytesForLargestPhysical = 0
        $largestTableName = ''
        $pctDs = if ($datasetTotal -gt 0) { [int](($dsIndex/[double]$datasetTotal)*100) } else { 100 }
        Write-Progress -Id 614 -ParentId 610 -Activity ("BigQuery Datasets $ProjectId") -Status ("Dataset {0}/{1} ({2}%) - {3}" -f $dsIndex,$datasetTotal,$pctDs,$datasetId) -PercentComplete $pctDs -ErrorAction SilentlyContinue
        Write-Log ("[DB][BigQuery] ({0}) DatasetStart {1} ({2}/{3})" -f $ProjectId,$datasetId,$dsIndex,$datasetTotal) -Level DEBUG

        # List tables in dataset (handle both array output and object with 'tables' + pagination)
        $tables = New-Object System.Collections.Generic.List[object]
        $pageToken = $null; $listAttempt = 0; $fetchedPages = 0
        do {
            $listAttempt++
            # Build bq ls command safely (avoid quoting issues and Invoke-Expression)
            $cmdArgs = @("--project_id=$ProjectId","ls","--dataset_id=$datasetId","-n","100000","--format=json")
            if ($pageToken) { $cmdArgs += "--page_token=$pageToken" }
            $tableListJson = & bq @cmdArgs 2>$null
            if ($LASTEXITCODE -ne 0 -or -not $tableListJson) { break }
            try {
                $parsed = $tableListJson | ConvertFrom-Json
                if ($parsed -is [System.Collections.IEnumerable] -and -not ($parsed.PSObject.Properties['tables'])) {
                    foreach ($t in ($parsed | Where-Object { $_ })) { $tables.Add($t) }
                    $pageToken = $null
                } elseif ($parsed.PSObject.Properties['tables']) {
                    foreach ($t in (@($parsed.tables) | Where-Object { $_ })) { $tables.Add($t) }
                    if ($parsed.PSObject.Properties['nextPageToken'] -and $parsed.nextPageToken) { $pageToken = $parsed.nextPageToken } else { $pageToken = $null }
                } else {
                    $tables.Add($parsed) | Out-Null; $pageToken = $null
                }
                $fetchedPages++
            } catch {
                $Errors.Value.Add(("Failed to parse table list JSON for {0}:{1} (attempt {2})" -f $ProjectId,$datasetId,$listAttempt))
                break
            }
        } while ($pageToken)
        if ($tables.Count -eq 0 -and $LASTEXITCODE -ne 0) { $Errors.Value.Add(("Failed to list tables in dataset {0}:{1}" -f $ProjectId,$datasetId)) }
        $tableTotal = $tables.Count
        Write-Log ("[DB][BigQuery] ({0}) Dataset={1} TableListPages={2} RawCount={3}" -f $ProjectId,$datasetId,$fetchedPages,$tableTotal) -Level DEBUG
        $tblIndex = 0
    if ($tableTotal -gt 0) { Write-Log ("[DB][BigQuery] ({0}) Dataset={1} TableCount={2}" -f $ProjectId,$datasetId,$tableTotal) -Level DEBUG }

    # --- Lightweight table metadata retrieval (default parallel ON, capped low) ---
    # Strategy:
    #   1. Try bulk INFORMATION_SCHEMA for large datasets (threshold configurable)
    #   2. If bulk fails or below threshold, fetch per-table metadata with at most $MaxBigQueryTableThreads active
    #   3. Adaptive throttle if free RAM < 1.5 GB (drop to 2 threads)
    # Parallel is ENABLED by default (can be turned off via -DisableBigQueryParallelTables)
    $tableResults = @()
    $maxThreads = if (Get-Variable -Name MaxBigQueryTableThreads -Scope Script -ErrorAction SilentlyContinue) { [int]$script:MaxBigQueryTableThreads } else { 5 }
    if ($maxThreads -lt 1) { $maxThreads = 1 }
    if ($tableTotal -lt $maxThreads) { $maxThreads = $tableTotal }

    # Determine if parallel table metadata retrieval is enabled.
    # New default: ON (previously OFF unless EnableBigQueryParallelTables was explicitly set)
    # Override precedence:
    #   1. $script:DisableBigQueryParallelTables = $true  -> force OFF
    #   2. $script:EnableBigQueryParallelTables  (legacy flag) respected if present
    $parallelEnabled = $true
    if (Get-Variable -Name EnableBigQueryParallelTables -Scope Script -ErrorAction SilentlyContinue) {
        try { $parallelEnabled = [bool]$script:EnableBigQueryParallelTables } catch { $parallelEnabled = $true }
    }
    if (Get-Variable -Name DisableBigQueryParallelTables -Scope Script -ErrorAction SilentlyContinue) {
        try { if ([bool]$script:DisableBigQueryParallelTables) { $parallelEnabled = $false } } catch {}
    }
    $useParallel = $parallelEnabled -and ($tableTotal -gt 0) -and ($maxThreads -gt 1) -and (-not $bulkSuccess)
    if ($useParallel) {
        Write-Log ("[DB][BigQuery] ({0}) Dataset={1} LightParallel Threads={2} Tables={3}" -f $ProjectId,$datasetId,$maxThreads,$tableTotal) -Level DEBUG
    } else {
        Write-Log ("[DB][BigQuery] ({0}) Dataset={1} SequentialTableMetadata" -f $ProjectId,$datasetId) -Level DEBUG
    }
        $bulkAttempted = $false; $bulkSuccess = $false
        # Try bulk INFORMATION_SCHEMA query if large dataset and not disabled
        $bulkThreshold = if (Get-Variable -Name BigQueryBulkQueryThreshold -Scope Script -ErrorAction SilentlyContinue) { [int]$script:BigQueryBulkQueryThreshold } else { 250 }
        $bulkDisabled  = $false; try { $bulkDisabled = [bool]$script:DisableBigQueryBulkQuery } catch {}
        if (-not $bulkDisabled -and $tableTotal -ge $bulkThreshold) {
            $bulkAttempted = $true
            Write-Log ("[DB][BigQuery] ({0}) Dataset={1} BulkQueryAttempt Tables={2}" -f $ProjectId,$datasetId,$tableTotal) -Level DEBUG
            # Use bq query --nouse_legacy_sql to pull logical & physical bytes per table.
            $qualified = "$ProjectId.$datasetId"
            # Build single INFORMATION_SCHEMA query (avoid fragile escaping of backticks inside interpolated string)
            $sql = 'SELECT table_id AS TableId, total_physical_bytes AS PhysicalBytes, total_logical_bytes AS LogicalBytes FROM `' + $qualified + '.INFORMATION_SCHEMA.TABLE_STORAGE`'
            $bulkJson = & bq --project_id=$ProjectId query --nouse_legacy_sql --format=json --max_rows=100000 "$sql" 2>$null
            if ($LASTEXITCODE -eq 0 -and $bulkJson) {
                try {
                    $parsedBulk = $bulkJson | ConvertFrom-Json
                    # Rows may be objects with fields table_id, total_physical_bytes, total_logical_bytes
                    foreach ($row in ($parsedBulk | Where-Object { $_ })) {
                        $tid = $row.TableId
                        if (-not $tid -and $row.PSObject.Properties['table_id']) { $tid = $row.table_id }
                        if ($tid) {
                            $logical = 0; $physical=0
                            if ($row.PSObject.Properties['LogicalBytes']) { $logical = [int64]$row.LogicalBytes }
                            elseif ($row.PSObject.Properties['total_logical_bytes']) { $logical = [int64]$row.total_logical_bytes }
                            if ($row.PSObject.Properties['PhysicalBytes']) { $physical = [int64]$row.PhysicalBytes }
                            elseif ($row.PSObject.Properties['total_physical_bytes']) { $physical = [int64]$row.total_physical_bytes }
                            $tableResults += [PSCustomObject]@{TableId=$tid; LogicalBytes=$logical; PhysicalBytes=$physical; Error=''}
                        }
                    }
                    if ($tableResults.Count -gt 0) {
                        $bulkSuccess = $true
                        Write-Log ("[DB][BigQuery] ({0}) Dataset={1} BulkQuerySuccess Rows={2}" -f $ProjectId,$datasetId,$tableResults.Count) -Level DEBUG
                    } else {
                        Write-Log ("[DB][BigQuery] ({0}) Dataset={1} BulkQueryEmpty FallingBack" -f $ProjectId,$datasetId) -Level WARN
                        $tableResults = @()
                    }
                } catch {
                    Write-Log ("[DB][BigQuery] ({0}) Dataset={1} BulkQueryParseFailed FallingBack" -f $ProjectId,$datasetId) -Level WARN
                    $tableResults = @()
                }
            } else {
                Write-Log ("[DB][BigQuery] ({0}) Dataset={1} BulkQueryFailed ExitCode={2} FallingBack" -f $ProjectId,$datasetId,$LASTEXITCODE) -Level WARN
            }
        }
        if ($bulkSuccess) { $useParallel = $false } # Skip per-table lookups entirely if bulk worked

        # Build tableIds list (handles mixed object shapes) only if per-table needed
        if (-not $bulkSuccess) {
            $tableIds = @()
            foreach ($t in $tables) {
                if ($t.tableReference -and $t.tableReference.tableId) { $tableIds += $t.tableReference.tableId }
                elseif ($t.tableId) { $tableIds += $t.tableId }
            }
        }

        if (-not $bulkSuccess -and $useParallel) {
            try {
                $iss = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()
                $pool = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspacePool(1,$maxThreads,$iss,$Host); $pool.Open()
                $active = @(); $idx = 0; $completed = 0
                while ($idx -lt $tableIds.Count -or $active.Count -gt 0) {
                    # Adaptive free memory check (drop to 2 threads if <1.5GB)
                    try {
                        $os = Get-CimInstance Win32_OperatingSystem -ErrorAction Stop
                        $freeGB = [math]::Round(($os.FreePhysicalMemory*1KB)/1GB,2)
                        if ($freeGB -lt 1.5 -and $maxThreads -gt 2) { $maxThreads = 2 }
                    } catch {}
                    while ($active.Count -lt $maxThreads -and $idx -lt $tableIds.Count) {
                        $tid = $tableIds[$idx]; $idx++
                        $ps = [PowerShell]::Create(); $ps.RunspacePool = $pool
                        [void]$ps.AddScript({
                            param($p,$ds,$t)
                            $r = [PSCustomObject]@{TableId=$t; LogicalBytes=[int64]0; PhysicalBytes=[int64]0; Error=''}
                            $j = & bq show --format=json --project_id=$p "$ds.$t" 2>$null
                            if ($LASTEXITCODE -eq 0 -and $j) {
                                try {
                                    $o = $j | ConvertFrom-Json
                                    if ($o.PSObject.Properties['numTotalLogicalBytes']) { $r.LogicalBytes = [int64]$o.numTotalLogicalBytes }
                                    elseif ($o.PSObject.Properties['numBytes']) { $r.LogicalBytes = [int64]$o.numBytes }
                                    if ($o.PSObject.Properties['numTotalPhysicalBytes']) { $r.PhysicalBytes = [int64]$o.numTotalPhysicalBytes }
                                } catch { $r.Error='parse' }
                            } else { $r.Error='show' }
                            return $r
                        }).AddArgument($ProjectId).AddArgument($datasetId).AddArgument($tid)
                        $active += [PSCustomObject]@{ PS=$ps; Handle=$ps.BeginInvoke(); Table=$tid }
                    }
                    $next=@()
                    foreach ($a in $active) {
                        if ($a.Handle.IsCompleted) {
                            try { $out = $a.PS.EndInvoke($a.Handle) } catch { $out=$null }
                            try { $a.PS.Dispose() } catch {}
                            if ($out) { $tableResults += $out }
                            $completed++
                            if ($tableIds.Count -gt 0 -and ($completed -eq 1 -or $completed -eq $tableIds.Count -or ($completed % 25 -eq 0))) {
                                $pct = [int](($completed/[double]$tableIds.Count)*100)
                                Write-Progress -Id 615 -ParentId 614 -Activity ("Tables $datasetId") -Status ("{0}/{1} ({2}%)" -f $completed,$tableIds.Count,$pct) -PercentComplete $pct -ErrorAction SilentlyContinue
                            }
                        } else { $next += $a }
                    }
                    $active = $next
                    if ($active.Count -gt 0 -and $idx -lt $tableIds.Count) { Start-Sleep -Milliseconds 50 }
                }
                $pool.Close(); $pool.Dispose()
            } catch {
                Write-Log ("[DB][BigQuery] ({0}) Dataset={1} LightParallelFailed -> Sequential: {2}" -f $ProjectId,$datasetId,$_.Exception.Message) -Level WARN
                $useParallel = $false
            }
        }

        if (-not $bulkSuccess -and -not $useParallel) {
            $tblIndex = 0
            foreach ($tid in $tableIds) {
                $tblIndex++
                if ($tableIds.Count -gt 0) {
                    $pctTbl = [int](($tblIndex/[double]$tableIds.Count)*100)
                    if ($tblIndex -eq 1 -or $tblIndex -eq $tableIds.Count -or ($tblIndex % 25 -eq 0)) {
                        Write-Progress -Id 615 -ParentId 614 -Activity ("Tables $datasetId") -Status ("{0}/{1} ({2}%)" -f $tblIndex,$tableIds.Count,$pctTbl) -PercentComplete $pctTbl -ErrorAction SilentlyContinue
                    }
                }
                $metaJson = & bq show --format=json --project_id=$ProjectId "$datasetId.$tid" 2>$null
                $obj = [PSCustomObject]@{TableId=$tid; LogicalBytes=[int64]0; PhysicalBytes=[int64]0; Error=''}
                if ($LASTEXITCODE -eq 0 -and $metaJson) {
                    try {
                        $tm = $metaJson | ConvertFrom-Json
                        if ($tm.PSObject.Properties['numTotalLogicalBytes']) { $obj.LogicalBytes = [int64]$tm.numTotalLogicalBytes }
                        elseif ($tm.PSObject.Properties['numBytes']) { $obj.LogicalBytes = [int64]$tm.numBytes }
                        if ($tm.PSObject.Properties['numTotalPhysicalBytes']) { $obj.PhysicalBytes = [int64]$tm.numTotalPhysicalBytes }
                    } catch { $obj.Error='parse' }
                } else { $obj.Error='show' }
                $tableResults += $obj
            }
        }
        # Aggregate results
        foreach ($r in $tableResults) {
            if ($r.Error -and $r.Error -ne '') { $Errors.Value.Add(("Failed table meta {0}:{1}.{2} ({3})" -f $ProjectId,$datasetId,$r.TableId,$r.Error)) }
            $totalLogicalBytes  += $r.LogicalBytes
            $totalPhysicalBytes += $r.PhysicalBytes
            if ($r.PhysicalBytes -gt $largestPhysicalBytes) { $largestPhysicalBytes = $r.PhysicalBytes; $largestLogicalBytesForLargestPhysical = $r.LogicalBytes; $largestTableName = $r.TableId }
            Write-Log ("[DB][BigQuery][Table-Info] ({0}) Dataset={1} Table={2} LogicalBytes={3} PhysicalBytes={4}" -f $ProjectId,$datasetId,$r.TableId,$r.LogicalBytes,$r.PhysicalBytes) -Level DEBUG
        }
        $tblIndex = $tableResults.Count
        if ($tableTotal -gt 0) { Write-Progress -Id 615 -ParentId 614 -Activity ("Tables $datasetId") -Status ("{0}/{1} (100%)" -f $tblIndex,$tableTotal) -PercentComplete 100 -ErrorAction SilentlyContinue }
        # --- End parallel table metadata retrieval ---
        if ($tableTotal -gt 0) { Write-Progress -Id 615 -Activity ("Tables $datasetId") -Completed -ErrorAction SilentlyContinue }
        Write-Log ("[DB][BigQuery] ({0}) DatasetEnd {1} Tables={2} LogicalBytes={3} PhysicalBytes={4} LargestTable={5}" -f $ProjectId,$datasetId,$tableTotal,$totalLogicalBytes,$totalPhysicalBytes,$largestTableName) -Level DEBUG

        $physicalGB = [math]::Round($totalPhysicalBytes / 1GB, 2)
        $logicalGB  = [math]::Round($totalLogicalBytes  / 1GB, 2)
        $largestPhysicalGB = [math]::Round($largestPhysicalBytes / 1GB,2)
        $largestLogicalGB  = [math]::Round($largestLogicalBytesForLargestPhysical / 1GB,2)

        $items += [PSCustomObject]@{
            Type                     = 'BigQueryDataset'
            Project                  = $ProjectId
            Name                     = $datasetId
            Region                   = $region
            Engine                   = 'BigQuery'
            StorageGB                = $physicalGB            # Backward compatibility (now physical size)
            TotalPhysicalSizeGB      = $physicalGB
            TotalLogicalSizeGB       = $logicalGB
            TableCount               = $tableTotal
            LargestTable             = $largestTableName
            LargestTablePhysicalGB   = $largestPhysicalGB
            LargestTableLogicalGB    = $largestLogicalGB
            Error                    = ''
        }
    }
    if ($datasetTotal -gt 0) { Write-Progress -Id 614 -Activity ("BigQuery Datasets $ProjectId") -Completed -ErrorAction SilentlyContinue }

    Write-Log "[DB] ($ProjectId) End BigQuery" -Level DEBUG
    return $items
}

function Invoke-GcpDbProjectInventory {
    param(
        [string]$ProjectId,
        [int]$TimeoutSec,
        [int]$ProjectIndex,
        [int]$TotalProjects,
        [ref]$Errors,
        [ref]$Statuses
    )
    $projStart = Get-Date
    $projStatus = [PSCustomObject]@{ Project=$ProjectId; Success=''; PermissionIssue=''; ApiDisabled=''; Timeout=''; Error=''; DurationSec=0 }
    $enabledSvcNames=@()
    try {
        $svcRaw = & gcloud services list --project $ProjectId --enabled --format="value(config.name)" 2>$null
        if ($LASTEXITCODE -eq 0 -and $svcRaw) { $enabledSvcNames = $svcRaw | Where-Object { $_ } }
    } catch {}
    $requiredMap = @{ CloudSQL='sqladmin.googleapis.com'; Spanner='spanner.googleapis.com'; BigQuery='bigquery.googleapis.com'; AlloyDB='alloydb.googleapis.com' }
    $missingSvcs = @(); foreach ($kv in $requiredMap.GetEnumerator()) { if ($enabledSvcNames -and ($enabledSvcNames -notcontains $kv.Value)) { $missingSvcs += $kv.Value } }
    if ($missingSvcs.Count -gt 0) { $projStatus.ApiDisabled='Y'; if (-not $projStatus.Error){ $projStatus.Error = 'MissingServices=' + ($missingSvcs -join ';') } }
    $displayTotal = if ($TotalProjects -gt 0) { $TotalProjects } else { 1 }
    Write-Progress -Id 610 -Activity 'DB Projects' -Status ("Project $ProjectId ($ProjectIndex/$displayTotal)") -PercentComplete ([int](($ProjectIndex-1)/[double]$displayTotal*100)) -ErrorAction SilentlyContinue
    $allItems=@()
    $projOk=$true
    try {
        $errRef = $Errors
        $statusRef = [ref]$projStatus
        $allItems += Get-GcpCloudSqlInstances -ProjectId $ProjectId -Errors $errRef -Status $statusRef
        $allItems += Get-GcpAlloyDbResources  -ProjectId $ProjectId -Errors $errRef -Status $statusRef
        $allItems += Get-GcpSpannerInstances  -ProjectId $ProjectId
        $allItems += Get-GcpBigQueryDatasets  -ProjectId $ProjectId -Errors $errRef -Status $statusRef
        if ($projStatus.PermissionIssue -or ($projStatus.Error -and $projStatus.Error -match 'permission')) { $projOk=$false }
        if ($errRef.Value.Count -gt 0) {
            # If last error pertains to this project and is critical, mark not ok
            $lastErr = $errRef.Value | Select-Object -Last 1
            if ($lastErr -match $ProjectId -and ($lastErr -match 'permission issue' -or $lastErr -match 'failed')) { $projOk=$false }
        }
    if ($projStatus.ApiDisabled -eq 'Y') { $projOk=$false }
    } catch {
        $errMsg = "General DB inventory failure for ${ProjectId}: $($_.Exception.Message)"
        $Errors.Value.Add($errMsg)
        if (-not $projStatus.Error) { $projStatus.Error = "General failure: $($_.Exception.Message)" }
        $projOk=$false
    }
    if ($projOk) { $script:DbProjectSuccess.Add($ProjectId); $projStatus.Success='Y' } else { $script:DbProjectFailed.Add($ProjectId); if (-not $projStatus.Error -and $Errors.Value.Count -gt 0) { $projStatus.Error = ($Errors.Value | Select-Object -Last 1) } }
    $projStatus.DurationSec = [math]::Round(((Get-Date)-$projStart).TotalSeconds,2)
    if ($projStatus.DurationSec -ge $TimeoutSec -and -not $projStatus.Success) { $projStatus.Timeout='Y'; if (-not $projStatus.Error) { $projStatus.Error='Timeout' } }
    $Statuses.Value.Add($projStatus) | Out-Null
    return @{ Items=$allItems; Status=$projStatus }
}

function Get-GcpDatabaseInventory {
    param(
        [Parameter(Mandatory)][string[]]$ProjectIds,
        [switch]$LightMode,
        [int] $MaxThreads = 8,
        [int] $ProjectTimeoutSec = 900
    )

    if (-not $ProjectIds -or $ProjectIds.Count -eq 0) {
        return [PSCustomObject]@{
            Data   = @();
            Errors = @('No projects supplied for DB inventory')
        }
    }

    $results = [System.Collections.Concurrent.ConcurrentBag[object]]::new()
    $errors  = [System.Collections.Concurrent.ConcurrentBag[string]]::new()
    $script:DbProjectSuccess = New-Object System.Collections.Concurrent.ConcurrentBag[string]
    $script:DbProjectFailed  = New-Object System.Collections.Concurrent.ConcurrentBag[string]
    # Use thread-safe bag for statuses when running in parallel
    if (-not $script:DbProjectStatuses) { $script:DbProjectStatuses = New-Object System.Collections.Concurrent.ConcurrentBag[object] }
    $projQueue = [System.Collections.Concurrent.ConcurrentQueue[string]]::new(); foreach ($p in $ProjectIds) { $projQueue.Enqueue($p) }
    $totalProjects = $ProjectIds.Count
    # Thread-safe progress counter (only add type once)
    if (-not ([System.AppDomain]::CurrentDomain.GetAssemblies() | Where-Object { $_.GetTypes() | Where-Object Name -eq 'DbProg' })) {
        Add-Type -TypeDefinition @"
public class DbProg { private static int counter = 0; public static int Increment(){ return System.Threading.Interlocked.Increment(ref counter);} }
"@
    }
    $scriptBlock = {
        param($queue,$res,$errs,$total,$statuses,$timeoutSec)
        $proj = $null; $index=0
        while ($queue.TryDequeue([ref]$proj)) {
            $index++
            try {
                $localErrs = New-Object System.Collections.Generic.List[string]
                $statusListRef = [ref]$statuses
                $inv = Invoke-GcpDbProjectInventory -ProjectId $proj -TimeoutSec $timeoutSec -ProjectIndex $index -TotalProjects $total -Errors ([ref]$localErrs) -Statuses $statusListRef
                foreach ($i in $inv.Items) { $res.Add($i) }
                foreach ($le in $localErrs) { $errs.Add($le) }
                $processed=[DbProg]::Increment(); $pctParent = if ($total -gt 0) { [int](($processed/[double]$total)*100) } else { 100 }
                Write-Progress -Id 610 -Activity 'DB Projects' -Status ("Project $ProjectId ($ProjectIndex/$displayTotal)") -PercentComplete $pctParent -ErrorAction SilentlyContinue
            } catch {
                $errs.Add("General DB inventory failure for ${proj}: $($_.Exception.Message)")
            }
        }
    }
    # Parallel runspaces disabled permanently for DB path due to inconsistent helper propagation.
    $usePool = $false
    $threads = [math]::Min($MaxThreads,[math]::Max(1,$ProjectIds.Count))
    # Always sequential execution
    & $scriptBlock $projQueue $results $errors $totalProjects $script:DbProjectStatuses $ProjectTimeoutSec
    foreach ($e in $errors) {
        $projName = ($ProjectIds | Where-Object { $e -like "*$_*" } | Select-Object -First 1); if (-not $projName) { $projName='' }
        $results.Add([PSCustomObject]@{Type='Error';Project=$projName;Name='';Region='';Zone='';Engine='';TierOrCapacity='';State='';StorageGB=0;Extra='';Error=($e -replace ',',';')})
    }
    return [PSCustomObject]@{ Data=$results.ToArray(); Errors=$errors.ToArray() }
}


function Get-GcpGkeInventory {
    <#
        .SYNOPSIS
            Modular GKE inventory (refactored from monolithic implementation)
        .OUTPUT
            Hashtable: @{ Clusters=[list]; PVs=[list]; PVCs=[list] }
        .NOTES
            Preserves prior output shape for downstream consumers.
    #>
    param(
        [string[]]$ProjectIds,
        [int]$MaxThreads = 4,          # Currently unused (sequential for safety w/ kubectl contexts)
        [int]$TaskTimeoutSec = 600,
        [switch]$LightMode
    )
    if (-not $ProjectIds -or $ProjectIds.Count -eq 0) { return @{ Clusters=@(); PVs=@(); PVCs=@() } }

    $script:gkeVerbose = $true
    if ($gkeVerbose) { Write-Log -Level INFO -Message ("[GKE-INIT] TargetProjects={0}" -f ($ProjectIds -join ',')) }

    # Detect kubectl (already ensured earlier if -Types GKE used)
    $resolvedKubectl = $null
    if (Get-Variable -Name __KubectlResolvedPath -Scope Script -ErrorAction SilentlyContinue) {
        if ($script:__KubectlResolvedPath -and (Test-Path $script:__KubectlResolvedPath)) { $resolvedKubectl = $script:__KubectlResolvedPath }
    }
    if (-not $resolvedKubectl) {
        $kcCmd = Get-Command kubectl -ErrorAction SilentlyContinue
        if ($kcCmd) { $resolvedKubectl = $kcCmd.Source }
    }
    $kubectlPresent = [bool]$resolvedKubectl
    if (-not $kubectlPresent) { Write-Log -Level WARN -Message "[GKE] kubectl not detected; will rely on disk heuristics if PV/PVC API path unavailable." }
    else { Write-Log -Level INFO -Message ("[GKE-ENV] kubectlPresent=True Path={0}" -f $resolvedKubectl) }

    # ----------------- Helper Functions -----------------
    function Get-GkeClustersForProject {
        param([string]$ProjectId)
        $clusters=@(); $apiDisabled=$false; $permIssue=$false
        try {
            $raw = & gcloud --quiet container clusters list --project $ProjectId --format=json 2>$null
            if ($LASTEXITCODE -ne 0) {
                $msg=($raw|Out-String)
                if ($msg -match '(?i)not enabled|has not been used|api .* enable') { $apiDisabled=$true }
                elseif ($msg -match '(?i)permission|denied|forbidden|403') { $permIssue=$true }
            } elseif (-not [string]::IsNullOrWhiteSpace(($raw|Out-String))) {
                try { $clusters = $raw | ConvertFrom-Json } catch {
                    $first= (($raw|Out-String) -split "`n")[0]
                    Write-Log -Level WARN -Message ("[GKE] ({0}) Cluster JSON parse failure FirstLine='{1}'" -f $ProjectId,($first -replace ',',';'))
                }
            }
        } catch { Write-Log -Level WARN -Message ("[GKE] ({0}) Exception listing clusters: {1}" -f $ProjectId,$_.Exception.Message) }
        return [PSCustomObject]@{Clusters=$clusters; ApiDisabled=$apiDisabled; PermIssue=$permIssue}
    }

    function Get-GkeProjectDisks {
        param([string]$ProjectId,[bool]$Skip)
        if ($Skip) { return @() }
        try {
            $diskRaw = & gcloud --quiet compute disks list --project $ProjectId --format=json 2>$null
            if ($LASTEXITCODE -eq 0 -and $diskRaw) { try { return ($diskRaw | ConvertFrom-Json) } catch {} }
        } catch {}
        return @()
    }

    function Convert-QuantityToGiB { param([string]$Val)
        if (-not $Val) { return 0 }
        if ($Val -match '^(?<num>[0-9]+\.?[0-9]*)(?<unit>(Ki|Mi|Gi|Ti|Pi|Ei|K|M|G|T|P|E))?$') {
            $num=[double]$Matches.num; $u=$Matches.unit
            switch ($u) {
                'Ki' { return $num/1024/1024 }
                'Mi' { return $num/1024 }
                'Gi' { return $num }
                'Ti' { return $num*1024 }
                'Pi' { return $num*1024*1024 }
                'Ei' { return $num*1024*1024*1024 }
                'K'  { return $num/1024/1024 }
                'M'  { return $num/1024 }
                'G'  { return $num }
                'T'  { return $num*1024 }
                'P'  { return $num*1024*1024 }
                'E'  { return $num*1024*1024*1024 }
                default { return $num }
            }
        } elseif ($Val -match '^(?<bytes>[0-9]+)$') { return ([double]$Matches.bytes)/1GB }
        return 0
    }

    function Get-GkeNodePoolResources { param($Cluster)
        $nodeCount=0; $vCpu=0; $memGb=0
        if (-not $Cluster.nodePools) { return @{NodeCount=0;vCpu=0;memGb=0} }
        $idx=0
        foreach ($np in $Cluster.nodePools) {
            $idx++
            $nNodes = 0
            if ($np.initialNodeCount) { $nNodes=[int]$np.initialNodeCount }
            elseif ($np.statusMessage -match 'size=([0-9]+)') { $nNodes=[int]$Matches[1] }
            if ($np.autoscaling -and $np.autoscaling.enabled -and $np.autoscaling.autoprovisioningEnabled -ne $true) {
                if ($np.statusMessage -match 'size=([0-9]+)') { $nNodes=[int]$Matches[1] }
            }
            $nodeCount += $nNodes
            $mt = $null; if ($np.config) { $mt=$np.config.machineType }
            if ($mt) {
                switch -regex ($mt) {
                    'e2-micro' { $vCpu += 2*$nNodes; $memGb += 1*$nNodes }
                    'e2-small' { $vCpu += 2*$nNodes; $memGb += 2*$nNodes }
                    'e2-medium'{ $vCpu += 2*$nNodes; $memGb += 4*$nNodes }
                    'e2-standard-(\d+)' { $c=[int]$Matches[1]; $vCpu += $c*$nNodes; $memGb += ($c*4)*$nNodes }
                    'n1-standard-(\d+)' { $c=[int]$Matches[1]; $vCpu += $c*$nNodes; $memGb += ($c*3.75)*$nNodes }
                    'n2-standard-(\d+)' { $c=[int]$Matches[1]; $vCpu += $c*$nNodes; $memGb += ($c*4)*$nNodes }
                }
            }
        }
        return @{NodeCount=$nodeCount;vCpu=$vCpu;memGb=$memGb}
    }

    function Get-GkePvAndPvcInfo {
        param(
            [string]$ProjectId,
            [string]$ClusterName,
            [string]$Region,
            [string]$Zone,
            [bool]$UseKubectl,
            [object[]]$ProjectDisks,
            [string]$KubectlPath
        )
        $pvDetails = New-Object System.Collections.Generic.List[object]
        $pvcDetails= New-Object System.Collections.Generic.List[object]
        $pvCap=0; $pvCount=0; $source=''
        if (-not $UseKubectl) {
            return @{ PvDetails=$pvDetails; PvcDetails=$pvcDetails; CapacityGiB=0; Count=0; Source=''; }
        }
        $tmpKube = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(),"kubeconfig_${ProjectId}_${ClusterName}_"+[guid]::NewGuid().ToString()+'.tmp')
        try {
            $env:KUBECONFIG=$tmpKube
            Write-Log -Level INFO -Message ("GKE :: Get PVC info started for {0}:{1}" -f $ProjectId,$ClusterName)
            $kubectlCmd = if ($KubectlPath -and (Test-Path $KubectlPath)) { $KubectlPath } else { 'kubectl' }
            if ($Zone) { $connectionOP = & gcloud --quiet container clusters get-credentials $ClusterName --zone $Zone --project $ProjectId 2>$null } else { $connectionOP = & gcloud --quiet container clusters get-credentials $ClusterName --region $Region --project $ProjectId 2>$null }

            Write-Log -Level DEBUG -Message ("[GKE] ({0}) {1} get-credentials output: {2}" -f $ProjectId,$ClusterName,($connectionOP|Out-String).Trim())

            if ($LASTEXITCODE -ne 0) { Write-Log -Level WARN -Message ("[GKE] ({0}) {1} get-credentials failed" -f $ProjectId,$ClusterName); return @{PvDetails=$pvDetails;PvcDetails=$pvcDetails;CapacityGiB=0;Count=0;Source=''} }
            # PV list
            Write-Log -Level INFO -Message ("GKE :: Get PV info started for {0}:{1}" -f $ProjectId,$ClusterName)
            $pvRaw = & $kubectlCmd get pv -o json --request-timeout=25s 2>$null
            Write-Log -Level DEBUG -Message ("[GKE] ({0}) {1} kubectl get pv exitcode={2}" -f $ProjectId,$ClusterName,$LASTEXITCODE)
            Write-Log -Level DEBUG -Message ("[GKE] ({0}) {1} kubectl get pv raw: {2}" -f $ProjectId,$ClusterName,($pvRaw|Out-String).Trim())
            if ($LASTEXITCODE -eq 0 -and $pvRaw) {
                try {
                    $pvObj = $pvRaw | ConvertFrom-Json -ErrorAction Stop
                } catch {
                    # Attempt to sanitize raw output (strip any leading non-JSON noise / trailing text)
                    try {
                        $pvText = ($pvRaw | Out-String).Trim()
                        $startIdx = $pvText.IndexOf('{')
                        if ($startIdx -gt 0) { $pvText = $pvText.Substring($startIdx) }
                        $endIdx = $pvText.LastIndexOf('}')
                        if ($endIdx -gt 0 -and $endIdx -lt ($pvText.Length-1)) { $pvText = $pvText.Substring(0,$endIdx+1) }
                        $pvObj = $pvText | ConvertFrom-Json -ErrorAction Stop
                        Write-Log -Level INFO -Message ("[GKE][PV-JSON-SANITIZE] ({0}) {1} Applied sanitation for kubectl pv output" -f $ProjectId,$ClusterName)
                    } catch {
                        Write-Log -Level WARN -Message ("[GKE] ({0}) {1} PV parse failed after sanitation: {2}" -f $ProjectId,$ClusterName,$_.Exception.Message)
                        $pvObj = $null
                    }
                }
                if ($pvObj -and $pvObj.items) {
                    foreach ($pv in ($pvObj.items | Where-Object { $_ })) {
                        if ($pv.status.phase -and $pv.status.phase -match 'Released|Failed') { continue }
                        $capRaw = if ($pv.spec.capacity.storage) { $pv.spec.capacity.storage } elseif ($pv.status.capacity.storage) { $pv.status.capacity.storage } else { $null }
                        $gi = Convert-QuantityToGiB $capRaw
                        if ($gi -gt 0) {
                            $pvCap += [double]([math]::Round($gi,4)); $pvCount++
                            $backingDisk = ''
                            if ($pv.spec.gcePersistentDisk -and $pv.spec.gcePersistentDisk.pdName) { $backingDisk = $pv.spec.gcePersistentDisk.pdName }
                            $pvDetails.Add([PSCustomObject]@{Project=$ProjectId;ClusterName=$ClusterName;PVName=$pv.metadata.name;Phase=$pv.status.phase;CapacityGiB=[math]::Round($gi,4);StorageClass=$pv.spec.storageClassName;
                                ReclaimPolicy=$pv.spec.persistentVolumeReclaimPolicy;AccessModes=(($pv.spec.accessModes) -join ';');VolumeMode=$pv.spec.volumeMode;ClaimNamespace=$pv.spec.claimRef.namespace;ClaimName=$pv.spec.claimRef.name;
                                Provisioner=$pv.metadata.annotations.'pv.kubernetes.io/provisioned-by';BackingDisk=$backingDisk;CreationTimestamp=([datetime]$pv.metadata.creationTimestamp);AgeDays=[math]::Round((New-TimeSpan -Start ([datetime]$pv.metadata.creationTimestamp) -End (Get-Date)).TotalDays,2) }) | Out-Null
                        }
                    }
                }
            }
            if ($pvCount -gt 0) { $source='pv' }

            # PVC list (always captured for detail)
            $pvcRaw = & $kubectlCmd get pvc -A -o json --request-timeout=30s 2>$null
            if ($LASTEXITCODE -eq 0 -and $pvcRaw) {
                $pvcCapSum = 0; $pvcCount = 0
                try {
                    $pvcObj = $pvcRaw | ConvertFrom-Json
                    foreach ($pvc in ($pvcObj.items | Where-Object { $_ })) {
                        $reqCap = if ($pvc.spec.resources.requests.storage) { $pvc.spec.resources.requests.storage } else { $null }
                        $actCap = if ($pvc.status.capacity.storage) { $pvc.status.capacity.storage } else { $reqCap }
                        $reqGi = Convert-QuantityToGiB $reqCap
                        $actGi = Convert-QuantityToGiB $actCap
                        if ($actGi -gt 0) { $pvcCapSum += [double]([math]::Round($actGi,4)); $pvcCount++ }
                        $pvcDetails.Add([PSCustomObject]@{Project=$ProjectId;ClusterName=$ClusterName;Namespace=$pvc.metadata.namespace;PVCName=$pvc.metadata.name;Status=$pvc.status.phase;RequestedGiB=[math]::Round($reqGi,4);CapacityGiB=[math]::Round($actGi,4);StorageClass=$pvc.spec.storageClassName;AccessModes=(($pvc.spec.accessModes)-join ';');VolumeMode=$pvc.spec.volumeMode;PVName=$pvc.spec.volumeName;CreationTimestamp=([datetime]$pvc.metadata.creationTimestamp);AgeDays=[math]::Round((New-TimeSpan -Start ([datetime]$pvc.metadata.creationTimestamp) -End (Get-Date)).TotalDays,2)}) | Out-Null
                    }
                    # If no PV objects were found, fall back to aggregated PVC capacities/count
                    if ($pvCount -eq 0 -and $pvcCount -gt 0) { $pvCap = [math]::Round($pvcCapSum,2); $pvCount = $pvcCount; $source = 'pvc' }
                    else {
                        # Reconcile if PV object count appears incomplete vs distinct PVC PVNames
                        $distinctPvNames = ($pvcDetails | Where-Object { $_.PVName } | Select-Object -ExpandProperty PVName -Unique)
                        if ($pvCount -gt 0 -and $distinctPvNames.Count -gt $pvCount) {
                            $uniqueCap = 0
                            $seen = @{}
                            foreach ($pvcEntry in ($pvcDetails | Where-Object { $_.PVName })) {
                                if (-not $seen.ContainsKey($pvcEntry.PVName)) {
                                    if ($pvcEntry.CapacityGiB -gt 0) { $uniqueCap += [double]$pvcEntry.CapacityGiB }
                                    $seen[$pvcEntry.PVName] = $true
                                }
                            }
                            if ($uniqueCap -gt 0 -and $uniqueCap -gt $pvCap) {
                                Write-Log -Level INFO -Message ("[GKE][PV-RECONCILE] ({0}) {1} Adjusting PVCount {2}->{3} PVCapGB {4}->{5}" -f $ProjectId,$ClusterName,$pvCount,$distinctPvNames.Count,$pvCap,[math]::Round($uniqueCap,2))
                                $pvCap = [math]::Round($uniqueCap,2)
                                $pvCount = $distinctPvNames.Count
                                $source = 'pvc-reconcile'
                            }
                        }
                    }
                } catch { Write-Log -Level WARN -Message ("[GKE] ({0}) {1} PVC parse failed: {2}" -f $ProjectId,$ClusterName,$_.Exception.Message) }
            }
        } finally { try { if (Test-Path $tmpKube) { Remove-Item $tmpKube -Force -ErrorAction SilentlyContinue } } catch {} }

        # Heuristic if still zero
        if ($pvCount -eq 0 -and $ProjectDisks -and $ProjectDisks.Count -gt 0) {
            $matched = @()
            foreach ($d in $ProjectDisks) {
                if ($d.name -match "gke-${ClusterName}.*-node-") { continue }
                $hasLabel=$false
                if ($d.labels) {
                    foreach ($ln in $d.labels.PSObject.Properties.Name) { if ($ln -match 'kubernetes-io-created-for-pv-name|kubernetes-io-created-for-pvc-name') { $hasLabel=$true; break } }
                }
                if ($hasLabel -or $d.name -like "gke-${ClusterName}-*") { $matched += $d }
            }
            if ($matched.Count -gt 0) {
                foreach ($m in $matched) { if ($m.sizeGb) { $pvCap += [double]$m.sizeGb } }
                $pvCount = $matched.Count; $source='disk'
                Write-Log -Level INFO -Message ("[GKE][DISK-HEURISTIC] ({0}) Cluster={1} Disks={2} CapGB={3}" -f $ProjectId,$ClusterName,$matched.Count,$pvCap)
            }
        }
        return @{ PvDetails=$pvDetails; PvcDetails=$pvcDetails; CapacityGiB=[math]::Round($pvCap,2); Count=$pvCount; Source=$source }
    }

    $allClusters   = New-Object System.Collections.Generic.List[object]
    $allPvDetails  = New-Object System.Collections.Generic.List[object]
    $allPvcDetails = New-Object System.Collections.Generic.List[object]

    foreach ($proj in $ProjectIds) {
        $projStart=[DateTime]::UtcNow
        Write-Log -Level INFO -Message ("[GKE-PROJECT-START] {0}" -f $proj)
        $clResult = Get-GkeClustersForProject -ProjectId $proj
        $projDisks = Get-GkeProjectDisks -ProjectId $proj -Skip ($clResult.ApiDisabled -or $clResult.PermIssue)
        if ($gkeVerbose) { Write-Log -Level DEBUG -Message ("[GKE-PROJECT] {0} DisksEnumerated={1}" -f $proj,$projDisks.Count) }
        $clusters = $clResult.Clusters
        if (-not $clusters -or $clusters.Count -eq 0) {
            Write-Log -Level INFO -Message ("[GKE-PROJECT] {0} ClusterCount=0 (ApiDisabled={1} PermIssue={2})" -f $proj,$clResult.ApiDisabled,$clResult.PermIssue)
        } else {
            Write-Log -Level INFO -Message ("[GKE-PROJECT] {0} ClusterCount={1}" -f $proj,$clusters.Count)
        }
        foreach ($c in ($clusters | Where-Object { $_ })) {
            $name = $c.name
            $location = if ($c.location) { $c.location } elseif ($c.zone) { ($c.zone -split '/')[-1] } elseif ($c.region) { $c.region } else { '' }
            $region = if ($location -match '^[a-z0-9-]+-[a-z]$') { $location -replace '-[a-z]$','' } else { $location }
            $zone = if ($location -match '^[a-z0-9-]+-[a-z]$') { $location } else { '' }
            $version = $c.currentMasterVersion; if (-not $version) { $version=$c.currentNodeVersion }
            $npRes = Get-GkeNodePoolResources -Cluster $c
            $pvInfo = Get-GkePvAndPvcInfo -ProjectId $proj -ClusterName $name -Region $region -Zone $zone -UseKubectl:$kubectlPresent -ProjectDisks $projDisks -KubectlPath $resolvedKubectl
            if ($pvInfo.PvDetails.Count -gt 0) { foreach ($pv in $pvInfo.PvDetails) { $allPvDetails.Add($pv)|Out-Null } }
            if ($pvInfo.PvcDetails.Count -gt 0) { foreach ($pvc in $pvInfo.PvcDetails) { $allPvcDetails.Add($pvc)|Out-Null } }
            $network = if ($c.network) { ($c.network -split '/')[-1] } else { '' }
            $subnet  = if ($c.subnetwork) { ($c.subnetwork -split '/')[-1] } else { '' }
            $releaseChannel = if ($c.releaseChannel.channel) { $c.releaseChannel.channel } else { '' }
            $obj = [PSCustomObject]@{
                Project       = $proj
                ClusterName   = $name
                Region        = $region
                Zone          = $zone
                Location      = $location
                KubernetesVer = $version
                NodeCount     = $npRes.NodeCount
                EstimatedvCPU = [int]$npRes.vCpu
                EstimatedMemGB= [int][math]::Round($npRes.memGb,0)
                PersistentVolumeCapacityGB = [double]$pvInfo.CapacityGiB
                PersistentVolumeCount      = $pvInfo.Count
                PersistentVolumeSource     = $pvInfo.Source
                ReleaseChannel= $releaseChannel
                Network       = $network
                Subnetwork    = $subnet
                Endpoint      = if ($c.endpoint){$c.endpoint}else{''}
                ApiDisabled   = $clResult.ApiDisabled
                PermissionIssue = $clResult.PermIssue
            }
            # Attach detail arrays (for downstream harvesting to remain compatible)
            $null = $obj | Add-Member -NotePropertyName _PvDetails  -NotePropertyValue ($pvInfo.PvDetails.ToArray())  -Force
            $null = $obj | Add-Member -NotePropertyName _PvcDetails -NotePropertyValue ($pvInfo.PvcDetails.ToArray()) -Force
            $allClusters.Add($obj)|Out-Null
            Write-Log -Level INFO -Message ("[GKE] Project={0} Cluster={1} Nodes={2} PVCapGB={3} PVCount={4} Source={5}" -f $proj,$name,$npRes.NodeCount,$pvInfo.CapacityGiB,$pvInfo.Count,$pvInfo.Source)
        }
        $dur=[math]::Round(([DateTime]::UtcNow-$projStart).TotalSeconds,2)
        Write-Log -Level INFO -Message ("[GKE-PROJECT-END] {0} Clusters={1} DurationSec={2}" -f $proj,($allClusters | Where-Object Project -eq $proj).Count,$dur)
    }

    return @{ Clusters=$allClusters; PVs=$allPvDetails; PVCs=$allPvcDetails }
}


# -------------------------
# Execution Flow
# -------------------------
$allProjects = Get-GcpProjects
# Additional normalization & case-insensitive resolution for user-specified projects
if ($Projects) {
    # Trim & dedupe input list
    $Projects = $Projects | ForEach-Object { $_.Trim() } | Where-Object { $_ } | Select-Object -Unique
    # Build case-insensitive lookup of all accessible projects
    $allLookup = @{}
    foreach ($p in $allProjects) { $allLookup[$p.ToLower()] = $p }
    $resolved=@(); $invalid=@()
    foreach ($p in $Projects) {
        $k = $p.ToLower()
        if ($allLookup.ContainsKey($k)) { $resolved += $allLookup[$k] } else { $invalid += $p }
    }
    if ($invalid.Count -gt 0) {
        Write-Warning ("Ignoring invalid project id(s): {0}" -f ($invalid -join ', '))
    }
    $targetProjects = $resolved | Select-Object -Unique
    if (-not $targetProjects -or $targetProjects.Count -eq 0) {
        Write-Error "No valid projects found from provided -Projects list."
        Stop-Transcript | Out-Null
        exit 1
    }
} else { $targetProjects = $allProjects }
Write-Host "Targeting $($targetProjects.Count) projects." -ForegroundColor Green

$invResults = @{}
if ($Selected.VM)      { $invResults = Get-GcpVMInventory -ProjectIds $targetProjects }
if ($Selected.STORAGE) { $invResults.StorageBuckets = Get-GcpStorageInventory -ProjectIds $targetProjects }
if ($Selected.FILESHARE) {
    # Compute filestore thread cap (10 or project count, whichever is smaller)
    $fsThreads = [Math]::Min(10,[Math]::Max(1,$targetProjects.Count))
    $invResults.FileShares = Get-GcpFileShareInventory -ProjectIds $targetProjects -LightMode:$LightMode -MaxThreads $fsThreads
}
if ($Selected.DB) {
    $dbThreads = [Math]::Min(8,[Math]::Max(1,$targetProjects.Count))
    $dbResult = Get-GcpDatabaseInventory -ProjectIds $targetProjects -LightMode:$LightMode -MaxThreads $dbThreads -ProjectTimeoutSec $DbProjectTimeoutSec
    $invResults.Databases = $dbResult.Data
    $invResults.DatabaseErrors = $dbResult.Errors
}
if ($Selected.GKE) {
    # Ensure kubectl prior to inventory (best-effort)
    $kubectlPath = Ensure-Kubectl
    if (-not $kubectlPath -and $ForceKubectl) {
        Write-Host '[GKE] kubectl missing; retrying provisioning once due to -ForceKubectl.' -ForegroundColor Yellow
        Start-Sleep -Seconds 2
        $kubectlPath = Ensure-Kubectl -Force
    }
    if ($ForceKubectl -and -not $kubectlPath) {
        Write-Error 'kubectl is required (-ForceKubectl specified) but could not be provisioned. Aborting GKE inventory.'
        Stop-Transcript | Out-Null
        exit 2
    }
    $gkeThreads = [Math]::Min(8,[Math]::Max(1,$targetProjects.Count))
    $gkeInv = Get-GcpGkeInventory -ProjectIds $targetProjects -MaxThreads $gkeThreads -LightMode:$LightMode
    $invResults.GKEClusters = $gkeInv.Clusters
    $invResults.GKEPVs      = $gkeInv.PVs
    $invResults.GKEPVCs     = $gkeInv.PVCs
}

# Global Filestore reconstruction (if we somehow ended up with zero successful shares but logs contain entries)
if ($Selected.FILESHARE) {
    $fsShares = $invResults.FileShares
    $fsGood = @()
    if ($fsShares) { $fsGood = $fsShares | Where-Object { -not ($_.Error -and $_.Error.Trim()) -and ($_.ShareName -ne '(error)') -and ($_.State -ne 'ERROR') } }
    if ((-not $fsGood) -or $fsGood.Count -eq 0) {
        if ($script:FileShareLogLines -and $script:FileShareLogLines.Count -gt 0) {
            $recoveredGlobal=@()
            foreach ($line in $script:FileShareLogLines) {
                if ($line -match '^\[FS\] Project=([^ ]+) .* Instance=([^ ]+) Share=([^ ]+) Tier=([^ ]+) Region=([^ ]+) Zone=([^ ]+) (?:CapacityGB|provisionedGib)=([0-9]+)') {
                    $gib=[int64]$Matches[7]
                    $gb=[math]::Round(($gib * 1024 * 1024 * 1024)/1e9,3)
                    $tib=[math]::Round($gib/1024,4)
                    $tb=[math]::Round($gb/1000,4)
                    $recoveredGlobal += [PSCustomObject]@{
                        Project=$Matches[1]; InstanceName=$Matches[2]; ShareName=$Matches[3]; Tier=$Matches[4]; Region=$Matches[5]; Zone=$Matches[6]; provisionedGb=$gb; provisionedGib=$gib; provisionedTb=$tb; provisionedTib=$tib; CapacityGB=$gib; Networks=''; IPAddresses=''; State='READY'; CreateTime=''; Labels=''; Protocol='NFS'; Error=''
                    }
                }
            }
            if ($recoveredGlobal.Count -gt 0) {
                Write-Host ("[FS-Global-Recover] Reconstructed {0} shares from log lines" -f $recoveredGlobal.Count) -ForegroundColor Yellow
                $invResults.FileShares = $recoveredGlobal
            } else {
                Write-Host "[FS-Global-Recover] No reconstructable [FS] log lines found." -ForegroundColor DarkYellow
            }
        } else {
            Write-Host "[FS-Global-Recover] No filestore log lines captured." -ForegroundColor DarkYellow
        }
    }
}

# ================= DB CSV EXPORTS =================
# Helper must be defined before usage
if (-not (Get-Command Add-BlankLines -ErrorAction SilentlyContinue)) {
    function Add-BlankLines {
        param(
            [Parameter(Mandatory=$true)][string]$Path,
            [int]$Count = 4
        )
        if (-not $Count -or $Count -le 0) { return }
        # Create file if it does not yet exist so Add-Content won't fail
        if (-not (Test-Path -LiteralPath $Path)) { New-Item -ItemType File -Path $Path -Force | Out-Null }
        for ($i=0; $i -lt $Count; $i++) { Add-Content -Path $Path -Value '' }
    }
}
if (-not (Get-Command Add-DbWorkloadSummary -ErrorAction SilentlyContinue)) {
function Add-DbWorkloadSummary {
    param(
        [string]$Path,
        [object[]]$Data,
        [string]$WorkloadName,
        [string]$CountLabel = 'Objects',
        [string]$SizeProperty = 'StorageGB',
        [string]$AltSizeLabel = ''
    )
    if (-not $Data -or $Data.Count -eq 0) { return }
    Add-BlankLines -Path $Path -Count 4
    Add-Content -Path $Path -Value ("### {0} Summary ###" -f $WorkloadName)
    $total = $Data.Count
    $hasSize = $false
    $sumSize = 0
    if ($SizeProperty -and ($Data | Select-Object -First 1).PSObject.Properties[$SizeProperty]) {
        $hasSize = $true
        $sumSize = ($Data | Measure-Object -Property $SizeProperty -Sum).Sum
    }
    if ($hasSize) {
        Add-Content -Path $Path -Value ("Total {0}, {1}, TotalSizeGB, {2}" -f $CountLabel,$total,$sumSize)
    } else {
        Add-Content -Path $Path -Value ("Total {0}, {1}" -f $CountLabel,$total)
    }
    Add-BlankLines -Path $Path -Count 1
    Add-Content -Path $Path -Value 'Per Project:'
    $projGroups = $Data | Group-Object Project | Sort-Object Name
    foreach ($pg in $projGroups) {
        if ($hasSize) {
            $pSize = ($pg.Group | Measure-Object -Property $SizeProperty -Sum).Sum
            Add-Content -Path $Path -Value ("Project, {0}, {1}, {2}, SizeGB, {3}" -f $pg.Name,$CountLabel,$pg.Count,$pSize)
        } else {
            Add-Content -Path $Path -Value ("Project, {0}, {1}, {2}" -f $pg.Name,$CountLabel,$pg.Count)
        }
    }
    if (($Data | Select-Object -First 1).PSObject.Properties['Region']) {
        Add-BlankLines -Path $Path -Count 1
        Add-Content -Path $Path -Value 'Per Project Region:'
        foreach ($pg in $projGroups) {
            $regionGroups = $pg.Group | Group-Object Region | Sort-Object Name
            foreach ($rg in $regionGroups) {
                if ($hasSize) {
                    $rSize = ($rg.Group | Measure-Object -Property $SizeProperty -Sum).Sum
                    Add-Content -Path $Path -Value ("ProjectRegion, {0}, {1}, {2}, {3}, SizeGB, {4}" -f $pg.Name,$rg.Name,$CountLabel,$rg.Count,$rSize)
                } else {
                    Add-Content -Path $Path -Value ("ProjectRegion, {0}, {1}, {2}, {3}" -f $pg.Name,$rg.Name,$CountLabel,$rg.Count)
                }
            }
        }
    }
}
}
if ($Selected.DB -and $invResults.Databases -and (@($invResults.Databases)).Count) {
    $dbAll      = @($invResults.Databases)
    $dbSuccess  = @($dbAll | Where-Object { -not ($_.Error -and $_.Error.Trim()) -and $_.Type -ne 'Error' })
    $dbErrors   = @($dbAll | Where-Object { ($_.Error -and $_.Error.Trim()) -or $_.Type -eq 'Error' })

    # Cloud SQL
    $cloudSql = $dbSuccess | Where-Object Type -eq 'CloudSQL'
    if ($cloudSql.Count -gt 0) {
        $cloudSqlCsv = Join-Path $outDir ("gcp_cloudsql_instances_" + $dateStr + ".csv")
        'Type,Project,Name,Region,Zone,Engine,TierOrCapacity,State,StorageGB,PrivateIPs,PublicIPs,OutgoingIPs,Error' | Out-File -FilePath $cloudSqlCsv -Encoding utf8
        foreach ($r in ($cloudSql | Sort-Object Project,Name)) {
            ("{0},{1},{2},{3},{4},{5},{6},{7},{8},{9},{10},{11},{12}" -f $r.Type,$r.Project,$r.Name,$r.Region,$r.Zone,$r.Engine,$r.TierOrCapacity,$r.State,$r.StorageGB,$r.PrivateIPs,$r.PublicIPs,$r.OutgoingIPs,($r.Error -replace ',',';')) | Out-File -FilePath $cloudSqlCsv -Append -Encoding utf8
        }
        Add-Content -Path $cloudSqlCsv -Value ''
        Add-Content -Path $cloudSqlCsv -Value '### Cloud SQL Summary ###'
        $totalInst = $cloudSql.Count
        $totalGb = ($cloudSql | Measure-Object StorageGB -Sum).Sum
        Add-Content -Path $cloudSqlCsv -Value ("Total Instances, {0}, TotalSizeGB, {1}" -f $totalInst,$totalGb)
        Add-Content -Path $cloudSqlCsv -Value 'Per Project:'
        foreach ($g in ($cloudSql | Group-Object Project | Sort-Object Name)) {
            $gGb = ($g.Group | Measure-Object StorageGB -Sum).Sum
            Add-Content -Path $cloudSqlCsv -Value ("Project, {0}, Instances, {1}, SizeGB, {2}" -f $g.Name,$g.Count,$gGb)
        }
    # Success / Failed sections
    Add-Content -Path $cloudSqlCsv -Value ''
    Add-Content -Path $cloudSqlCsv -Value 'SuccessfulProjects:'
    Add-Content -Path $cloudSqlCsv -Value 'Project,Instances,TotalSizeGB'
    foreach ($grp in ($cloudSql | Group-Object Project | Sort-Object Name)) { $sz=($grp.Group|Measure-Object StorageGB -Sum).Sum; Add-Content -Path $cloudSqlCsv -Value ("{0},{1},{2}" -f $grp.Name,$grp.Count,$sz) }
    $failedCsql=@(); foreach ($p in $targetProjects) { if (-not ($cloudSql | Where-Object Project -eq $p)) { $st=$null; if ($script:DbProjectStatuses){ $st=$script:DbProjectStatuses|Where-Object Project -eq $p|Select-Object -First 1 }; if ($st -and ($st.PermissionIssue -eq 'Y' -or $st.ApiDisabled -eq 'Y' -or $st.Timeout -eq 'Y' -or ($st.Error -and ($st.Error -match 'CloudSQL' -or $st.Error -match 'sql')))) { $failedCsql+=$st } } }
        if ($failedCsql.Count -gt 0) {
            Add-Content -Path $cloudSqlCsv -Value ''
            Add-Content -Path $cloudSqlCsv -Value 'FailedProjects:'
            Add-Content -Path $cloudSqlCsv -Value 'Project,Error'
            foreach ($f in ($failedCsql|Sort-Object Project)) {
                $errMsg = if ($f.ApiDisabled -eq 'Y') { 'CloudSQL API Disabled' } elseif ($f.PermissionIssue -eq 'Y') { 'CloudSQL PermissionDenied' } elseif ($f.Timeout -eq 'Y') { 'Timeout' } elseif ($f.Error -and $f.Error.Trim()) { $f.Error } else { 'NoData / UnknownFailure' }
                $errMsg = ($errMsg -replace ',',';')
                Add-Content -Path $cloudSqlCsv -Value ("{0},{1}" -f $f.Project,$errMsg)
            }
        }
        Write-Host "Cloud SQL CSV written: $(Split-Path $cloudSqlCsv -Leaf)" -ForegroundColor Cyan
    Add-DbWorkloadSummary -Path $cloudSqlCsv -Data $cloudSql -WorkloadName 'Cloud SQL Instances' -CountLabel 'Instances' -SizeProperty 'StorageGB'
    }

    # Spanner
    # Normalize to array so a single instance does not collapse into a lone PSCustomObject (which would make .Count $null)
    $spanner = @($dbSuccess | Where-Object Type -eq 'SpannerInstance')
    if ($spanner.Count -gt 0) {
        $spannerCsv = Join-Path $outDir ("gcp_spanner_instances_" + $dateStr + ".csv")
        'Type,Project,Name,Region,Zone,Engine,TierOrCapacity,State,StorageGB,Error' | Out-File -FilePath $spannerCsv -Encoding utf8
        foreach ($r in ($spanner | Sort-Object Project,Name)) {
            ("{0},{1},{2},{3},{4},{5},{6},{7},{8},{9}" -f $r.Type,$r.Project,$r.Name,$r.Region,$r.Zone,$r.Engine,$r.TierOrCapacity,$r.State,$r.StorageGB,($r.Error -replace ',',';')) | Out-File -FilePath $spannerCsv -Append -Encoding utf8
        }
    Add-Content -Path $spannerCsv -Value ''
    Add-Content -Path $spannerCsv -Value 'SuccessfulProjects:'
    Add-Content -Path $spannerCsv -Value 'Project,Instances'
    foreach ($grp in ($spanner | Group-Object Project | Sort-Object Name)) { Add-Content -Path $spannerCsv -Value ("{0},{1}" -f $grp.Name,$grp.Count) }
    $failedSp=@(); foreach ($p in $targetProjects) { if (-not ($spanner | Where-Object Project -eq $p)) { $st=$null; if ($script:DbProjectStatuses){ $st=$script:DbProjectStatuses|Where-Object Project -eq $p|Select-Object -First 1 }; if ($st -and ($st.PermissionIssue -eq 'Y' -or $st.ApiDisabled -eq 'Y' -or $st.Timeout -eq 'Y' -or ($st.Error -and $st.Error -match 'Spanner'))) { $failedSp+=$st } } }
    if ($failedSp.Count -gt 0) { Add-Content -Path $spannerCsv -Value ''; Add-Content -Path $spannerCsv -Value 'FailedProjects:'; Add-Content -Path $spannerCsv -Value 'Project,Reason'; foreach ($f in ($failedSp|Sort-Object Project)) { $reason = if ($f.ApiDisabled -eq 'Y') { 'ApiDisabled' } elseif ($f.PermissionIssue -eq 'Y') { 'PermissionDenied' } elseif ($f.Timeout -eq 'Y') { 'Timeout' } elseif ($f.Error -match 'permission') { 'PermissionDenied' } elseif ($f.Error) { 'Error' } else { 'Unknown' }; Add-Content -Path $spannerCsv -Value ("{0},{1}" -f $f.Project,$reason) } }
        Write-Host "Spanner CSV written: $(Split-Path $spannerCsv -Leaf)" -ForegroundColor Cyan
    Add-DbWorkloadSummary -Path $spannerCsv -Data $spanner -WorkloadName 'Spanner Instances' -CountLabel 'Instances' -SizeProperty 'StorageGB'
    }
    else {
        Write-Host '[Spanner] No instances discovered; CSV not generated (expected if API disabled or permission denied).' -ForegroundColor Yellow
    }

    # BigQuery
    # Normalize BigQuery datasets to an array so single-item results don't drop the Count property
    $bigQuery = @($dbSuccess | Where-Object Type -eq 'BigQueryDataset')
    # Always create the dataset CSV if we even attempted DB work (so downstream processes can rely on file existence)
    $bqCsv = Join-Path $outDir ("gcp_bigquery_datasets_" + $dateStr + ".csv")
    'Type,Project,Dataset_Name,Region,Engine,TotalPhysicalSizeGB,TotalLogicalSizeGB,TableCount,LargestTable,LargestTablePhysicalGB,LargestTableLogicalGB,Error' | Out-File -FilePath $bqCsv -Encoding utf8
    if ($bigQuery.Count -gt 0) {
        foreach ($r in ($bigQuery | Sort-Object Project,Name)) {
            $largestTableNameSafe = ($r.LargestTable -replace ',',';')
            ("{0},{1},{2},{3},{4},{5},{6},{7},{8},{9},{10},{11}" -f $r.Type,$r.Project,$r.Name,$r.Region,$r.Engine,$r.TotalPhysicalSizeGB,$r.TotalLogicalSizeGB,$r.TableCount,$largestTableNameSafe,$r.LargestTablePhysicalGB,$r.LargestTableLogicalGB,($r.Error -replace ',',';')) | Out-File -FilePath $bqCsv -Append -Encoding utf8
        }
        Add-Content -Path $bqCsv -Value ''
        Add-Content -Path $bqCsv -Value 'SuccessfulProjects:'
        Add-Content -Path $bqCsv -Value 'Project,Datasets'
        foreach ($grp in ($bigQuery | Group-Object Project | Sort-Object Name)) { Add-Content -Path $bqCsv -Value ("{0},{1}" -f $grp.Name,$grp.Count) }
    } else {
        # leave only header row; indicates 0 datasets found
        Add-Content -Path $bqCsv -Value ''
        Add-Content -Path $bqCsv -Value 'SuccessfulProjects:'
        Add-Content -Path $bqCsv -Value 'Project,Datasets'
    }
    $failedBq=@(); foreach ($p in $targetProjects) { if (-not ($bigQuery | Where-Object Project -eq $p)) { $st=$null; if ($script:DbProjectStatuses){ $st=$script:DbProjectStatuses|Where-Object Project -eq $p|Select-Object -First 1 }; if ($st -and ($st.PermissionIssue -eq 'Y' -or $st.ApiDisabled -eq 'Y' -or $st.Timeout -eq 'Y' -or ($st.Error -and ($st.Error -match 'BigQuery' -or $st.Error -match 'bq-cli')))) { $failedBq+=$st } } }
    if ($failedBq.Count -gt 0) { Add-Content -Path $bqCsv -Value ''; Add-Content -Path $bqCsv -Value 'FailedProjects:'; Add-Content -Path $bqCsv -Value 'Project,Reason'; foreach ($f in ($failedBq|Sort-Object Project)) { $reason = if ($f.ApiDisabled -eq 'Y') { 'ApiDisabled' } elseif ($f.PermissionIssue -eq 'Y') { 'PermissionDenied' } elseif ($f.Timeout -eq 'Y') { 'Timeout' } elseif ($f.Error -match 'permission') { 'PermissionDenied' } elseif ($f.Error) { 'Error' } else { 'Unknown' }; Add-Content -Path $bqCsv -Value ("{0},{1}" -f $f.Project,$reason) } }
    Write-Host "BigQuery CSV written: $(Split-Path $bqCsv -Leaf)" -ForegroundColor Cyan
    Add-DbWorkloadSummary -Path $bqCsv -Data $bigQuery -WorkloadName 'BigQuery Datasets' -CountLabel 'Datasets' -SizeProperty 'TotalPhysicalSizeGB'

    # Optional: detailed tables CSV when any BigQuery dataset has tables (aggregate table info already summarized per dataset)
    $anyTables = $false; if ($bigQuery.Count -gt 0) { $anyTables = ($bigQuery | Where-Object { $_.TableCount -gt 0 }).Count -gt 0 }
    if ($anyTables) {
        $bqTablesCsv = Join-Path $outDir ("gcp_bigquery_tables_" + $dateStr + ".csv")
        'Project,Dataset,TableName,LogicalBytes,PhysicalBytes,LogicalGB,PhysicalGB' | Out-File -FilePath $bqTablesCsv -Encoding utf8
        # Reconstruct table details from logs (DEBUG lines) if not explicitly captured elsewhere
        foreach ($line in $script:ChildLogQueue) {
            if ($line -match '\[DB\]\[BigQuery\]\[Table-Info\] \(([^)]+)\) Dataset=([^ ]+) Table=([^ ]+) LogicalBytes=([0-9]+) PhysicalBytes=([0-9]+)') {
                $proj=$Matches[1]; $ds=$Matches[2]; $tbl=$Matches[3]; $lbytes=[int64]$Matches[4]; $pbytes=[int64]$Matches[5]
                $lgb=[math]::Round($lbytes/1GB,2); $pgb=[math]::Round($pbytes/1GB,2)
                Add-Content -Path $bqTablesCsv -Value ("{0},{1},{2},{3},{4},{5},{6}" -f $proj,$ds,$tbl,$lbytes,$pbytes,$lgb,$pgb)
            }
        }
        if ((Get-Content $bqTablesCsv | Measure-Object -Line).Lines -le 1) { # no rows added
            Add-Content -Path $bqTablesCsv -Value '# No table detail lines found in log (ensure DEBUG logging enabled)'
        } else {
            Write-Host "BigQuery Tables CSV written: $(Split-Path $bqTablesCsv -Leaf)" -ForegroundColor Cyan
        }
    }

    # AlloyDB Clusters
    $alloyClusters = $dbSuccess | Where-Object Type -eq 'alloydbcluster'
    if ($alloyClusters.Count -gt 0) {
        $alloyCsv = Join-Path $outDir ("gcp_alloydb_clusters_" + $dateStr + ".csv")
        'Type,Project,Name,Region,Zone,Engine,TierOrCapacity,State,StorageBytes,StorageMB,StorageGB,MonitoringStatus,Extra' | Out-File -FilePath $alloyCsv -Encoding utf8
        foreach ($cl in ($alloyClusters | Sort-Object Project,Name)) {
            $bytesVal = if ($cl.PSObject.Properties['StorageBytes']) { [int64]$cl.StorageBytes } else { 0 }
            $gbVal = if ($bytesVal -gt 0) { [int][math]::Round($bytesVal/1GB,0) } elseif ($cl.StorageGB -gt 0) { [int]$cl.StorageGB } else { 0 }
            $monStat = if ($cl.PSObject.Properties['MonitoringStatus']) { $cl.MonitoringStatus } else { '' }
            $mbVal = if ($cl.PSObject.Properties['StorageMB']) { [int][math]::Round($cl.StorageMB,0) } elseif ($bytesVal -gt 0) { [int][math]::Round($bytesVal/1MB,0) } else { 0 }
            ("{0},{1},{2},{3},{4},{5},{6},{7},{8},{9},{10},{11},{12}" -f 'alloydbcluster',$cl.Project,$cl.Name,$cl.Region,$cl.Zone,$cl.Engine,$cl.TierOrCapacity,$cl.State,$bytesVal,$mbVal,$gbVal,$monStat,($cl.Extra -replace ',',';')) | Out-File -FilePath $alloyCsv -Append -Encoding utf8
        }
    Add-Content -Path $alloyCsv -Value ''
    Add-Content -Path $alloyCsv -Value 'SuccessfulProjects:'
    Add-Content -Path $alloyCsv -Value 'Project,PrimaryClusters,TotalStorageGB'
    foreach ($grp in ($alloyClusters | Group-Object Project | Sort-Object Name)) { $sumGb = ($grp.Group | Measure-Object StorageGB -Sum).Sum; Add-Content -Path $alloyCsv -Value ("{0},{1},{2}" -f $grp.Name,$grp.Count,$sumGb) }
    $failedAlloy=@(); foreach ($p in $targetProjects) { if (-not ($alloyClusters | Where-Object Project -eq $p)) { $st=$null; if ($script:DbProjectStatuses){ $st=$script:DbProjectStatuses|Where-Object Project -eq $p|Select-Object -First 1 }; if ($st -and ($st.PermissionIssue -eq 'Y' -or $st.ApiDisabled -eq 'Y' -or $st.Timeout -eq 'Y' -or ($st.Error -and $st.Error -match 'AlloyDB'))) { $failedAlloy+=$st } } }
        if ($failedAlloy.Count -gt 0) {
            Add-Content -Path $alloyCsv -Value ''
            Add-Content -Path $alloyCsv -Value 'FailedProjects:'
            Add-Content -Path $alloyCsv -Value 'Project,Error'
            foreach ($f in ($failedAlloy|Sort-Object Project)) {
                $errMsg = if ($f.ApiDisabled -eq 'Y') { 'AlloyDB API Disabled' } elseif ($f.PermissionIssue -eq 'Y') { 'AlloyDB PermissionDenied' } elseif ($f.Timeout -eq 'Y') { 'Timeout' } elseif ($f.Error -and $f.Error.Trim()) { $f.Error } else { 'NoData / UnknownFailure' }
                $errMsg = ($errMsg -replace ',',';')
                Add-Content -Path $alloyCsv -Value ("{0},{1}" -f $f.Project,$errMsg)
            }
        }
        Write-Host "AlloyDB clusters CSV written: $(Split-Path $alloyCsv -Leaf)" -ForegroundColor Cyan
    Add-DbWorkloadSummary -Path $alloyCsv -Data $alloyClusters -WorkloadName 'AlloyDB Primary Clusters' -CountLabel 'Clusters' -SizeProperty 'StorageGB'
    }
}

# -------------------------
# Output CSVs
# -------------------------
Write-Progress -Id 5 -Activity "Generating Output Files" -Status "Exporting CSV files..." -PercentComplete 0

function Write-PlainCsv {
    param(
        [Parameter(Mandatory)]$Data,
        [Parameter(Mandatory)][string]$Path,
        [switch]$Append  # <-- new switch
    )
    if (-not $Data -or ($Data | Measure-Object).Count -eq 0) { return }

    $first = $Data | Select-Object -First 1
    $cols = $first.PSObject.Properties | Where-Object { $_.MemberType -eq 'NoteProperty' } | Select-Object -ExpandProperty Name

    if (-not $Append) {
        # Overwrite only if not appending
        Set-Content -Path $Path -Value ($cols -join ',')
    } elseif (-not (Test-Path $Path)) {
        # If appending but file doesn't exist, write headers first
        Set-Content -Path $Path -Value ($cols -join ',')
    }

    foreach ($row in $Data) {
        $values = foreach ($c in $cols) {
            $v = $row.$c
            if ($null -eq $v) { '' } else { ($v.ToString() -replace '"','') -replace "[\r\n]",' ' }
        }
        Add-Content -Path $Path -Value ($values -join ',')
    }
}


# -------------------------
# Summary append helpers (appends to existing CSV files with blank lines then a section)
#  Add-BlankLines defined earlier with guard

function Add-VmInfoSummary {
    if (-not $VmData -or $VmData.Count -eq 0) { return }
    Add-BlankLines -Path $Path -Count 4
    Add-Content -Path $Path -Value ("Total VMs, {0}" -f $VmData.Count)
    $projGroups = $VmData | Group-Object Project | Sort-Object Name
    foreach ($pg in $projGroups) {
        Add-Content -Path $Path -Value ("Project, {0}, VMs, {1}" -f $pg.Name,$pg.Count)
    }
    # Per Project -> Region
    Add-Content -Path $Path -Value 'Per Project Region:'
    foreach ($pg in $projGroups) {
        $regionGroups = $pg.Group | Group-Object Region | Sort-Object Name
        foreach ($rg in $regionGroups) {
            Add-Content -Path $Path -Value ("ProjectRegion, {0}, {1}, VMs, {2}" -f $pg.Name,$rg.Name,$rg.Count)
        }
    }
    # Per Project -> Zone
    Add-Content -Path $Path -Value 'Per Project Zone:'
    foreach ($pg in $projGroups) {
        $zoneGroups = $pg.Group | Group-Object Zone | Sort-Object Name
        foreach ($zg in $zoneGroups) {
            Add-Content -Path $Path -Value ("ProjectZone, {0}, {1}, {2}, VMs, {3}" -f $pg.Name,($zg.Group | Select-Object -First 1).Region,$zg.Name,$zg.Count)
        }
    }
}


function Add-DiskSummary {
    param([string]$Path,[object[]]$DiskData,[string]$Title)
    if (-not $DiskData -or $DiskData.Count -eq 0) { return }
    Add-BlankLines -Path $Path -Count 4
    Add-Content -Path $Path -Value ("### {0} Summary ###" -f $Title)
    $uniqueDisks = $DiskData | Select-Object -Property DiskName,SizeGB,Project,Region,Zone -Unique
    $totalCount = ($uniqueDisks | Measure-Object).Count
    $totalGB = ($uniqueDisks | Measure-Object SizeGB -Sum).Sum
    Add-Content -Path $Path -Value ("Total Disks, {0}, TotalSizeGB, {1}" -f $totalCount,$totalGB)
    # Per Project
    Add-Content -Path $Path -Value 'Per Project:'
    $projGroups = $uniqueDisks | Group-Object Project | Sort-Object Name
    foreach ($pg in $projGroups) {
        $pGB = ($pg.Group | Measure-Object SizeGB -Sum).Sum
        Add-Content -Path $Path -Value ("Project, {0}, Disks, {1}, SizeGB, {2}" -f $pg.Name,$pg.Count,$pGB)
    }
    # Per Project Region
    Add-Content -Path $Path -Value 'Per Project Region:'
    foreach ($pg in $projGroups) {
        $regionGroups = $pg.Group | Group-Object Region | Sort-Object Name
        foreach ($rg in $regionGroups) {
            $rGB = ($rg.Group | Measure-Object SizeGB -Sum).Sum
            Add-Content -Path $Path -Value ("ProjectRegion, {0}, {1}, Disks, {2}, SizeGB, {3}" -f $pg.Name,$rg.Name,$rg.Count,$rGB)
        }
    }
}

function Add-BucketSummary {
    param([string]$Path,[object[]]$Buckets)
    if (-not $Buckets -or $Buckets.Count -eq 0) { return }
    Add-BlankLines -Path $Path -Count 4
    Add-Content -Path $Path -Value '### Storage Buckets Summary ###'
    $totalCount = $Buckets.Count
    $totalBytes = ($Buckets | Measure-Object UsedCapacityBytes -Sum).Sum
    $totalGB = [math]::Round($totalBytes/1e9,3)
    Add-Content -Path $Path -Value ("Total Buckets, {0}, TotalSizeGB, {1}" -f $totalCount,$totalGB)
    # Per Project
    Add-Content -Path $Path -Value 'Per Project:'
    $projGroups = $Buckets | Group-Object Project | Sort-Object Name
    foreach ($pg in $projGroups) {
        $pBytes = ($pg.Group | Measure-Object UsedCapacityBytes -Sum).Sum
        $pGB = [math]::Round($pBytes/1e9,3)
        Add-Content -Path $Path -Value ("Project, {0}, Buckets, {1}, SizeGB, {2}" -f $pg.Name,$pg.Count,$pGB)
    }
    # Per Project Location
    Add-Content -Path $Path -Value 'Per Project Location:'
    foreach ($pg in $projGroups) {
        $locGroups = $pg.Group | Group-Object Location | Sort-Object Name
        foreach ($lg in $locGroups) {
            $lBytes = ($lg.Group | Measure-Object UsedCapacityBytes -Sum).Sum
            $lGB = [math]::Round($lBytes/1e9,3)
            Add-Content -Path $Path -Value ("ProjectLocation, {0}, {1}, Buckets, {2}, SizeGB, {3}" -f $pg.Name,$lg.Name,$lg.Count,$lGB)
        }
    }
}

# Always generate all VM-related CSVs if VM inventory is selected and data exists
if ($Selected.VM -and $invResults.VMs -and $invResults.VMs.Count) {
    $vmCsv = Join-Path $outDir ("gcp_vm_instance_info_" + $dateStr + ".csv")
    if (Test-Path $vmCsv) { Remove-Item $vmCsv -Force }
    Add-Content -Path $vmCsv -Value '======= VM Instance Summary ======='
    Add-VmInfoSummary -Path $vmCsv -VmData $invResults.VMs
    Add-Content -Path $vmCsv -Value '======= VM Instance Details ======='
    Add-Content -Path $vmCsv -Value ("Project, VMInstanceName,VMSize,OS,Region,Zone,VMId,DiskCount,VMDiskSizeGB")
    Write-PlainCsv -Data $invResults.VMs -Path $vmCsv -Append
    Write-Host "VMs CSV written: $(Split-Path $vmCsv -Leaf)" -ForegroundColor Cyan

    if ($invResults.AttachedDisks -and $invResults.AttachedDisks.Count) {
        $attachedCsv = Join-Path $outDir ("gcp_disks_attached_to_vm_instances_" + $dateStr + ".csv")
    $attachedData = $invResults.AttachedDisks | Select-Object DiskName,VMName,Project,Region,Zone,IsRegional,Encrypted,DiskType,SizeGB
        Write-PlainCsv -Data $attachedData -Path $attachedCsv
        Write-Host "Attached disks CSV written: $(Split-Path $attachedCsv -Leaf)" -ForegroundColor Cyan
        Add-DiskSummary -Path $attachedCsv -DiskData $attachedData -Title 'Attached Disks'
    }

    if ($invResults.UnattachedDisks -and $invResults.UnattachedDisks.Count) {
        $unattachedCsv = Join-Path $outDir ("gcp_disks_unattached_to_vm_instances_" + $dateStr + ".csv")
    $unattachedData = $invResults.UnattachedDisks | Select-Object DiskName,VMName,Project,Region,Zone,IsRegional,Encrypted,DiskType,SizeGB
        Write-PlainCsv -Data $unattachedData -Path $unattachedCsv
        Write-Host "Unattached disks CSV written: $(Split-Path $unattachedCsv -Leaf)" -ForegroundColor Cyan
        Add-DiskSummary -Path $unattachedCsv -DiskData $unattachedData -Title 'Unattached Disks'
    }

    # Append VM Successful/Failed project sections
    if ($script:VmProjectStatuses) {
        Add-Content -Path $vmCsv -Value ''
        Add-Content -Path $vmCsv -Value 'SuccessfulProjects:'
        Add-Content -Path $vmCsv -Value 'Project,VMs,Disks'
        $vmSucc = $script:VmProjectStatuses | Where-Object { $_.Success -eq $true -or $_.Success -eq 'True' -or $_.Success -eq 'Y' }
        foreach ($s in ($vmSucc | Sort-Object Project)) { Add-Content -Path $vmCsv -Value ("{0},{1},{2}" -f $s.Project,$s.VMCount,$s.DiskCount) }
        $vmFail = $script:VmProjectStatuses | Where-Object { -not ($_.Success -eq $true -or $_.Success -eq 'True' -or $_.Success -eq 'Y') }
        if ($vmFail) {
            Add-Content -Path $vmCsv -Value ''
            Add-Content -Path $vmCsv -Value 'FailedProjects:'
            Add-Content -Path $vmCsv -Value 'Project,Reason'
            foreach ($f in ($vmFail | Sort-Object Project)) {
                $reason = if ($f.Timeout) { 'Timeout' } elseif ($f.PermissionIssue) { 'PermissionDenied' } elseif ($f.ApiDisabled) { 'ApiDisabled' } else { 'Unknown' }
                Add-Content -Path $vmCsv -Value ("{0},{1}" -f $f.Project,$reason)
            }
        }
    }
}

# Always generate Storage Bucket CSV if Storage inventory is selected and data exists
if ($Selected.STORAGE -and $invResults.StorageBuckets -and $invResults.StorageBuckets.Count) {
    $bktCsv = Join-Path $outDir ("gcp_storage_buckets_info_" + $dateStr + ".csv")
    Write-PlainCsv -Data $invResults.StorageBuckets -Path $bktCsv
    Write-Host "Buckets CSV written: $(Split-Path $bktCsv -Leaf)" -ForegroundColor Cyan
    Add-BucketSummary -Path $bktCsv -Buckets $invResults.StorageBuckets
    # Append Bucket Successful/Failed project sections
    if ($script:StorageProjectStatuses) {
        Add-Content -Path $bktCsv -Value ''
        Add-Content -Path $bktCsv -Value 'SuccessfulProjects:'
        Add-Content -Path $bktCsv -Value 'Project,Buckets,TotalSizeGB'
        $bucketGroups = $invResults.StorageBuckets | Group-Object Project
        # Backfill Success flag if older objects lack it but we have bucket data
        foreach ($st in $script:StorageProjectStatuses) {
            if (-not $st.PSObject.Properties.Name -contains 'Success') {
                $hasBuckets = ($invResults.StorageBuckets | Where-Object Project -eq $st.Project)
                $st | Add-Member -NotePropertyName Success -NotePropertyValue (if ($hasBuckets) { 'Y' } else { 'N' })
            }
        }
        $succStatuses = $script:StorageProjectStatuses | Where-Object { $_.Success -eq 'Y' }
        foreach ($s in ($succStatuses | Sort-Object Project)) {
            $grp = $bucketGroups | Where-Object Name -eq $s.Project
            $sizeBytes = if ($grp) { ($grp.Group | Measure-Object UsedCapacityBytes -Sum).Sum } else { 0 }
            $sizeGB = if ($sizeBytes) { [math]::Round($sizeBytes/1e9,3) } else { 0 }
            Add-Content -Path $bktCsv -Value ("{0},{1},{2}" -f $s.Project,$s.BucketCount,$sizeGB)
        }
        $failStatuses = $script:StorageProjectStatuses | Where-Object { $_.Success -ne 'Y' }
        if ($failStatuses) {
            Add-Content -Path $bktCsv -Value ''
            Add-Content -Path $bktCsv -Value 'FailedProjects:'
            Add-Content -Path $bktCsv -Value 'Project,Reason'
            foreach ($f in ($failStatuses | Sort-Object Project)) {
                $reason = if ($f.Timeout -eq 'Y') { 'Timeout' } elseif ($f.PermissionIssue -eq 'Y') { 'PermissionDenied' } elseif ($f.Error -and $f.Error -match 'permission|denied') { 'PermissionDenied' } elseif ($f.Error -and $f.Error -match 'Timeout') { 'Timeout' } elseif ($f.Error) { 'Error' } else { 'Unknown' }
                Add-Content -Path $bktCsv -Value ("{0},{1}" -f $f.Project,$reason)
            }
        }
    }
}

# Always generate Filestore CSV if Filestore inventory is selected and data exists
if ($Selected.FILESHARE -and $invResults.FileShares -and (@($invResults.FileShares)).Count) {
    $fsCsv = Join-Path $outDir ("gcp_filestore_info_" + $dateStr + ".csv")
    $allFs = @($invResults.FileShares)  # normalize enumeration
    $success = @($allFs | Where-Object { -not ($_.Error -and $_.Error.Trim()) -and ($_.ShareName -ne '(error)') -and ($_.State -ne 'ERROR') })
    $fail = @($allFs | Where-Object { ($_.Error -and $_.Error.Trim()) -or ($_.ShareName -eq '(error)') -or ($_.State -eq 'ERROR') })
    $successCount = (@($success)).Count
    if ($successCount -gt 0) {
        $success = $success | Sort-Object Project, InstanceName, ShareName
        Write-PlainCsv -Data $success -Path $fsCsv
    } else {
    $dummy = [PSCustomObject]@{ Project=''; InstanceName=''; ShareName=''; Tier=''; Region=''; Zone=''; provisionedGb=''; provisionedGib=''; provisionedTb=''; provisionedTib=''; Networks=''; IPAddresses=''; State=''; CreateTime=''; Labels=''; Protocol=''; Error='' }
        Write-PlainCsv -Data @($dummy) -Path $fsCsv
        (Get-Content $fsCsv | Select-Object -First 1) | Set-Content $fsCsv
    }
    # Insert SuccessfulProjects summary (only if we actually had successes)
    if ($successCount -gt 0) {
        Add-Content -Path $fsCsv -Value ''
        Add-Content -Path $fsCsv -Value 'SuccessfulProjects:'
        Add-Content -Path $fsCsv -Value 'Project,Shares,TotalProvisionedGiB,TotalProvisionedGB,TotalProvisionedTB,TotalProvisionedTiB'
        foreach ($grp in ($success | Group-Object Project | Sort-Object Name)) {
            $sumGib = (($grp.Group | Measure-Object provisionedGib -Sum).Sum)
            if (-not $sumGib) { $sumGib = 0 }
            $sumGb  = [math]::Round(($sumGib * 1024 * 1024 * 1024)/1e9,3)
            $sumTib = [math]::Round($sumGib/1024,4)
            $sumTb  = [math]::Round($sumGb/1000,4)
            Add-Content -Path $fsCsv -Value ("{0},{1},{2},{3},{4},{5}" -f $grp.Name,$grp.Count,[int64]$sumGib,$sumGb,$sumTb,$sumTib)
        }
    }

    if ((@($fail)).Count -gt 0) {
        Add-Content -Path $fsCsv -Value ''
        Add-Content -Path $fsCsv -Value 'FailedProjects:'
        Add-Content -Path $fsCsv -Value 'Project,Error'
        foreach ($e in ($fail | Sort-Object Project -Unique)) {
            $msg=$e.Error
            if ($msg -match '(?s)(ERROR:.*?)(Activation|Google developers|Cloud Filestore API has not been used|$)') { $msg = $Matches[1] }
            $first = ($msg -split "`n")[0]
            if ($first.Length -gt 260) { $first=$first.Substring(0,257)+'...' }
            $first=$first -replace ',',';'
            Add-Content -Path $fsCsv -Value ("{0},{1}" -f $e.Project,$first)
        }
        # Add a more granular section listing individual failed shares with instance & share name if available
        Add-Content -Path $fsCsv -Value ''
        Add-Content -Path $fsCsv -Value 'FailedShares:'
        Add-Content -Path $fsCsv -Value 'Project,InstanceName,ShareName,State,ErrorSnippet'
        foreach ($e in ($fail | Sort-Object Project, InstanceName, ShareName)) {
            $emsg = $e.Error
            if ($emsg -match '(?s)(ERROR:.*?)(Activation|Google developers|Cloud Filestore API has not been used|$)') { $emsg = $Matches[1] }
            $line = ($emsg -split "`n")[0]
            if ($line.Length -gt 180) { $line = $line.Substring(0,177)+'...' }
            $line = $line -replace ',',';'
            Add-Content -Path $fsCsv -Value ("{0},{1},{2},{3},{4}" -f $e.Project,$e.InstanceName,$e.ShareName,$e.State,$line)
        }
    }
    Write-Host "Filestore CSV written: $(Split-Path $fsCsv -Leaf) (Good=$successCount Fail=$((@($fail)).Count))" -ForegroundColor Cyan
}

# GKE clusters CSV
if ($Selected.GKE -and $invResults.GKEClusters -and $invResults.GKEClusters.Count) {
    $gkeCsv = Join-Path $outDir ("gcp_gke_clusters_" + $dateStr + ".csv")
    # Order & columns
    $gkeData = $invResults.GKEClusters | Select-Object Project,ClusterName,Region,Zone,Location,KubernetesVer,NodeCount,EstimatedvCPU,EstimatedMemGB,PersistentVolumeCapacityGB,PersistentVolumeCount,PersistentVolumeSource,ReleaseChannel,Network,Subnetwork,Endpoint
    Write-PlainCsv -Data $gkeData -Path $gkeCsv
    # Summary append
    Add-Content -Path $gkeCsv -Value ''
    Add-Content -Path $gkeCsv -Value 'Summary:'
    $totalClusters = $gkeData.Count
    $sumNodes = ($gkeData | Measure-Object NodeCount -Sum).Sum
    $sumvCPU = ($gkeData | Measure-Object EstimatedvCPU -Sum).Sum
    $sumMem  = ($gkeData | Measure-Object EstimatedMemGB -Sum).Sum
    $sumPvCap = ($gkeData | Measure-Object PersistentVolumeCapacityGB -Sum).Sum
    $sumPvCnt = ($gkeData | Measure-Object PersistentVolumeCount -Sum).Sum
    Add-Content -Path $gkeCsv -Value ("TotalClusters,{0},TotalNodes,{1},TotalvCPU,{2},TotalMemGB,{3},TotalPVCapGB,{4},TotalPVCount,{5}" -f $totalClusters,$sumNodes,$sumvCPU,$sumMem,$sumPvCap,$sumPvCnt)
    Add-Content -Path $gkeCsv -Value 'Per Project:'
    foreach ($grp in ($gkeData | Group-Object Project | Sort-Object Name)) {
        $nNodes = ($grp.Group | Measure-Object NodeCount -Sum).Sum
        $nVcpu  = ($grp.Group | Measure-Object EstimatedvCPU -Sum).Sum
        $nMem   = ($grp.Group | Measure-Object EstimatedMemGB -Sum).Sum
        $nPvCap = ($grp.Group | Measure-Object PersistentVolumeCapacityGB -Sum).Sum
        $nPvCnt = ($grp.Group | Measure-Object PersistentVolumeCount -Sum).Sum
    Add-Content -Path $gkeCsv -Value ("Project,{0},Clusters,{1},Nodes,{2},vCPU,{3},MemGB,{4},PVCapGB,{5},PVCount,{6}" -f $grp.Name,$grp.Count,$nNodes,$nVcpu,$nMem,$nPvCap,$nPvCnt)
    }
    Write-Host "GKE clusters CSV written: $(Split-Path $gkeCsv -Leaf)" -ForegroundColor Cyan

    # PV CSV
    if ($invResults.GKEPVs -and $invResults.GKEPVs.Count) {
        $pvCsv = Join-Path $outDir ("gcp_gke_persistent_volumes_" + $dateStr + ".csv")
        $pvOrdered = $invResults.GKEPVs | Sort-Object Project,ClusterName,ClaimNamespace,PVName
        Write-PlainCsv -Data $pvOrdered -Path $pvCsv
        Add-Content -Path $pvCsv -Value ''
        Add-Content -Path $pvCsv -Value 'Summary:'
        $totalPv = $pvOrdered.Count
        $sumPvCap = ($pvOrdered | Measure-Object CapacityGiB -Sum).Sum
        Add-Content -Path $pvCsv -Value ("TotalPVs,{0},TotalCapacityGiB,{1}" -f $totalPv,$sumPvCap)
        Add-Content -Path $pvCsv -Value 'Per Cluster:'
        foreach ($grp in ($pvOrdered | Group-Object Project,ClusterName)) {
            $cap = ($grp.Group | Measure-Object CapacityGiB -Sum).Sum
            $first = $grp.Group | Select-Object -First 1
            Add-Content -Path $pvCsv -Value ("Project,{0},Cluster,{1},PVs,{2},CapacityGiB,{3}" -f $first.Project,$first.ClusterName,$grp.Count,$cap)
        }
        Write-Host "GKE PV CSV written: $(Split-Path $pvCsv -Leaf)" -ForegroundColor Cyan
    }
    # PVC CSV
    if ($invResults.GKEPVCs -and $invResults.GKEPVCs.Count) {
        $pvcCsv = Join-Path $outDir ("gcp_gke_persistent_volume_claims_" + $dateStr + ".csv")
        $pvcOrdered = $invResults.GKEPVCs | Sort-Object Project,ClusterName,Namespace,PVCName
        Write-PlainCsv -Data $pvcOrdered -Path $pvcCsv
        Add-Content -Path $pvcCsv -Value ''
        Add-Content -Path $pvcCsv -Value 'Summary:'
        $totalPvc = $pvcOrdered.Count
        $sumReq = ($pvcOrdered | Measure-Object RequestedGiB -Sum).Sum
        $sumCap = ($pvcOrdered | Measure-Object CapacityGiB -Sum).Sum
        Add-Content -Path $pvcCsv -Value ("TotalPVCs,{0},TotalRequestedGiB,{1},TotalCapacityGiB,{2}" -f $totalPvc,$sumReq,$sumCap)
        Add-Content -Path $pvcCsv -Value 'Per Cluster:'
        foreach ($grp in ($pvcOrdered | Group-Object Project,ClusterName)) {
            $req = ($grp.Group | Measure-Object RequestedGiB -Sum).Sum
            $cap = ($grp.Group | Measure-Object CapacityGiB -Sum).Sum
            $first = $grp.Group | Select-Object -First 1
            Add-Content -Path $pvcCsv -Value ("Project,{0},Cluster,{1},PVCs,{2},RequestedGiB,{3},CapacityGiB,{4}" -f $first.Project,$first.ClusterName,$grp.Count,$req,$cap)
        }
        Add-Content -Path $pvcCsv -Value 'Per Cluster Namespace:'
        foreach ($grp in ($pvcOrdered | Group-Object Project,ClusterName,Namespace)) {
            $req = ($grp.Group | Measure-Object RequestedGiB -Sum).Sum
            $cap = ($grp.Group | Measure-Object CapacityGiB -Sum).Sum
            $first = $grp.Group | Select-Object -First 1
            Add-Content -Path $pvcCsv -Value ("Project,{0},Cluster,{1},Namespace,{2},PVCs,{3},RequestedGiB,{4},CapacityGiB,{5}" -f $first.Project,$first.ClusterName,$first.Namespace,$grp.Count,$req,$cap)
        }
        Write-Host "GKE PVC CSV written: $(Split-Path $pvcCsv -Leaf)" -ForegroundColor Cyan
    }
}



# -------------------------
# Build Summary (Custom ordering with spacer rows)
# Order:
# 1. Overall (VM + Buckets)
# 2. 4 blank rows
# 3. Project-level (VM + Buckets)
# 4. 4 blank rows
# 5. For each project: Region-level VM rows, 2 blank rows, Zone-level VM rows
function New-BlankSummaryRow { return [PSCustomObject]@{Level='';ResourceType='';Project='';Region='';Zone='';Count='';TotalSizeGB='';TotalSizeTB='';TotalSizeTiB=''} }

$summaryRows = @()

$vmData = $invResults.VMs
$attached = $invResults.AttachedDisks
$bucketData = $invResults.StorageBuckets
$fileshareData = $invResults.FileShares
if ($fileshareData) { $fileshareData = $fileshareData | Where-Object { -not ($_.Error -and $_.Error.Trim()) } }
$databaseData = $invResults.Databases
if ($databaseData) { $databaseData = $databaseData | Where-Object { -not ($_.Error -and $_.Error.Trim()) -and $_.Type -ne 'Error' } }
 $gkeData = $invResults.GKEClusters

# Fallback: if user requested only DB (or DB included) and nothing discovered, capture zero rows so summary isn't just blank spacers
if ($Selected.DB -and (-not $databaseData -or $databaseData.Count -eq 0)) {
    # Overall zero row
    $summaryRows += [PSCustomObject]@{Level='Overall';ResourceType='Database';Project='All';Region='All';Zone='All';Count=0;TotalSizeGB=0;TotalSizeTB=0;TotalSizeTiB=0}
    # Per project zero rows (maintains consistency with other resource summaries)
    foreach ($proj in $targetProjects) {
        $summaryRows += [PSCustomObject]@{Level='Project';ResourceType='Database';Project=$proj;Region='All';Zone='All';Count=0;TotalSizeGB=0;TotalSizeTB=0;TotalSizeTiB=0}
    }
}

# Overall rows
if ($Selected.VM -and $vmData) {
    $overallDiskSizeGB = ($attached | Select-Object -Property DiskName,SizeGB -Unique | Measure-Object SizeGB -Sum).Sum
}
if ($Selected.STORAGE -and $bucketData) {
    $overallBucketBytes = ($bucketData | Measure-Object UsedCapacityBytes -Sum).Sum
}

# Combined overall row (only if both resource types selected and present)
if (($Selected.VM -and $vmData) -and ($Selected.STORAGE -and $bucketData)) {
    if (-not $overallDiskSizeGB) { $overallDiskSizeGB = 0 }
    if (-not $overallBucketBytes) { $overallBucketBytes = 0 }
    $overallBucketGB = [math]::Round($overallBucketBytes/1e9,3)
    $combinedGB = [math]::Round(($overallDiskSizeGB + $overallBucketGB),3)
    $summaryRows += [PSCustomObject]@{Level='Overall';ResourceType='AllResources';Project='All';Region='All';Zone='All';Count='N/A';TotalSizeGB=$combinedGB;TotalSizeTB=[math]::Round($combinedGB/1e3,4);TotalSizeTiB=[math]::Round($combinedGB/1024,4)}
}

# Individual overall rows
if ($Selected.VM -and $vmData) {
    if (-not $overallDiskSizeGB) { $overallDiskSizeGB = 0 }
    $summaryRows += [PSCustomObject]@{Level='Overall';ResourceType='VM';Project='All';Region='All';Zone='All';Count=$vmData.Count;TotalSizeGB=[int64]$overallDiskSizeGB;TotalSizeTB=[math]::Round($overallDiskSizeGB/1e3,4);TotalSizeTiB=[math]::Round($overallDiskSizeGB/1024,4)}
}
if ($Selected.STORAGE -and $bucketData) {
    if (-not $overallBucketBytes) { $overallBucketBytes = 0 }
    $summaryRows += [PSCustomObject]@{Level='Overall';ResourceType='StorageBucket';Project='All';Region='All';Zone='All';Count=$bucketData.Count;TotalSizeGB=[math]::Round($overallBucketBytes/1e9,3);TotalSizeTB=[math]::Round($overallBucketBytes/1e12,4);TotalSizeTiB=[math]::Round(($overallBucketBytes/1GB)/1024,4)}
}
if ($Selected.FILESHARE -and $fileshareData) {
    $overallFsGiB = ($fileshareData | Measure-Object provisionedGib -Sum).Sum
    if (-not $overallFsGiB) { $overallFsGiB = 0 }
    $overallFsGB  = [math]::Round(($overallFsGiB * 1024 * 1024 * 1024)/1e9,3)
    $overallFsTiB = [math]::Round($overallFsGiB/1024,4)
    $overallFsTB  = [math]::Round($overallFsGB/1000,4)
    $summaryRows += [PSCustomObject]@{Level='Overall';ResourceType='Fileshare';Project='All';Region='All';Zone='All';Count=(@($fileshareData)).Count;TotalSizeGB=$overallFsGB;TotalSizeTB=$overallFsTB;TotalSizeTiB=$overallFsTiB}
}
if ($Selected.DB -and $databaseData) {
    $overallDbGB = ($databaseData | Measure-Object StorageGB -Sum).Sum
    if (-not $overallDbGB) { $overallDbGB = 0 }
    $summaryRows += [PSCustomObject]@{Level='Overall';ResourceType='Database';Project='All';Region='All';Zone='All';Count=$databaseData.Count;TotalSizeGB=[math]::Round($overallDbGB,3);TotalSizeTB=[math]::Round($overallDbGB/1e3,4);TotalSizeTiB=[math]::Round($overallDbGB/1024,4)}
}
if ($Selected.GKE -and $gkeData) {
    $overallGkePvCap = ($gkeData | Measure-Object PersistentVolumeCapacityGB -Sum).Sum
    if (-not $overallGkePvCap) { $overallGkePvCap = 0 }
    $summaryRows += [PSCustomObject]@{Level='Overall';ResourceType='GKECluster';Project='All';Region='All';Zone='All';Count=$gkeData.Count;TotalSizeGB=[math]::Round($overallGkePvCap,2);TotalSizeTB=[math]::Round($overallGkePvCap/1e3,4);TotalSizeTiB=[math]::Round($overallGkePvCap/1024,4)}
}

# 4 spacer rows
1..4 | ForEach-Object { $summaryRows += (New-BlankSummaryRow) }

# Project-level rows
if ($Selected.VM -and $vmData) {
    foreach ($proj in $targetProjects) {
        $projVMs = $vmData | Where-Object Project -eq $proj
        if (-not $projVMs) { continue }
        $projDisks = $attached | Where-Object Project -eq $proj | Select-Object DiskName,SizeGB -Unique
        $projGB = ($projDisks | Measure-Object SizeGB -Sum).Sum
        $summaryRows += [PSCustomObject]@{Level='Project';ResourceType='VM';Project=$proj;Region='All';Zone='All';Count=$projVMs.Count;TotalSizeGB=[int64]$projGB;TotalSizeTB=[math]::Round($projGB/1e3,4);TotalSizeTiB=[math]::Round($projGB/1024,4)}
    }
}
if ($Selected.STORAGE -and $bucketData) {
    foreach ($proj in $targetProjects) {
        $projBuckets = $bucketData | Where-Object Project -eq $proj
        if (-not $projBuckets) { continue }
        $projBytes = ($projBuckets | Measure-Object UsedCapacityBytes -Sum).Sum
        $summaryRows += [PSCustomObject]@{Level='Project';ResourceType='StorageBucket';Project=$proj;Region='All';Zone='All';Count=$projBuckets.Count;TotalSizeGB=[math]::Round($projBytes/1e9,3);TotalSizeTB=[math]::Round($projBytes/1e12,4);TotalSizeTiB=[math]::Round(($projBytes/1GB)/1024,4)}
    }
}
if ($Selected.FILESHARE -and $fileshareData) {
    foreach ($proj in $targetProjects) {
        $projFS = $fileshareData | Where-Object Project -eq $proj
        if (-not $projFS) { continue }
        $projFsGiB = ($projFS | Measure-Object provisionedGib -Sum).Sum
        if (-not $projFsGiB) { $projFsGiB = 0 }
        $projFsGB = [math]::Round(($projFsGiB * 1024 * 1024 * 1024)/1e9,3)
        $projFsTiB = [math]::Round($projFsGiB/1024,4)
        $projFsTB  = [math]::Round($projFsGB/1000,4)
        $summaryRows += [PSCustomObject]@{Level='Project';ResourceType='Fileshare';Project=$proj;Region='All';Zone='All';Count=(@($projFS)).Count;TotalSizeGB=$projFsGB;TotalSizeTB=$projFsTB;TotalSizeTiB=$projFsTiB}
    }
}
if ($Selected.DB -and $databaseData) {
    foreach ($proj in $targetProjects) {
        $projDB = $databaseData | Where-Object Project -eq $proj
        if (-not $projDB) { continue }
        $projDbGB = ($projDB | Measure-Object StorageGB -Sum).Sum
        if (-not $projDbGB) { $projDbGB = 0 }
        $summaryRows += [PSCustomObject]@{Level='Project';ResourceType='Database';Project=$proj;Region='All';Zone='All';Count=$projDB.Count;TotalSizeGB=[math]::Round($projDbGB,3);TotalSizeTB=[math]::Round($projDbGB/1e3,4);TotalSizeTiB=[math]::Round($projDbGB/1024,4)}
    }
}
if ($Selected.GKE -and $gkeData) {
    foreach ($proj in $targetProjects) {
        $projGke = $gkeData | Where-Object Project -eq $proj
        if (-not $projGke) { continue }
    $projGkePvCap = ($projGke | Measure-Object PersistentVolumeCapacityGB -Sum).Sum
    if (-not $projGkePvCap) { $projGkePvCap = 0 }
    $summaryRows += [PSCustomObject]@{Level='Project';ResourceType='GKECluster';Project=$proj;Region='All';Zone='All';Count=$projGke.Count;TotalSizeGB=[math]::Round($projGkePvCap,2);TotalSizeTB=[math]::Round($projGkePvCap/1e3,4);TotalSizeTiB=[math]::Round($projGkePvCap/1024,4)}
    }
}

# 4 spacer rows
1..4 | ForEach-Object { $summaryRows += (New-BlankSummaryRow) }
# -------------------------
# NEW: Cumulative Region-level rows across ALL projects (placed after project-level)
# For VMs: sum distinct disks per region to avoid double counting
# For Buckets: group by Location (mapped to Region column)
# -------------------------
if ($Selected.VM -and $vmData) {
    $regionGroupsAll = $vmData | Group-Object Region | Sort-Object Name
    foreach ($rg in $regionGroupsAll) {
        $region = $rg.Name
        # Distinct disks by name in this region
        $regionDisksAll = $attached | Where-Object { $_.Region -eq $region } | Select-Object DiskName,SizeGB -Unique
        $regionGBAll = ($regionDisksAll | Measure-Object SizeGB -Sum).Sum
        if (-not $regionGBAll) { $regionGBAll = 0 }
        $summaryRows += [PSCustomObject]@{Level='Region';ResourceType='VM';Project='All';Region=$region;Zone='All';Count=$rg.Count;TotalSizeGB=[int64]$regionGBAll;TotalSizeTB=[math]::Round($regionGBAll/1e3,4);TotalSizeTiB=[math]::Round($regionGBAll/1024,4)}
    }
}
if ($Selected.STORAGE -and $bucketData) {
    $locGroupsAll = $bucketData | Group-Object Location | Sort-Object Name
    foreach ($lg in $locGroupsAll) {
        $locBytes = ($lg.Group | Measure-Object UsedCapacityBytes -Sum).Sum
        if (-not $locBytes) { $locBytes = 0 }
        $summaryRows += [PSCustomObject]@{Level='Region';ResourceType='StorageBucket';Project='All';Region=$lg.Name;Zone='All';Count=$lg.Count;TotalSizeGB=[math]::Round($locBytes/1e9,3);TotalSizeTB=[math]::Round($locBytes/1e12,4);TotalSizeTiB=[math]::Round(($locBytes/1GB)/1024,4)}
    }
}
if ($Selected.FILESHARE -and $fileshareData) {
    $fsRegionGroups = $fileshareData | Group-Object Region | Sort-Object Name
    foreach ($rg in $fsRegionGroups) {
        $rGiB = ($rg.Group | Measure-Object provisionedGib -Sum).Sum
        if (-not $rGiB) { $rGiB = 0 }
        $rGB  = [math]::Round(($rGiB * 1024 * 1024 * 1024)/1e9,3)
        $rTiB = [math]::Round($rGiB/1024,4)
        $rTB  = [math]::Round($rGB/1000,4)
        $summaryRows += [PSCustomObject]@{Level='Region';ResourceType='Fileshare';Project='All';Region=$rg.Name;Zone='All';Count=$rg.Count;TotalSizeGB=$rGB;TotalSizeTB=$rTB;TotalSizeTiB=$rTiB}
    }
}
if ($Selected.DB -and $databaseData) {
    $dbRegionGroups = $databaseData | Group-Object Region | Sort-Object Name
    foreach ($rg in $dbRegionGroups) {
        $rGb = ($rg.Group | Measure-Object StorageGB -Sum).Sum
        if (-not $rGb) { $rGb = 0 }
        $summaryRows += [PSCustomObject]@{Level='Region';ResourceType='Database';Project='All';Region=$rg.Name;Zone='All';Count=$rg.Count;TotalSizeGB=[math]::Round($rGb,3);TotalSizeTB=[math]::Round($rGb/1e3,4);TotalSizeTiB=[math]::Round($rGb/1024,4)}
    }
}
if ($Selected.GKE -and $gkeData) {
    $gkeRegionGroups = $gkeData | Group-Object Region | Sort-Object Name
    foreach ($rg in $gkeRegionGroups) {
    $rgGkePvCap = ($rg.Group | Measure-Object PersistentVolumeCapacityGB -Sum).Sum
    if (-not $rgGkePvCap) { $rgGkePvCap = 0 }
    $summaryRows += [PSCustomObject]@{Level='Region';ResourceType='GKECluster';Project='All';Region=$rg.Name;Zone='All';Count=$rg.Count;TotalSizeGB=[math]::Round($rgGkePvCap,2);TotalSizeTB=[math]::Round($rgGkePvCap/1e3,4);TotalSizeTiB=[math]::Round($rgGkePvCap/1024,4)}
    }
}

# 4 spacer rows
1..4 | ForEach-Object { $summaryRows += (New-BlankSummaryRow) }

# Per project region + zone breakdown (VMs only)
if ($Selected.VM -and $vmData) {
    foreach ($proj in $targetProjects) {
        $projVMs = $vmData | Where-Object Project -eq $proj
        if (-not $projVMs) { continue }
    
    # Header row indicating upcoming region/zone breakdown for this project
    $summaryRows += [PSCustomObject]@{Level="Per region/zone in project [$proj]";ResourceType='';Project='';Region='';Zone='';Count='';TotalSizeGB='';TotalSizeTB='';TotalSizeTiB=''}
    
        # Regions
        $regionGroups = $projVMs | Group-Object Region | Sort-Object Name
        foreach ($rg in $regionGroups) {
            $region = $rg.Name
            $regionDisks = $attached | Where-Object { $_.Project -eq $proj -and $_.Region -eq $region } | Select-Object DiskName,SizeGB -Unique
            $regionGB = ($regionDisks | Measure-Object SizeGB -Sum).Sum
            $summaryRows += [PSCustomObject]@{Level='Region';ResourceType='VM';Project=$proj;Region=$region;Zone='All';Count=$rg.Count;TotalSizeGB=[int64]$regionGB;TotalSizeTB=[math]::Round($regionGB/1e3,4);TotalSizeTiB=[math]::Round($regionGB/1024,4)}
        }

        # 2 spacer rows between region and zone section
        1..2 | ForEach-Object { $summaryRows += (New-BlankSummaryRow) }

        # Zones (group ensures correct counts; avoids missing zone counts)
        $zoneGroups = $projVMs | Group-Object Zone | Sort-Object Name
        foreach ($zg in $zoneGroups) {
            $zone = $zg.Name
            $zoneDisks = $attached | Where-Object { $_.Project -eq $proj -and $_.Zone -eq $zone } | Select-Object DiskName,SizeGB -Unique
            $zoneGB = ($zoneDisks | Measure-Object SizeGB -Sum).Sum
            $summaryRows += [PSCustomObject]@{Level='Zone';ResourceType='VM';Project=$proj;Region=(($projVMs | Where-Object Zone -eq $zone | Select-Object -First 1).Region);Zone=$zone;Count=$zg.Count;TotalSizeGB=[int64]$zoneGB;TotalSizeTB=[math]::Round($zoneGB/1e3,4);TotalSizeTiB=[math]::Round($zoneGB/1024,4)}
        }

        # Spacer between projects (optional single blank row)
        $summaryRows += (New-BlankSummaryRow)
    }
}

$summaryCsv = Join-Path $outDir ("gcp_inventory_summary_" + $dateStr + ".csv")
Write-PlainCsv -Data $summaryRows -Path $summaryCsv
Write-Host "Inventory summary exported: $(Split-Path $summaryCsv -Leaf)" -ForegroundColor Green

# -------------------------
# Finalize log, then ZIP
# -------------------------
Write-Progress -Id 5 -Activity "Generating Output Files" -Status "Finalizing log..." -PercentComplete 75
Stop-Transcript | Out-Null   # end transcript (separate from detail log)


Write-Progress -Id 5 -Activity "Generating Output Files" -Status "Creating ZIP archive..." -PercentComplete 90
$zipFile = Join-Path $PWD ("gcp_sizing_" + $dateStr + ".zip")
Add-Type -AssemblyName System.IO.Compression.FileSystem

try {
    [IO.Compression.ZipFile]::CreateFromDirectory($outDir, $zipFile)
    Write-Host "ZIP archive created: $zipFile" -ForegroundColor Green
} catch {
    Write-Warning "Failed to create ZIP archive: $_"
}

Write-Progress -Id 5 -Activity "Generating Output Files" -Status "Cleaning up..." -PercentComplete 95
try {
    Remove-Item -Path $outDir -Recurse -Force
    Write-Host "Temporary directory removed: $outDir" -ForegroundColor Green
} catch {
    Write-Warning "Cleanup failed (directory may be locked): $_"
}

Write-Progress -Id 5 -Activity "Generating Output Files" -Completed
Write-Host "`nInventory complete. Results in $zipFile`n" -ForegroundColor Green
Write-Host "All output files (including the log) are compressed into the ZIP archive." -ForegroundColor Cyan
