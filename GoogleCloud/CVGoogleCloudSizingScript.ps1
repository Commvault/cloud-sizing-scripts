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
    Optional. Restrict inventory to specific resource types. Valid values: VM, Storage.
    If omitted, both VMs and Storage will be inventoried.

.PARAMETER Projects
    Optional. Target specific GCP projects by name or ID. If omitted, all accessible projects will be processed.

.OUTPUTS
    Creates timestamped output directory with:
    - gcp_vm_info_YYYY-MM-DD_HHMMSS.csv
    - gcp_disks_attached_to_vms_YYYY-MM-DD_HHMMSS.csv
    - gcp_disks_unattached_to_vms_YYYY-MM-DD_HHMMSS.csv
    - gcp_storage_buckets_info_YYYY-MM-DD_HHMMSS.csv
    - gcp_inventory_summary_YYYY-MM-DD_HHMMSS.csv
    - gcp_sizing_script_output_YYYY-MM-DD_HHMMSS.log
    - gcp_sizing_YYYY-MM-DD_HHMMSS.zip

.NOTES
    Requires Google Cloud SDK (gcloud CLI and gsutil) installed and authenticated.
    Must be run by a user with appropriate GCP permissions.
#>


<#
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

5. Run the script:
    ./CVGoogleCloudSizingScript.ps1
    ./CVGoogleCloudSizingScript.ps1 -Types VM,Storage
    ./CVGoogleCloudSizingScript.ps1 -Projects "my-gcp-project-1","my-gcp-project-2"

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
    .\CVGoogleCloudSizingScript.ps1
    .\CVGoogleCloudSizingScript.ps1 -Types VM
    .\CVGoogleCloudSizingScript.ps1 -Projects "my-gcp-project"

EXAMPLE USAGE
-------------
     .\CVGoogleCloudSizingScript.ps1
     # Inventories VMs and Storage Buckets in all accessible projects

     .\CVGoogleCloudSizingScript.ps1 -Types VM,Storage
     # Explicitly inventories VMs and Storage Buckets in all projects (same as default)

     .\CVGoogleCloudSizingScript.ps1 -Types VM
     # Only inventories Compute Engine VMs in all projects

     .\CVGoogleCloudSizingScript.ps1 -Projects "my-gcp-project-1","my-gcp-project-2"
     # Inventories VMs and Storage Buckets in only the specified projects

     .\CVGoogleCloudSizingScript.ps1 -Types Storage -Projects "my-gcp-project-1"
     # Only inventories Storage Buckets in the specified project
#>

param(
    [ValidateSet("VM","Storage", IgnoreCase=$true)]
    [string[]]$Types,
    [string[]]$Projects
)

# Minimal logging mode (set to $false to re-enable verbose progress lines)
$MinimalOutput = $true

# -------------------------
# Setup output + transcript
# -------------------------
$dateStr = (Get-Date).ToString("yyyy-MM-dd_HHmmss")
$outDir = Join-Path -Path $PWD -ChildPath ("gcp-inv-" + $dateStr)
New-Item -Path $outDir -ItemType Directory -Force | Out-Null

$transcriptFile = Join-Path $outDir ("gcp_sizing_script_output_" + $dateStr + ".log")
Start-Transcript -Path $transcriptFile -Append | Out-Null

# Global concurrent queue to collect log lines from child runspaces (VM & Bucket)
if (-not (Get-Variable -Name ChildLogQueue -Scope Global -ErrorAction SilentlyContinue)) {
    $Global:ChildLogQueue = [System.Collections.Concurrent.ConcurrentQueue[string]]::new()
}


Write-Host "=== GCP Resource Inventory Started ===" -ForegroundColor Green
if ($Types)    { Write-Host "  Types: $($Types -join ', ')" -ForegroundColor Green }
if ($Projects) { Write-Host "  Projects: $($Projects -join ', ')" -ForegroundColor Green }


# Resource type mapping
$ResourceTypeMap = @{
    "VM"      = "VMs"
    "STORAGE" = "StorageBuckets"
}

# Normalize types
if ($Types) {
    $Types = $Types | ForEach-Object { $_.Trim().ToUpper() }
    $Selected = @{}
    foreach ($t in $Types) {
        if ($ResourceTypeMap.ContainsKey($t)) { $Selected[$t] = $true }
    }
    if ($Selected.Count -eq 0) {
        Write-Host "No valid -Types specified. Use: VM, Storage"
        exit 1
    }
} else {
    $Selected = @{}
    $ResourceTypeMap.Keys | ForEach-Object { $Selected[$_] = $true }
}

# -------------------------
# Helpers
# -------------------------
function Get-GcpProjects {
    try {
        $json = gcloud projects list --format=json | ConvertFrom-Json
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

# -------------------------
# Lightweight logger (previous script referenced Write-Log but it wasn't defined)
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
}



# -------------------------
# VM + Disk Inventory (fast)
# -------------------------
function Get-GcpVMInventory {
    param([string[]]$ProjectIds)
    # ScriptBlock executed per project (returns inventory + log lines)
    $vmProjectScriptBlock = {
        param($proj,$minimalFlag)
        $log = New-Object System.Collections.Generic.List[string]
        $startUtc = [DateTime]::UtcNow
        $log.Add("[VM-Project-Start] $proj") | Out-Null
        function Get-RegionFromZoneInner { param([string]$zone) if (-not $zone) { return 'Unknown' }; $z = $zone -replace '.*/',''; return ($z -replace '-[a-z]$','') }
        $apiDisabled = $false
        # Always suppress interactive prompts
        $env:CLOUDSDK_CORE_DISABLE_PROMPTS = '1'
        # Instances
        try {
            $vmRaw = & gcloud --quiet compute instances list --project $proj --format=json 2>&1
            if ($LASTEXITCODE -ne 0) {
                $msg = ($vmRaw | Out-String).Trim()
                if ($msg -match '(?i)not enabled|has not been used|is disabled|API .* not enabled') { $apiDisabled = $true }
                $log.Add("[VM-Project-Warn] $proj instances list failed exit=$LASTEXITCODE msg=$msg") | Out-Null
                $vmList = @()
            } else {
                if ([string]::IsNullOrWhiteSpace(($vmRaw|Out-String))) { $vmList=@() } else { try { $vmList = $vmRaw | ConvertFrom-Json } catch { $log.Add("[VM-Project-Error] $proj instances JSON parse failed: $($_.Exception.Message)") | Out-Null; $vmList=@() } }
            }
        } catch { $log.Add("[VM-Project-Error] $proj instances command threw: $($_.Exception.Message)") | Out-Null; $vmList=@() }
        if ($apiDisabled) {
            $log.Add("[VM-Project-Skip] $proj Compute API disabled - skipping VMs & disks") | Out-Null
            return [PSCustomObject]@{ Project=$proj; VMs=@(); AttachedDisks=@(); AllDisks=@(); UnattachedDisks=@(); Logs=$log; DurationSec=[math]::Round(([DateTime]::UtcNow - $startUtc).TotalSeconds,2) }
        }
        # Disks
        try {
            $diskRaw = & gcloud --quiet compute disks list --project $proj --format=json 2>&1
            if ($LASTEXITCODE -ne 0) {
                $dmsg = ($diskRaw | Out-String).Trim()
                $log.Add("[VM-Project-Warn] $proj disks list failed exit=$LASTEXITCODE msg=$dmsg") | Out-Null
                $diskListAll = @()
            } else {
                if ([string]::IsNullOrWhiteSpace(($diskRaw|Out-String))) { $diskListAll=@() } else { try { $diskListAll = $diskRaw | ConvertFrom-Json } catch { $log.Add("[VM-Project-Error] $proj disks JSON parse failed: $($_.Exception.Message)") | Out-Null; $diskListAll=@() } }
            }
        } catch { $log.Add("[VM-Project-Error] $proj disks command threw: $($_.Exception.Message)") | Out-Null; $diskListAll=@() }
        if (-not $vmList) { $vmList=@() }; if (-not $diskListAll) { $diskListAll=@() }
        $diskMap = @{}; foreach ($d in $diskListAll) { if ($d.name) { $diskMap[$d.name]=$d } }
        $projectVMs=@(); $projectAttached=@(); $projectAllDisks=@(); $projectUnattached=@()
        $vmIndex=0
        foreach ($vm in $vmList) {
            $vmIndex++
            $zone = ($vm.zone -replace '.*/','')
            $region = Get-RegionFromZoneInner -zone $vm.zone
            $osType='Linux'
            try {
                if ($vm.disks) {
                    foreach ($vd in $vm.disks) {
                        if ($vd.licenses) { foreach ($lic in $vd.licenses) { if ($lic -match 'windows') { $osType='Windows'; break } }; if ($osType -eq 'Windows') { break } }
                    }
                }
                if ($osType -ne 'Windows' -and $vm.labels) { $lbl = $vm.labels.PSObject.Properties.Name | Where-Object { $_ -match 'windows' }; if ($lbl) { $osType='Windows' } }
            } catch { $osType='Linux' }
            $vmDiskGB=0
            if ($vm.disks) {
                foreach ($disk in $vm.disks) {
                    $diskName = ($disk.source -split '/')[-1]
                    if ($diskName -and $diskMap.ContainsKey($diskName)) {
                        $d=$diskMap[$diskName]; $vmDiskGB += [int64]$d.sizeGb
                        $isRegional = ($null -ne $d.region)
                        $projectAttached += [PSCustomObject]@{
                            DiskName=$d.name; VMName=$vm.name; Project=$proj; Region= if ($d.region){($d.region -split '/')[-1]} else {(Get-RegionFromZoneInner -zone $d.zone)}; Zone= if ($d.region){''} else {($d.zone -split '/')[-1]}; IsRegional=[bool]$isRegional; Encrypted= if ($d.diskEncryptionKey -or $d.encryptionKey){'Yes'} else {'No'}; DiskType=($d.type -replace '.*/',''); SizeGB=[int64]$d.sizeGb }
                    }
                }
            }
            $diskCountLocal = if ($vm.disks){$vm.disks.Count}else{0}
            $log.Add( "[VM] Project=$proj $vmIndex/$($vmList.Count) Name=$($vm.name) Type=$(($vm.machineType -replace '.*/','')) Region=$region Zone=$zone Disks=$diskCountLocal DiskGB=$vmDiskGB" ) | Out-Null
            $projectVMs += [PSCustomObject]@{ Project=$proj; VMName=$vm.name; VMSize=($vm.machineType -replace '.*/',''); OS=$osType; Region=$region; Zone=$zone; VMId=$vm.id; DiskCount=$diskCountLocal; VMDiskSizeGB=[int64]$vmDiskGB }
        }
        foreach ($disk in $diskListAll) {
            $isRegional = ($null -ne $disk.region)
            $diskObj = [PSCustomObject]@{ DiskName=$disk.name; VMName= if ($disk.users -and $disk.users.Count -gt 0){ ($disk.users | ForEach-Object { ($_ -split '/')[-1] }) -join ',' } else { $null }; Project=$proj; Region= if ($disk.region){($disk.region -split '/')[-1]} else {(Get-RegionFromZoneInner -zone $disk.zone)}; Zone= if ($disk.region){''} else {($disk.zone -split '/')[-1]}; IsRegional=[bool]$isRegional; Encrypted= if ($disk.diskEncryptionKey -or $disk.encryptionKey){'Yes'} else {'No'}; DiskType=($disk.type -replace '.*/',''); SizeGB=[int64]$disk.sizeGb }
            $projectAllDisks += $diskObj; if (-not $disk.users -or $disk.users.Count -eq 0){ $projectUnattached += $diskObj }
        }
        $log.Add("[VM-Project-End] $proj VMs=$($projectVMs.Count) Disks=$($projectAllDisks.Count)") | Out-Null
        $durationSec = [math]::Round(([DateTime]::UtcNow - $startUtc).TotalSeconds,2)
        return [PSCustomObject]@{ Project=$proj; VMs=$projectVMs; AttachedDisks=$projectAttached; AllDisks=$projectAllDisks; UnattachedDisks=$projectUnattached; Logs=$log; DurationSec=$durationSec }
    }

    $maxThreads = [Math]::Min(10,[Math]::Max(1,$ProjectIds.Count))
    $iss = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()
    $pool = [RunspaceFactory]::CreateRunspacePool(1,$maxThreads,$iss,$Host); $pool.Open()
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
                }
            } else { $still += $rs }
        }
        $runspaces = $still
        # Update progress
        $pct = if ($total -gt 0) { [math]::Round(($completed / $total)*100,1) } else { 100 }
        $elapsedMin = [math]::Round(([DateTime]::UtcNow - $overallStart).TotalMinutes,3)
        $avgMin = if ($durations.Count -gt 0) { [math]::Round(($durations | Measure-Object -Average | Select -Expand Average),3) } else { 0 }
        $remaining = $total - $completed; $etaMin = if ($avgMin -gt 0 -and $remaining -gt 0) { [math]::Round($avgMin * $remaining,2) } else { 0 }
        $rate = if ($elapsedMin -gt 0) { [math]::Round($allVMs.Count / ($elapsedMin*60),2) } else { 0 }
        $status = "Projects {0}/{1} ({2}%) | CumVMs={3} | Rate={4}/s | ElapsedMin={5} | ETA_Min={6}" -f $completed,$total,$pct,$allVMs.Count,$rate,$elapsedMin,$etaMin
        Write-Progress -Id 101 -Activity 'VM Projects' -Status $status -PercentComplete $pct
        if ($runspaces.Count -gt 0) { Start-Sleep -Milliseconds 200 }
    }
    Write-Progress -Id 101 -Activity 'VM Projects' -Completed
    $pool.Close(); $pool.Dispose()
    return @{ VMs=$allVMs; AttachedDisks=$allAttached; AllDisks=$allAllDisks; UnattachedDisks=$allUnattached }
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
                $sizes = gcloud storage objects list "gs://$BucketName" --project $Project --format="value(size)" 2>$null
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

## Removed old nested project/bucket runspace approach; replaced by two-phase global model


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
            $raw = & gcloud storage buckets list --project $project --format=json 2>&1
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
        $projectStatuses += [PSCustomObject]@{ Project=$display; BucketCount=$bucketCountCsv; PermissionIssue= if ($pr.PermissionIssue){'Y'} else {''} }
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
        $totalBytes = ($sized | Measure-Object UsedCapacityBytes -Sum).Sum
        $totalGB = if ($totalBytes) { [math]::Round($totalBytes/1e9,3) } else { 0 }
        Write-Progress -Id 401 -Activity 'Bucket Sizing' -Status ("Buckets {0}/{1} ({2}%) SizedGB={3} ElapsedSec={4}" -f $done,$totalBuckets,$pct,$totalGB,$elapsed) -PercentComplete $pct
        if ($bucketRunspaces.Count -gt 0) { Start-Sleep -Milliseconds 200 }
    }
    Write-Progress -Id 401 -Activity 'Bucket Sizing' -Completed
    $pool2.Close(); $pool2.Dispose()

    $script:StorageProjectStatuses = $projectStatuses
    return $sized
}


# -------------------------
# Execution Flow
# -------------------------
$allProjects = Get-GcpProjects
if ($Projects) {
    $targetProjects = $Projects | Where-Object { $allProjects -contains $_ }
    if (-not $targetProjects) {
        Write-Error "No valid projects found from provided list."
        Stop-Transcript | Out-Null
        exit 1
    }
} else {
    $targetProjects = $allProjects
}
Write-Host "Targeting $($targetProjects.Count) projects." -ForegroundColor Green

$invResults = @{}
if ($Selected.VM)      { $invResults = Get-GcpVMInventory -ProjectIds $targetProjects }
if ($Selected.STORAGE) { $invResults.StorageBuckets = Get-GcpStorageInventory -ProjectIds $targetProjects }

# -------------------------
# Output CSVs
# -------------------------
Write-Progress -Id 5 -Activity "Generating Output Files" -Status "Exporting CSV files..." -PercentComplete 0

# Plain CSV writer (no double quotes). Assumes field values do not contain commas.
function Write-PlainCsv {
    param(
        [Parameter(Mandatory)]$Data,
        [Parameter(Mandatory)][string]$Path
    )
    if (-not $Data -or ($Data | Measure-Object).Count -eq 0) { return }
    $first = $Data | Select-Object -First 1
    $cols = $first.PSObject.Properties | Where-Object { $_.MemberType -eq 'NoteProperty' } | Select-Object -ExpandProperty Name
    Set-Content -Path $Path -Value ($cols -join ',')
    foreach ($row in $Data) {
        $values = foreach ($c in $cols) {
            $v = $row.$c
            if ($null -eq $v) { '' } else { ($v.ToString() -replace '"','') -replace "[\r\n]",' ' }
        }
        Add-Content -Path $Path -Value ($values -join ',')
    }
}

# -------------------------
# Summary append helpers (appends to existing CSV files with 4 blank lines then a section)
# -------------------------
function Add-BlankLines([string]$Path,[int]$Count=4){ 1..$Count | ForEach-Object { Add-Content -Path $Path -Value '' } }

function Add-VmInfoSummary {
    param([string]$Path,[object[]]$VmData)
    if (-not $VmData -or $VmData.Count -eq 0) { return }
    Add-BlankLines -Path $Path -Count 4
    Add-Content -Path $Path -Value '### VM Info Summary ###'
    Add-Content -Path $Path -Value ("Total VMs, {0}" -f $VmData.Count)
    # Per Project
    $projGroups = $VmData | Group-Object Project | Sort-Object Name
    Add-Content -Path $Path -Value 'Per Project:'
    foreach ($pg in $projGroups) { Add-Content -Path $Path -Value ("Project, {0}, VMs, {1}" -f $pg.Name,$pg.Count) }
    # Per Project -> Region
    Add-Content -Path $Path -Value 'Per Project Region:'
    foreach ($pg in $projGroups) {
        $regionGroups = $pg.Group | Group-Object Region | Sort-Object Name
        foreach ($rg in $regionGroups) { Add-Content -Path $Path -Value ("ProjectRegion, {0}, {1}, VMs, {2}" -f $pg.Name,$rg.Name,$rg.Count) }
    }
    # Per Project -> Zone
    Add-Content -Path $Path -Value 'Per Project Zone:'
    foreach ($pg in $projGroups) {
        $zoneGroups = $pg.Group | Group-Object Zone | Sort-Object Name
        foreach ($zg in $zoneGroups) { Add-Content -Path $Path -Value ("ProjectZone, {0}, {1}, {2}, VMs, {3}" -f $pg.Name,($zg.Group | Select-Object -First 1).Region,$zg.Name,$zg.Count) }
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
    Write-PlainCsv -Data $invResults.VMs -Path $vmCsv
    Write-Host "VMs CSV written: $(Split-Path $vmCsv -Leaf)" -ForegroundColor Cyan
    Add-VmInfoSummary -Path $vmCsv -VmData $invResults.VMs

    if ($invResults.AttachedDisks -and $invResults.AttachedDisks.Count) {
    $attachedCsv = Join-Path $outDir ("gcp_disks_attached_to_vm_instances_" + $dateStr + ".csv")
        Write-PlainCsv -Data $invResults.AttachedDisks -Path $attachedCsv
        Write-Host "Attached disks CSV written: $(Split-Path $attachedCsv -Leaf)" -ForegroundColor Cyan
    Add-DiskSummary -Path $attachedCsv -DiskData $invResults.AttachedDisks -Title 'Attached Disks'
    }

    if ($invResults.UnattachedDisks -and $invResults.UnattachedDisks.Count) {
    $unattachedCsv = Join-Path $outDir ("gcp_disks_unattached_to_vm_instances_" + $dateStr + ".csv")
        Write-PlainCsv -Data $invResults.UnattachedDisks -Path $unattachedCsv
        Write-Host "Unattached disks CSV written: $(Split-Path $unattachedCsv -Leaf)" -ForegroundColor Cyan
    Add-DiskSummary -Path $unattachedCsv -DiskData $invResults.UnattachedDisks -Title 'Unattached Disks'
    }
}

if ($Selected.STORAGE -and $invResults.StorageBuckets -and $invResults.StorageBuckets.Count) {
    $bktCsv = Join-Path $outDir ("gcp_storage_buckets_info_" + $dateStr + ".csv")
    Write-PlainCsv -Data $invResults.StorageBuckets -Path $bktCsv
    Write-Host "Buckets CSV written: $(Split-Path $bktCsv -Leaf)" -ForegroundColor Cyan
    Add-BucketSummary -Path $bktCsv -Buckets $invResults.StorageBuckets
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
