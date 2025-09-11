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
    # NOTE: Manual validation performed after normalization to allow inputs like -Types "VM,Storage"
    [string[]]$Types,
    [string[]]$Projects
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
    $allowed = @('VM','STORAGE')
    $bad = $Types | Where-Object { $allowed -notcontains ($_.Trim().ToUpper()) }
    if ($bad.Count -gt 0) {
        Write-Error ("Invalid value(s) for -Types: {0}. Valid values: VM, Storage" -f ($bad -join ', '))
        return
    }
}

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
    'FILESHARE' = 'FileShares'
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
        function Get-RegionFromZoneInner { param([string]$zone) if (-not $zone) { return 'Unknown' }; $z = $zone -replace '.*/',''; return ($z -replace '-[a-z]$','') }
        $apiDisabled = $false
        $permIssue   = $false
        $env:CLOUDSDK_CORE_DISABLE_PROMPTS = '1'

        # Instances
        try {
            $vmRaw = & gcloud --quiet compute instances list --project $proj --format=json 2>&1
            if ($LASTEXITCODE -ne 0) {
                $msg = ($vmRaw | Out-String).Trim()
                if     ($msg -match '(?i)not enabled|has not been used|is disabled|API .* not enabled') { $apiDisabled = $true }
                elseif ($msg -match '(?i)permission|denied|forbidden|403|PERMISSION_DENIED|insufficientPermissions') { $permIssue  = $true }
                $log.Add("[VM-Project-Warn] $proj instances list failed exit=$LASTEXITCODE msg=$msg") | Out-Null
                $vmList = @()
            } else {
                if ([string]::IsNullOrWhiteSpace(($vmRaw | Out-String))) { $vmList = @() }
                else { try { $vmList = $vmRaw | ConvertFrom-Json } catch { $log.Add("[VM-Project-Error] $proj instances JSON parse failed: $($_.Exception.Message)") | Out-Null; $vmList=@() } }
            }
        } catch {
            $log.Add("[VM-Project-Error] $proj instances command threw: $($_.Exception.Message)") | Out-Null
            $vmList=@()
        }
        if ($apiDisabled) {
            $log.Add("[VM-Project-Skip] $proj Compute API disabled - skipping VMs & disks") | Out-Null
            return [PSCustomObject]@{ Project=$proj; VMs=@(); AttachedDisks=@(); AllDisks=@(); UnattachedDisks=@(); Logs=$log; DurationSec=[math]::Round(([DateTime]::UtcNow - $startUtc).TotalSeconds,2) }
        }
        if ($permIssue) {
            $log.Add("[VM-Project-Skip] $proj Compute API permission issue - skipping VMs & disks") | Out-Null
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
                if ([string]::IsNullOrWhiteSpace(($diskRaw | Out-String))) { $diskListAll=@() }
                else { try { $diskListAll = $diskRaw | ConvertFrom-Json } catch { $log.Add("[VM-Project-Error] $proj disks JSON parse failed: $($_.Exception.Message)") | Out-Null; $diskListAll=@() } }
            }
        } catch {
            $log.Add("[VM-Project-Error] $proj disks command threw: $($_.Exception.Message)") | Out-Null
            $diskListAll=@()
        }

        if (-not $vmList)      { $vmList      = @() }
        if (-not $diskListAll) { $diskListAll = @() }

        # Build disk lookup map. Original implementation keyed only by disk name which can collide across zones/regions.
        # We now key by (a) selfLink (preferred), and (b) composite "<zoneOrRegion>|<name>". We still keep first-seen entries for composite keys.
        $diskMap = @{}
        foreach ($d in $diskListAll) {
            if (-not $d.name) { continue }
            $selfKey = $null
            if ($d.PSObject.Properties.Name -contains 'selfLink' -and $d.selfLink) { $selfKey = $d.selfLink.ToLower() }
            elseif ($d.PSObject.Properties.Name -contains 'id' -and $d.id) { $selfKey = ("id://" + $d.id) }
            else {
                # Fallback synthetic key
                $zr = if ($d.region) { ($d.region -split '/')[-1] } elseif ($d.zone) { ($d.zone -split '/')[-1] } else { '' }
                $selfKey = ("//disk/" + $zr + "/" + $d.name).ToLower()
            }
            $diskMap[$selfKey] = $d
            $zr2 = if ($d.region) { ($d.region -split '/')[-1] } elseif ($d.zone) { ($d.zone -split '/')[-1] } else { '' }
            $composite = ($zr2 + '|' + $d.name).ToLower()
            if (-not $diskMap.ContainsKey($composite)) { $diskMap[$composite] = $d }
        }

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
            try {
                if ($vm.disks) {
                    foreach ($vd in $vm.disks) {
                        if ($vd.licenses) {
                            foreach ($lic in $vd.licenses) { if ($lic -match 'windows') { $osType='Windows'; break } }
                            if ($osType -eq 'Windows') { break }
                        }
                    }
                }
                if ($osType -ne 'Windows' -and $vm.labels) {
                    $lbl = $vm.labels.PSObject.Properties.Name | Where-Object { $_ -match 'windows' }
                    if ($lbl) { $osType='Windows' }
                }
            } catch { $osType='Linux' }

            $vmDiskGB = 0
            if ($vm.disks) {
                if (-not $script:__AllDiskKeys) { $script:__AllDiskKeys = @{} }
                foreach ($disk in $vm.disks) {
                    $diskName = ($disk.source -split '/')[-1]
                    $primaryKey = if ($disk.source) { $disk.source.ToLower() } else { $null }
                    $d = $null
                    $fromMap = $false
                    if ($primaryKey -and $diskMap.ContainsKey($primaryKey)) { $d = $diskMap[$primaryKey]; $fromMap=$true }
                    else {
                        # Derive composite key from URL for lookup
                        $zrMatch = ''
                        if ($disk.source -match '/zones/([^/]+)/') { $zrMatch = $Matches[1] }
                        elseif ($disk.source -match '/regions/([^/]+)/') { $zrMatch = $Matches[1] }
                        $altKey = ($zrMatch + '|' + $diskName).ToLower()
                        if ($diskMap.ContainsKey($altKey)) { $d = $diskMap[$altKey]; $fromMap=$true }
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
                        if ($d.region) { $regionVal = ($d.region -split '/')[-1]; $isRegional = $true } else { $zoneVal = ($d.zone -split '/')[-1]; $regionVal = Get-RegionFromZoneInner -zone $d.zone }
                        if ($d.diskEncryptionKey -or $d.encryptionKey) { $enc='Yes' }
                        $typeVal = ($d.type -replace '.*/','')
                        if ($d.PSObject.Properties.Name -contains 'selfLink') { $selfLinkVal = $d.selfLink }
                    } else {
                        # Use attachment data only
                        if ($disk.PSObject.Properties.Name -contains 'diskSizeGb' -and $disk.diskSizeGb) { $sizeGbVal = [int64]$disk.diskSizeGb }
                        if ($disk.source -match '/zones/([^/]+)/') { $zoneVal = $Matches[1]; $regionVal = Get-RegionFromZoneInner -zone $zoneVal }
                        elseif ($disk.source -match '/regions/([^/]+)/') { $regionVal = $Matches[1]; $isRegional = $true }
                        if ($disk.PSObject.Properties.Name -contains 'type' -and $disk.type) { $typeVal = ($disk.type -replace '.*/','') }
                        $selfLinkVal = $disk.source
                    }
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
                        Project=$proj; InstanceName=''; ShareName='(error)'; Tier=''; Region=''; Zone=''; CapacityGB=0; Networks=''; IPAddresses=''; State='ERROR'; CreateTime=''; Labels=''; Protocol=''; Error=$msg
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
                    $capGB = 0; try { if ($share.capacityGb) { $capGB=[int64]$share.capacityGb } } catch {}
                    $log.Add("[FS] Project=$proj $idx/$($instances.Count) Instance=$instanceShort Share=$shareName Tier=$tier Region=$region Zone=$zone CapacityGB=$capGB")|Out-Null
                    $shares += [PSCustomObject]@{
                        Project=$proj
                        InstanceName=$instanceShort
                        ShareName=$shareName
                        Tier=$tier
                        Region=$region
                        Zone=$zone
                        CapacityGB=$capGB
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
                $log.Add("[FS] Project=$proj $idx/$($instances.Count) Instance=$instanceShort Tier=$tier Region=$region Zone=$zone CapacityGB=0 (NoShares)")|Out-Null
                $shares += [PSCustomObject]@{
                    Project=$proj
                    InstanceName=$instanceShort
                    ShareName=''
                    Tier=$tier
                    Region=$region
                    Zone=$zone
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
                                if ($line -match '^\[FS\] Project=([^ ]+) [^ ]+ Instance=([^ ]+) Share=([^ ]+) Tier=([^ ]+) Region=([^ ]+) Zone=([^ ]+) CapacityGB=([0-9]+)') {
                                    $recovered += [PSCustomObject]@{
                                        Project=$Matches[1]; InstanceName=$Matches[2]; ShareName=$Matches[3]; Tier=$Matches[4]; Region=$Matches[5]; Zone=$Matches[6]; CapacityGB=[int64]$Matches[7]; Networks=''; IPAddresses=''; State='READY'; CreateTime=''; Labels=''; Protocol='NFS'; Error=''
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
                                if ($line -match '^\[FS\] Project=([^ ]+) [^ ]+ Instance=([^ ]+) Share=([^ ]+) Tier=([^ ]+) Region=([^ ]+) Zone=([^ ]+) CapacityGB=([0-9]+)') {
                                    $recovered += [PSCustomObject]@{
                                        Project=$Matches[1]; InstanceName=$Matches[2]; ShareName=$Matches[3]; Tier=$Matches[4]; Region=$Matches[5]; Zone=$Matches[6]; CapacityGB=[int64]$Matches[7]; Networks=''; IPAddresses=''; State='READY'; CreateTime=''; Labels=''; Protocol='NFS'; Error=''
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

# Global Filestore reconstruction (if we somehow ended up with zero successful shares but logs contain entries)
if ($Selected.FILESHARE) {
    $fsShares = $invResults.FileShares
    $fsGood = @()
    if ($fsShares) { $fsGood = $fsShares | Where-Object { -not ($_.Error -and $_.Error.Trim()) -and ($_.ShareName -ne '(error)') -and ($_.State -ne 'ERROR') } }
    if ((-not $fsGood) -or $fsGood.Count -eq 0) {
        if ($script:FileShareLogLines -and $script:FileShareLogLines.Count -gt 0) {
            $recoveredGlobal=@()
            foreach ($line in $script:FileShareLogLines) {
                if ($line -match '^\[FS\] Project=([^ ]+) [^ ]+ Instance=([^ ]+) Share=([^ ]+) Tier=([^ ]+) Region=([^ ]+) Zone=([^ ]+) CapacityGB=([0-9]+)') {
                    $recoveredGlobal += [PSCustomObject]@{
                        Project=$Matches[1]; InstanceName=$Matches[2]; ShareName=$Matches[3]; Tier=$Matches[4]; Region=$Matches[5]; Zone=$Matches[6]; CapacityGB=[int64]$Matches[7]; Networks=''; IPAddresses=''; State='READY'; CreateTime=''; Labels=''; Protocol='NFS'; Error=''
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
        $dummy = [PSCustomObject]@{ Project=''; InstanceName=''; ShareName=''; Tier=''; Region=''; Zone=''; CapacityGB=''; Networks=''; IPAddresses=''; State=''; CreateTime=''; Labels=''; Protocol=''; Error='' }
        Write-PlainCsv -Data @($dummy) -Path $fsCsv
        (Get-Content $fsCsv | Select-Object -First 1) | Set-Content $fsCsv
    }
    # Insert SuccessfulProjects summary (only if we actually had successes)
    if ($successCount -gt 0) {
        Add-Content -Path $fsCsv -Value ''
        Add-Content -Path $fsCsv -Value 'SuccessfulProjects:'
        Add-Content -Path $fsCsv -Value 'Project,Shares,TotalCapacityGB'
        foreach ($grp in ($success | Group-Object Project | Sort-Object Name)) {
            $cap = (($grp.Group | Measure-Object CapacityGB -Sum).Sum)
            Add-Content -Path $fsCsv -Value ("{0},{1},{2}" -f $grp.Name,$grp.Count,[int64]$cap)
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
    }
    Write-Host "Filestore CSV written: $(Split-Path $fsCsv -Leaf) (Good=$successCount Fail=$((@($fail)).Count))" -ForegroundColor Cyan
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
    $overallFsGB = ($fileshareData | Measure-Object CapacityGB -Sum).Sum
    if (-not $overallFsGB) { $overallFsGB = 0 }
    $summaryRows += [PSCustomObject]@{Level='Overall';ResourceType='Fileshare';Project='All';Region='All';Zone='All';Count=$fileshareData.Count;TotalSizeGB=[int64]$overallFsGB;TotalSizeTB=[math]::Round($overallFsGB/1e3,4);TotalSizeTiB=[math]::Round($overallFsGB/1024,4)}
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
        $projFsGB = ($projFS | Measure-Object CapacityGB -Sum).Sum
        $summaryRows += [PSCustomObject]@{Level='Project';ResourceType='Fileshare';Project=$proj;Region='All';Zone='All';Count=$projFS.Count;TotalSizeGB=[int64]$projFsGB;TotalSizeTB=[math]::Round($projFsGB/1e3,4);TotalSizeTiB=[math]::Round($projFsGB/1024,4)}
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
        $rGb = ($rg.Group | Measure-Object CapacityGB -Sum).Sum
        if (-not $rGb) { $rGb = 0 }
        $summaryRows += [PSCustomObject]@{Level='Region';ResourceType='Fileshare';Project='All';Region=$rg.Name;Zone='All';Count=$rg.Count;TotalSizeGB=[int64]$rGb;TotalSizeTB=[math]::Round($rGb/1e3,4);TotalSizeTiB=[math]::Round($rGb/1024,4)}
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
