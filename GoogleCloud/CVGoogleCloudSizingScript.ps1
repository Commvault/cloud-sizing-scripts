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

# -------------------------
# Setup output + transcript
# -------------------------
$dateStr = (Get-Date).ToString("yyyy-MM-dd_HHmmss")
$outDir = Join-Path -Path $PWD -ChildPath ("gcp-inv-" + $dateStr)
New-Item -Path $outDir -ItemType Directory -Force | Out-Null

$logFile = Join-Path $outDir ("gcp_sizing_script_output_" + $dateStr + ".log")
Start-Transcript -Path $logFile -Append | Out-Null

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
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet('INFO','WARN','ERROR','DEBUG')][string]$Level = 'INFO'
    )
    $ts = (Get-Date).ToString('s')
    # Host output is captured by Start-Transcript already
    switch ($Level) {
        'ERROR' { Write-Host "[$ts] [$Level] $Message" -ForegroundColor Red }
        'WARN'  { Write-Host "[$ts] [$Level] $Message" -ForegroundColor Yellow }
        'DEBUG' { Write-Host "[$ts] [$Level] $Message" -ForegroundColor DarkGray }
        default { Write-Host "[$ts] [$Level] $Message" -ForegroundColor Gray }
    }
}

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

# -------------------------
# VM + Disk Inventory (fast)
# -------------------------
function Get-GcpVMInventory {
    param([string[]]$ProjectIds)

    $VMs = @()
    $AttachedDisks = @()
    $AllDisks = @()
    $UnattachedDisks = @()

    $projIndex = 0
    foreach ($proj in $ProjectIds) {
        $projIndex++
        $projPercent = if ($ProjectIds.Count) { [math]::Round(($projIndex / $ProjectIds.Count) * 100,1) } else { 100 }
        Write-Progress -Id 1 -Activity "Processing GCP VM workload" `
            -Status "Project $projIndex/$($ProjectIds.Count) ($projPercent%): $proj" `
            -PercentComplete $projPercent

        Write-Host "`n=== Project: $proj ===" -ForegroundColor Yellow

        # Get all instances and all disks ONCE per project
        $vmList = @()
        $diskListAll = @()
        try {
            $vmList = gcloud compute instances list --project $proj --format=json | ConvertFrom-Json
        } catch { Write-Warning "Failed to list instances in ${proj}: $_" }

        try {
            $diskListAll = gcloud compute disks list --project $proj --format=json | ConvertFrom-Json
        } catch { Write-Warning "Failed to list disks in ${proj}: $_" }

        if (-not $vmList)   { $vmList = @() }
        if (-not $diskListAll) { $diskListAll = @() }

        # Build disk map for O(1) lookups
        $diskMap = @{}
        foreach ($d in $diskListAll) { $diskMap[$d.name] = $d }

        # Track attached disk names per project
        $attachedDiskNames = New-Object System.Collections.Generic.HashSet[string]

    # VM loop (no per-disk describe calls)
    $vmCount = 0
    foreach ($vm in $vmList) {
            $vmCount++
            $vmPercent = if ($vmList.Count) { [math]::Round(($vmCount / $vmList.Count) * 100, 1) } else { 100 }
            Write-Progress -Id 2 -ParentId 1 -Activity "Processing VMs" `
        -Status "VM $vmCount/$($vmList.Count) ($vmPercent%): $($vm.name)" `
        -PercentComplete $vmPercent

            # Inner per-VM workload progress (3 steps) - reset per VM
            $vmWorkTotal = 3
            $vmWorkStep = 0
            function Update-VMWorkProgress {
                param([string]$Phase,[int]$Current,[int]$Total)
                if ($Total -le 0) { $Total = 1 }
                $pctRaw = ($Current / $Total) * 100
                if ($pctRaw -gt 100) { $pctRaw = 100 }
                $pct = [math]::Round($pctRaw,0)
                Write-Progress -Id 21 -ParentId 2 -Activity "VM Workload" -Status ("{0} ({1}/{2})" -f $Phase,$Current,$Total) -PercentComplete $pct
            }
            $vmWorkStep++; Update-VMWorkProgress -Phase "Detecting OS" -Current $vmWorkStep -Total $vmWorkTotal

            $zone = ($vm.zone -replace '.*/','')
            $region = Get-RegionFromZone -zone $vm.zone

            # OS detection (improved): check disk licenses first (fast, in list output), then labels.
            $osType = "Linux"
            try {
                if ($vm.disks) {
                    foreach ($vd in $vm.disks) {
                        if ($vd.licenses) {
                            foreach ($lic in $vd.licenses) {
                                if ($lic -match 'windows') { $osType = 'Windows'; break }
                            }
                        }
                        if ($osType -eq 'Windows') { break }
                    }
                }
                if ($osType -ne 'Windows' -and $vm.labels) {
                    $lbl = $vm.labels.PSObject.Properties.Name | Where-Object { $_ -match 'windows' }
                    if ($lbl) { $osType = 'Windows' }
                }
            } catch { $osType = 'Linux' }

            # Sum attached disk sizes via diskMap
            $vmWorkStep++; Update-VMWorkProgress -Phase "Aggregating Disks" -Current $vmWorkStep -Total $vmWorkTotal
            $vmDiskGB = 0
            if ($vm.disks) {
                foreach ($disk in $vm.disks) {
                    $diskName = ($disk.source -split '/')[-1]
                    if ($diskName) { $null = $attachedDiskNames.Add($diskName) }
                    $d = $null
                    if ($diskMap.ContainsKey($diskName)) { $d = $diskMap[$diskName] }
                    if ($d) {
                        $vmDiskGB += [int64]$d.sizeGb
                        $isRegional = ($null -ne $d.region)
                        $AttachedDisks += [PSCustomObject]@{
                            DiskName   = $d.name
                            VMName     = $vm.name
                            Project    = $proj
                            Region     = if ($d.region) { ($d.region -split '/')[-1] } else { (Get-RegionFromZone -zone $d.zone) }
                            Zone       = if ($d.region) { "" } else { ($d.zone -split '/')[-1] }
                            IsRegional = [bool]$isRegional
                            Encrypted  = if ($d.diskEncryptionKey -or $d.encryptionKey) { 'Yes' } else { 'No' }
                            DiskType   = ($d.type -replace '.*/','')
                            SizeGB     = [int64]$d.sizeGb
                        }
                    }
                }
            }

            $diskCount = if ($vm.disks) { $vm.disks.Count } else { 0 }

            # Per-VM log line for transcript (fast, no extra gcloud calls)
            Write-Host ("[VM] {0} | {1}/{2} ({3}%) | Name={4} | Type={5} | Region={6} | Zone={7} | Disks={8} | DiskGB={9}" -f $proj, $vmCount, $vmList.Count, $vmPercent, $vm.name, ($vm.machineType -replace '.*/',''), $region, $zone, $diskCount, $vmDiskGB) -ForegroundColor DarkCyan

            $vmWorkStep++; Update-VMWorkProgress -Phase "Recording VM" -Current $vmWorkStep -Total $vmWorkTotal
            $VMs += [PSCustomObject]@{
                Project      = $proj
                VMName       = $vm.name
                VMSize       = ($vm.machineType -replace '.*/','')
                OS           = $osType
                Region       = $region
                Zone         = $zone
                VMId         = $vm.id
                DiskCount    = $diskCount
                VMDiskSizeGB = [int64]$vmDiskGB
            }
            # Complete inner progress for this VM
            Write-Progress -Id 21 -Activity "VM Workload" -Completed
        }

        Write-Progress -Id 2 -Activity "Processing VMs" -Completed

        # Build AllDisks / UnattachedDisks (fast using 'users' field)
        foreach ($disk in $diskListAll) {
            $isRegional = ($null -ne $disk.region)
            $allDiskObj = [PSCustomObject]@{
                DiskName   = $disk.name
                VMName     = if ($disk.users -and $disk.users.Count -gt 0) { ($disk.users | ForEach-Object { ($_ -split '/')[-1] }) -join ',' } else { $null }
                Project    = $proj
                Region     = if ($disk.region) { ($disk.region -split '/')[-1] } else { (Get-RegionFromZone -zone $disk.zone) }
                Zone       = if ($disk.region) { "" } else { ($disk.zone -split '/')[-1] }
                IsRegional = [bool]$isRegional
                Encrypted  = if ($disk.diskEncryptionKey -or $disk.encryptionKey) { 'Yes' } else { 'No' }
                DiskType   = ($disk.type -replace '.*/','')
                SizeGB     = [int64]$disk.sizeGb
            }
            $AllDisks += $allDiskObj
            if (-not $disk.users -or $disk.users.Count -eq 0) {
                $UnattachedDisks += $allDiskObj
            }
        }
    }

    Write-Progress -Id 1 -Activity "Processing GCP VM workload" -Completed

    return @{
        VMs             = $VMs
        AttachedDisks   = $AttachedDisks
        AllDisks        = $AllDisks
        UnattachedDisks = $UnattachedDisks
    }
}

# -------------------------
# Storage Inventory (fast, gcloud-only)
# -------------------------
function Get-GcpStorageInventory {
    param([string[]]$ProjectIds)

    $StorageBuckets = @()
    $projIndex = 0

    foreach ($proj in $ProjectIds) {
        $projIndex++
        Write-Progress -Id 3 -Activity "Processing Storage" `
            -Status "Project $proj ($projIndex/$($ProjectIds.Count))" `
            -PercentComplete ([math]::Round(($projIndex / $ProjectIds.Count) * 100, 1))

        Write-Host "`n=== Storage Project: $proj ===" -ForegroundColor Yellow

        $bucketList = @()
        try {
            $bucketList = gcloud storage buckets list --project $proj --format=json | ConvertFrom-Json
        } catch {
            Write-Warning "Failed to list buckets in ${proj}: $_"
            $bucketList = @()
        }

        $bCount = 0
        $bTotal = ($bucketList | Measure-Object).Count

    foreach ($bucket in $bucketList) {
            $bCount++
            $percent = if ($bTotal) { [math]::Round(($bCount / $bTotal) * 100, 1) } else { 100 }
            Write-Progress -Id 4 -ParentId 3 -Activity "Sizing Buckets" `
                -Status "Bucket $bCount of ${bTotal}: $($bucket.name)" `
                -PercentComplete $percent

            # Inner per-bucket workload progress (3 steps) - reset per bucket
            $bucketWorkTotal = 3
            $bucketWorkStep = 0
            function Update-BucketWorkProgress {
                param([string]$Phase,[int]$Current,[int]$Total)
                if ($Total -le 0) { $Total = 1 }
                $pctRaw = ($Current / $Total) * 100
                if ($pctRaw -gt 100) { $pctRaw = 100 }
                $pct = [math]::Round($pctRaw,0)
                Write-Progress -Id 41 -ParentId 4 -Activity "Bucket Workload" -Status ("{0} ({1}/{2})" -f $Phase,$Current,$Total) -PercentComplete $pct
            }
            $bucketWorkStep++; Update-BucketWorkProgress -Phase "Measuring Size" -Current $bucketWorkStep -Total $bucketWorkTotal

            $sizeBytes = Get-BucketSizeBytes -BucketName $bucket.name -Project $proj

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

            $bucketWorkStep++; Update-BucketWorkProgress -Phase "Recording Bucket" -Current $bucketWorkStep -Total $bucketWorkTotal
            $StorageBuckets += [PSCustomObject]@{
                StorageBucket       = $bucket.name
                Project             = $proj
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

            Write-Host ("[BUCKET] {0} | {1}/{2} ({3}%) | Name={4} | Loc={5} | Class={6} | Bytes={7} | GiB={8} | GB(dec)={9}" -f $proj,$bCount,$bTotal,$percent,$bucket.name,$bucket.location,$bucket.storageClass,$bytes,$sizeGiB,$sizeGBDec) -ForegroundColor DarkGreen
            Write-Log "Collected bucket: $($bucket.name) in ${proj}, bytes: $bytes, sizeGiB: $sizeGiB, sizeGB(dec): $sizeGBDec"
            $bucketWorkStep++; Update-BucketWorkProgress -Phase "Completed" -Current $bucketWorkStep -Total $bucketWorkTotal
            Write-Progress -Id 41 -Activity "Bucket Workload" -Completed
        }

        Write-Progress -Id 4 -Activity "Sizing Buckets" -Completed
    }

    Write-Progress -Id 3 -Activity "Processing Storage" -Completed
    return $StorageBuckets
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
    $vmCsv = Join-Path $outDir ("gcp_vm_info_" + $dateStr + ".csv")
    Write-PlainCsv -Data $invResults.VMs -Path $vmCsv
    Write-Host "VMs CSV written: $(Split-Path $vmCsv -Leaf)" -ForegroundColor Cyan
    Add-VmInfoSummary -Path $vmCsv -VmData $invResults.VMs

    if ($invResults.AttachedDisks -and $invResults.AttachedDisks.Count) {
        $attachedCsv = Join-Path $outDir ("gcp_disks_attached_to_vms_" + $dateStr + ".csv")
        Write-PlainCsv -Data $invResults.AttachedDisks -Path $attachedCsv
        Write-Host "Attached disks CSV written: $(Split-Path $attachedCsv -Leaf)" -ForegroundColor Cyan
    Add-DiskSummary -Path $attachedCsv -DiskData $invResults.AttachedDisks -Title 'Attached Disks'
    }

    if ($invResults.UnattachedDisks -and $invResults.UnattachedDisks.Count) {
        $unattachedCsv = Join-Path $outDir ("gcp_disks_unattached_to_vms_" + $dateStr + ".csv")
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
    $summaryRows += [PSCustomObject]@{Level='Overall';ResourceType='AllResources';Project='All';Region='All';Zone='All';Count=($vmData.Count + $bucketData.Count);TotalSizeGB=$combinedGB;TotalSizeTB=[math]::Round($combinedGB/1e3,4);TotalSizeTiB=[math]::Round($combinedGB/1024,4)}
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
Stop-Transcript | Out-Null   # ensure log is fully flushed before zipping

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
