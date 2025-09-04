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

function Parse-RegionFromZone([string]$zone) {
    if (-not $zone) { return "Unknown" }
    $z = $zone -replace '.*/',''
    return ($z -replace '-[a-z]$','')
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
        Write-Progress -Id 1 -Activity "Processing GCP Projects" `
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

            $zone = ($vm.zone -replace '.*/','')
            $region = Parse-RegionFromZone $vm.zone

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
                            Region     = if ($d.region) { ($d.region -split '/')[-1] } else { Parse-RegionFromZone $d.zone }
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
        }

        Write-Progress -Id 2 -Activity "Processing VMs" -Completed

        # Build AllDisks / UnattachedDisks (fast using 'users' field)
        foreach ($disk in $diskListAll) {
            $isRegional = ($null -ne $disk.region)
            $allDiskObj = [PSCustomObject]@{
                DiskName   = $disk.name
                VMName     = if ($disk.users -and $disk.users.Count -gt 0) { ($disk.users | ForEach-Object { ($_ -split '/')[-1] }) -join ',' } else { $null }
                Project    = $proj
                Region     = if ($disk.region) { ($disk.region -split '/')[-1] } else { Parse-RegionFromZone $disk.zone }
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

    Write-Progress -Id 1 -Activity "Processing GCP Projects" -Completed

    return @{
        VMs             = $VMs
        AttachedDisks   = $AttachedDisks
        AllDisks        = $AllDisks
        UnattachedDisks = $UnattachedDisks
    }
}

# -------------------------
# Storage Inventory (fast)
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

        $bucketNames = @()
        try {
            $bucketNames = gcloud storage buckets list --project $proj --format="value(name)"
        } catch {
            Write-Warning "Failed to list buckets in ${proj}: $_"
            $bucketNames = @()
        }

        $bCount = 0
        $bTotal = ($bucketNames | Measure-Object).Count
        foreach ($bucket in $bucketNames) {
            $bCount++
            $percent = if ($bTotal) { [math]::Round(($bCount / $bTotal) * 100, 1) } else { 100 }
            Write-Progress -Id 4 -ParentId 3 -Activity "Sizing Buckets" `
                -Status "Bucket $bCount of ${bTotal}: $bucket" `
                -PercentComplete $percent

            # Fast size: gsutil du -s returns "BYTES  gs://bucket"
            $sizeBytes = 0
            try {
                $du = gsutil du -s ("gs://{0}" -f $bucket) 2>$null
                if ($du) {
                    $parts = $du -split '\s+'
                    if ($parts.Length -gt 0) {
                        [int64]::TryParse($parts[0], [ref]$sizeBytes) | Out-Null
                    }
                }
            } catch {
                Write-Warning "Failed to size bucket $bucket in ${proj}: $_"
                $sizeBytes = 0
            }

            $StorageBuckets += [PSCustomObject]@{
                StorageBucket     = $bucket
                Project           = $proj
                UsedCapacityBytes = [int64]$sizeBytes
                UsedCapacityGiB   = [math]::Round($sizeBytes / 1GB, 0)
                UsedCapacityTiB   = [math]::Round(($sizeBytes / 1GB) / 1024, 4)
                UsedCapacityGB    = [math]::Round($sizeBytes / 1e9, 3)
                UsedCapacityTB    = [math]::Round($sizeBytes / 1e12, 4)
            }
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

# Always generate all VM-related CSVs if VM inventory is selected and data exists
if ($Selected.VM -and $invResults.VMs -and $invResults.VMs.Count) {
    $vmCsv = Join-Path $outDir ("gcp_vm_info_" + $dateStr + ".csv")
    $invResults.VMs | Export-Csv $vmCsv -NoTypeInformation
    Write-Host "VMs CSV written: $(Split-Path $vmCsv -Leaf)" -ForegroundColor Cyan

    if ($invResults.AttachedDisks -and $invResults.AttachedDisks.Count) {
        $attachedCsv = Join-Path $outDir ("gcp_disks_attached_to_vms_" + $dateStr + ".csv")
        $invResults.AttachedDisks | Export-Csv $attachedCsv -NoTypeInformation
        Write-Host "Attached disks CSV written: $(Split-Path $attachedCsv -Leaf)" -ForegroundColor Cyan
    }

    if ($invResults.UnattachedDisks -and $invResults.UnattachedDisks.Count) {
        $unattachedCsv = Join-Path $outDir ("gcp_disks_unattached_to_vms_" + $dateStr + ".csv")
        $invResults.UnattachedDisks | Export-Csv $unattachedCsv -NoTypeInformation
        Write-Host "Unattached disks CSV written: $(Split-Path $unattachedCsv -Leaf)" -ForegroundColor Cyan
    }
}

if ($Selected.STORAGE -and $invResults.StorageBuckets -and $invResults.StorageBuckets.Count) {
    $bktCsv = Join-Path $outDir ("gcp_storage_buckets_info_" + $dateStr + ".csv")
    $invResults.StorageBuckets | Export-Csv $bktCsv -NoTypeInformation
    Write-Host "Buckets CSV written: $(Split-Path $bktCsv -Leaf)" -ForegroundColor Cyan
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
    $summaryRows += [PSCustomObject]@{Level='Overall';ResourceType='VM';Project='All';Region='All';Zone='All';Count=$vmData.Count;TotalSizeGB=[int64]$overallDiskSizeGB;TotalSizeTB=[math]::Round($overallDiskSizeGB/1e3,4);TotalSizeTiB=[math]::Round($overallDiskSizeGB/1024,4)}
}
if ($Selected.STORAGE -and $bucketData) {
    $overallBucketBytes = ($bucketData | Measure-Object UsedCapacityBytes -Sum).Sum
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
$summaryRows | Export-Csv -Path $summaryCsv -NoTypeInformation
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
