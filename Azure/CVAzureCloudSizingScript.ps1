<#  
.SYNOPSIS  
    Azure Cloud Sizing Script - Comprehensive VM and Storage Account inventory and sizing analysis
.DESCRIPTION  
    Inventories Azure Virtual Machines and Storage Accounts across all or specified subscriptions.
    Calculates disk sizes for VMs and storage capacity utilization for Storage Accounts.
    Generates detailed CSV reports with comprehensive sizing information in multiple units (GB, TB, TiB).
    Includes hierarchical progress tracking and comprehensive logging.
    Outputs timestamped CSV files and creates a ZIP archive of all results.

.PARAMETER Types
    Optional. Restrict inventory to specific resource types.
    Valid values: VM, Storage
    If not specified, all supported resource types will be inventoried.
    
.PARAMETER Subscriptions
    Optional. Target specific subscriptions by name or ID.
    If not specified, all accessible subscriptions will be processed.
    
.EXAMPLE  
    .\CVAzureCloudSizingScript.ps1  
    # Inventories VMs and Storage Accounts in all accessible subscriptions  
.EXAMPLE  
    .\CVAzureCloudSizingScript.ps1 -Types VM,Storage  
    # Explicitly inventories VMs and Storage Accounts in all subscriptions (same as default)
.EXAMPLE  
    .\CVAzureCloudSizingScript.ps1 -Types VM
    # Only inventories Virtual Machines in all subscriptions
.EXAMPLE  
    .\CVAzureCloudSizingScript.ps1 -Subscriptions "Production","Development"  
    # Inventories VMs and Storage Accounts in only the Production and Development subscriptions
.EXAMPLE  
    .\CVAzureCloudSizingScript.ps1 -Types Storage -Subscriptions "Production"  
    # Only inventories Storage Accounts in the Production subscription  
    
.OUTPUTS
    Creates timestamped output directory with the following files:
    - azure_vm_info_YYYY-MM-DD_HHMMSS.csv - VM inventory with disk sizing
    - azure_storage_accounts_info_YYYY-MM-DD_HHMMSS.csv - Storage Account inventory with capacity metrics
    - azure_inventory_summary_YYYY-MM-DD_HHMMSS.csv - Comprehensive summary with regional breakdowns
    - azure_sizing_script_output_YYYY-MM-DD_HHMMSS.log - Complete execution log
    - azure_sizing_YYYY-MM-DD_HHMMSS.zip - ZIP archive containing all output files
    
.NOTES
    Requires Azure PowerShell modules: Az.Accounts, Az.Compute, Az.Storage, Az.Monitor, Az.Resources
    Script must be run by a user with appropriate Azure permissions to read VMs and Storage Accounts
    VM disk sizing includes both OS disks and data disks with error handling for inaccessible disks
    Storage Account metrics are retrieved from Azure Monitor for the last 24 hours

    SETUP INSTRUCTIONS FOR AZURE CLOUD SHELL (Recommended):

    1. Learn about Azure Cloud Shell:
       Visit: https://docs.microsoft.com/en-us/azure/cloud-shell/overview

    2. Verify Azure permissions:
       Ensure your Azure AD account has "Reader" role on target subscriptions
       Additional "Reader and Data Access" role may be needed for storage metrics

    3. Access Azure Cloud Shell:
       - Login to Azure Portal with verified account
       - Open Azure Cloud Shell (PowerShell mode)

    4. Upload this script:
       Use the Cloud Shell file upload feature to upload CVAzureCloudSizingScript.ps1

    5. Run the script:
       ./CVAzureCloudSizingScript.ps1
       ./CVAzureCloudSizingScript.ps1 -Types VM,Storage
       ./CVAzureCloudSizingScript.ps1 -Subscriptions "Production","Development"

    SETUP INSTRUCTIONS FOR LOCAL SYSTEM:

    1. Install PowerShell 7:
       Download from: https://github.com/PowerShell/PowerShell/releases

    2. Install required Azure PowerShell modules:
       Install-Module Az.Accounts,Az.Compute,Az.Storage,Az.Monitor,Az.Resources -Force

    3. Verify Azure permissions:
       Ensure your Azure AD account has "Reader" role on target subscriptions

    4. Connect to Azure:
       Connect-AzAccount

    5. Run the script:
       .\CVAzureCloudSizingScript.ps1
       .\CVAzureCloudSizingScript.ps1 -Types VM
       .\CVAzureCloudSizingScript.ps1 -Subscriptions "MySubscription"

    OUTPUT AND RESULTS:

    The script generates a timestamped ZIP file containing:
    - VM inventory CSV with disk sizing details
    - Storage Account inventory CSV with capacity metrics  
    - Summary CSV with regional breakdowns and totals
    - Complete execution log file

    All individual CSV files are automatically compressed into a single ZIP archive.
    The temporary directory is cleaned up after ZIP creation.
    Copy the generated ZIP file and share it with the requesting party.

    TROUBLESHOOTING:

    - Ensure proper Azure permissions before running
    - If module import fails, run: Import-Module Az -Force
    - For large environments, expect longer execution times
    - Check the log file in the ZIP for detailed error information
#>  
  
param(  
    [string[]]$Types, # Choices: VM, Storage  
    [string[]]$Subscriptions # Subscription names or IDs to target (if not specified, all subscriptions will be processed)
)  

# Set culture to en-US for consistent date and time formatting
$CurrentCulture = [System.Globalization.CultureInfo]::CurrentCulture
[System.Threading.Thread]::CurrentThread.CurrentCulture = 'en-US'
[System.Threading.Thread]::CurrentThread.CurrentUICulture = 'en-US'
  
# Resource type mapping  
$ResourceTypeMap = @{  
    "VM"         = "VMs"  
    "STORAGE"    = "StorageAccounts"
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
  
# Output Directory and Logging
$dateStr = Get-Date -Format "yyyy-MM-dd_HHmmss"  
$outdir = Join-Path -Path $PWD -ChildPath ("az-inv-" + $dateStr)  
New-Item -ItemType Directory -Force -Path $outdir | Out-Null

# Create comprehensive log file that captures everything
$logFile = Join-Path $outdir "azure_sizing_script_output_$dateStr.log"

# Start transcript to capture everything to log file
Start-Transcript -Path $logFile -Append

# Log script start and parameters
Write-Host "=== Azure Resource Inventory Started ===" -ForegroundColor Green
Write-Host "Script Parameters:" -ForegroundColor Green
if ($Types) { Write-Host "  Types: $($Types -join ', ')" -ForegroundColor Green }
if ($Subscriptions) { Write-Host "  Subscriptions: $($Subscriptions -join ', ')" -ForegroundColor Green }
if ($BlobLimit -gt 0) { Write-Host "  BlobLimit: $BlobLimit" -ForegroundColor Green }  
  
# Load modules  
$modules = @(  
    'Az.Accounts','Az.Compute','Az.Storage','Az.Monitor','Az.Resources'  
)  
foreach ($m in $modules) {  
    try { Import-Module $m -ErrorAction Stop } catch { Write-Warning "Could not load $m" }  
}  
  
# Subscription discovery  
$allSubs = Get-AzSubscription  
if ($allSubs -isnot [array]) { $allSubs = @($allSubs) }  

# Filter subscriptions if specified
if ($Subscriptions) {
    Write-Host "Filtering subscriptions based on provided list..." -ForegroundColor Yellow
    $subs = @()
    foreach ($subFilter in $Subscriptions) {
        $matchedSubs = $allSubs | Where-Object { 
            $_.Name -eq $subFilter -or 
            $_.Id -eq $subFilter -or 
            $_.SubscriptionId -eq $subFilter 
        }
        if ($matchedSubs) {
            $subs += $matchedSubs
        } else {
            Write-Warning "Subscription '$subFilter' not found or not accessible"
        }
    }
    
    if ($subs.Count -eq 0) {
        Write-Error "No valid subscriptions found from the provided list. Exiting."
        exit 1
    }
} else {
    Write-Host "No subscription filter specified, targeting all accessible subscriptions..." -ForegroundColor Yellow
    $subs = $allSubs
}

Write-Host "Targeting $($subs.Count) subscriptions: $($subs.Name -join ', ')" -ForegroundColor Green  

# Global output arrays for all resources
$VMs = @()  
$StorageAccounts = @()  

# Process each subscription sequentially
$subIdx = 0
foreach ($sub in $subs) {  
    $subIdx++
    Write-Progress -Id 1 -Activity "Processing Azure Subscriptions" -Status "Subscription $subIdx of $($subs.Count): $($sub.Name)" -PercentComplete (($subIdx / $subs.Count) * 100)

    $ErrorActionPreference = "Stop"
    Set-AzContext -SubscriptionId $sub.Id | Out-Null
  
    # VMs  
    if ($Selected.VM) {  
        try {
            $vmList = Get-AzVM
            if ($vmList) {
                $vmCount = 0
                foreach ($vm in $vmList) {  
                    $vmCount++
                    $vmPercentComplete = [math]::Round(($vmCount / $vmList.Count) * 100, 1)
                    Write-Progress -Id 2 -ParentId 1 -Activity "Processing Virtual Machines" -Status "Processing VM $vmCount of $($vmList.Count) - $vmPercentComplete% complete" -PercentComplete $vmPercentComplete
                    
                    # Calculate disk information
                $diskCount = 0
                $totalDiskSizeGB = 0         
                # OS Disk - get actual disk object for size
                if ($vm.StorageProfile.OsDisk) {
                    $diskCount++
                    try {
                        if ($vm.StorageProfile.OsDisk.DiskSizeGB) {
                        
                            $totalDiskSizeGB += $vm.StorageProfile.OsDisk.DiskSizeGB
                        } elseif($vm.StorageProfile.OsDisk.ManagedDisk) {
                            # Managed disk - get the disk resource
                            $osDiskName = $vm.StorageProfile.OsDisk.Name
                            $osDisk = Get-AzDisk -ResourceGroupName $vm.ResourceGroupName -DiskName $osDiskName -ErrorAction SilentlyContinue
                            if ($osDisk -and $osDisk.DiskSizeGB) {
                                $totalDiskSizeGB += $osDisk.DiskSizeGB
                            }
                        } else {
                            Write-Warning "Could not get data disk size for disk $($vm.StorageProfile.OsDisk.Name) on VM $($vm.Name)."

                        }
                    } catch {
                        Write-Warning "Could not get OS disk size for VM $($vm.Name): $_"
                    }
                }
                
                # Data Disks - get actual disk objects for sizes
                if ($vm.StorageProfile.DataDisks) {
                    $diskCount += $vm.StorageProfile.DataDisks.Count
                    foreach ($dataDisk in $vm.StorageProfile.DataDisks) {
                        try {
                            if ($dataDisk.DiskSizeGB) {
                              
                                $totalDiskSizeGB += $dataDisk.DiskSizeGB
                            } elseif ($dataDisk.ManagedDisk) {
                                # Managed disk and VM is Powered Off - get the disk resource
                                $diskName = $dataDisk.Name
                                $disk = Get-AzDisk -ResourceGroupName $vm.ResourceGroupName -DiskName $diskName -ErrorAction SilentlyContinue
                                if ($disk -and $disk.DiskSizeGB) {
                                    $totalDiskSizeGB += $disk.DiskSizeGB
                                }
                            } 
                        } catch {
                            Write-Warning "Could not get data disk size for disk $($dataDisk.Name) on VM $($vm.Name): $_"
                        }
                    }
                }
                
                $VMs += [PSCustomObject]@{  
                    Subscription   = $sub.Name  
                    ResourceGroup  = $vm.ResourceGroupName  
                    VMName         = $vm.Name  
                    VMSize         = $vm.HardwareProfile.VmSize  
                    OS             = $vm.StorageProfile.OsDisk.OsType  
                    Region         = $vm.Location  
                    VMId           = $vm.VMID
                    DiskCount      = $diskCount
                    VMDiskSizeGB   = $totalDiskSizeGB  
                }  
            }
            }
            Write-Progress -Id 2 -Activity "Processing Virtual Machines" -Completed
        } catch {
            Write-Warning "Error getting VMs: $_"
        }  
    }  
    # Storage Accounts  
    if ($Selected.STORAGE) {  
        try {
            $accounts = Get-AzStorageAccount  
            if ($accounts) {
                $saCount = 0
                foreach ($sa in $accounts) {  
                    $saCount++
                    $saPercentComplete = [math]::Round(($saCount / $accounts.Count) * 100, 1)
                    Write-Progress -Id 3 -ParentId 1 -Activity "Processing Storage Accounts" -Status "Processing Storage Account $saCount of $($accounts.Count) - $saPercentComplete% complete" -PercentComplete $saPercentComplete
                    
                    try {
                        # Use Get-AzMetric to get detailed storage account metrics
                    $resourceId = $sa.Id
                    $metrics = @("BlobCapacity", "ContainerCount", "BlobCount")
                    $containerCount = 0
                    $blobCount = 0
                    $blobCapacity = 0
                    
                    try {
                        $blobMetrics = Get-AzMetric -ResourceId "$resourceId/blobServices/default" -MetricNames $metrics -AggregationType Maximum -StartTime (Get-Date).AddDays(-1) -WarningAction SilentlyContinue
                        $containerCount = ($blobMetrics | Where-Object { $_.id -like "*ContainerCount" }).Data.Maximum | Select-Object -Last 1
                        $blobCount = ($blobMetrics | Where-Object { $_.id -like "*BlobCount" }).Data.Maximum | Select-Object -Last 1
                        $blobCapacity = ($blobMetrics | Where-Object { $_.id -like "*BlobCapacity" }).Data.Maximum | Select-Object -Last 1
                    } catch {
                        Write-Warning "Error getting blob metrics for $($sa.StorageAccountName): $_"
                    }
                    
                    $azSAObj = [ordered] @{}
                    $azSAObj.Add("StorageAccount",$sa.StorageAccountName)
                    $azSAObj.Add("StorageAccountType",$sa.Kind)
                    $azSAObj.Add("HNSEnabled(ADLSGen2)",$sa.EnableHierarchicalNamespace)
                    $azSAObj.Add("StorageAccountSkuName",$sa.Sku.Name)
                    $azSAObj.Add("StorageAccountAccessTier",$sa.AccessTier)
                    $azSAObj.Add("Subscription",$sub.Name)
                    $azSAObj.Add("Region",$sa.PrimaryLocation)
                    $azSAObj.Add("ResourceGroup",$sa.ResourceGroupName)
                    $azSAObj.Add("UsedCapacityBytes",$blobCapacity)
                    $azSAObj.Add("UsedCapacityGiB",[math]::round(($blobCapacity / 1073741824), 0))
                    $azSAObj.Add("UsedCapacityTiB",[math]::round(($blobCapacity / 1073741824 / 1024), 4))
                    $azSAObj.Add("UsedCapacityGB",[math]::round(($blobCapacity / 1000000000), 3))
                    $azSAObj.Add("UsedCapacityTB",[math]::round(($blobCapacity / 1000000000000), 4))
                    $azSAObj.Add("UsedBlobCapacityBytes",$blobCapacity)
                    $azSAObj.Add("UsedBlobCapacityGiB",[math]::round(($blobCapacity / 1073741824), 0))
                    $azSAObj.Add("UsedBlobCapacityTiB",[math]::round(($blobCapacity / 1073741824 / 1024), 4))
                    $azSAObj.Add("UsedBlobCapacityGB",[math]::round(($blobCapacity / 1000000000), 3))
                    $azSAObj.Add("UsedBlobCapacityTB",[math]::round(($blobCapacity / 1000000000000), 4))
                    $azSAObj.Add("BlobContainerCount",$containerCount)
                    $azSAObj.Add("BlobCount",$blobCount)
                    $StorageAccounts += New-Object -TypeName PSObject -Property $azSAObj
                } catch {
                    Write-Warning "Error getting storage metrics for $($sa.StorageAccountName): $_"
                }
            }
            }
            Write-Progress -Id 3 -Activity "Processing Storage Accounts" -Completed
        } catch {
            Write-Warning "Error getting storage accounts: $_"
        }  
    }  
}  

# Complete subscription progress
Write-Progress -Id 1 -Activity "Processing Azure Subscriptions" -Completed  
Write-Host "`n=== All Subscriptions Processed Successfully ===" -ForegroundColor Green
Write-Host "Total VMs found: $($VMs.Count)" -ForegroundColor Cyan
Write-Host "Total Storage Accounts found: $($StorageAccounts.Count)" -ForegroundColor Cyan

# Write all resources to CSV files
Write-Progress -Id 4 -Activity "Generating Output Files" -Status "Writing CSV files..." -PercentComplete 0

if ($Selected.VM -and $VMs.Count) { 
    Write-Progress -Id 4 -Activity "Generating Output Files" -Status "Writing VMs CSV..." -PercentComplete 25
    $VMs | Export-Csv (Join-Path $outdir "azure_vm_info_$dateStr.csv") -NoTypeInformation
    Write-Host "azure_vm_info_$dateStr.csv file has been written to $outdir" -ForegroundColor Cyan
}
if ($Selected.STORAGE -and $StorageAccounts.Count) { 
    Write-Progress -Id 4 -Activity "Generating Output Files" -Status "Writing Storage Accounts CSV..." -PercentComplete 50
    $StorageAccounts | Export-Csv (Join-Path $outdir "azure_storage_accounts_info_$dateStr.csv") -NoTypeInformation 
    Write-Host "azure_storage_accounts_info_$dateStr.csv file has been written to $outdir" -ForegroundColor Cyan
}

# Create comprehensive summary CSV  
$summaryRows = @()  

# Add overall resource type counts first
foreach ($k in $ResourceTypeMap.Keys) {  
    if ($Selected[$k]) {  
        # Calculate total disk size for VMs or storage capacity for Storage Accounts
        $ResourceType = ""
        $totalSize = 0
        $totalSizeTB = 0
        $totalSizeTiB = 0
        $count = 0
        
        if ($k -eq "VM" -and $VMs.Count -gt 0) {
            $ResourceType = "VM"
            $count = $VMs.Count
            $totalSize = ($VMs | Measure-Object -Property VMDiskSizeGB -Sum).Sum
            if ($totalSize -eq $null) { $totalSize = 0 }
            $totalSizeTB = [math]::Round($totalSize / 1000, 4)
            $totalSizeTiB = [math]::Round($totalSize / 1024, 4)
        } elseif ($k -eq "STORAGE" -and $StorageAccounts.Count -gt 0) {
            $ResourceType = "Storage Account"
            $count = $StorageAccounts.Count
            $totalCapacityBytes = ($StorageAccounts | Measure-Object -Property UsedCapacityBytes -Sum).Sum
            if ($totalCapacityBytes -eq $null) { $totalCapacityBytes = 0 }
            $totalSize = [math]::round(($totalCapacityBytes / 1000000000), 2)  # Convert bytes to GB
            $totalSizeTB = [math]::Round($totalCapacityBytes / 1000000000000, 4)  # Convert bytes to TB
            $totalSizeTiB = [math]::Round($totalCapacityBytes / 1099511627776, 4)  # Convert bytes to TiB
        }
        
        # Only add summary row if we have resources of this type
        if ($count -gt 0) {
            $summaryRows += [PSCustomObject]@{ 
                ResourceType = $ResourceType
                Region = "All"
                Count = $count
                TotalSizeGB = $totalSize
                TotalSizeTB = $totalSizeTB
                TotalSizeTiB = $totalSizeTiB
            }
        }  
    }  
}

# Add gap after overall totals
$summaryRows += [PSCustomObject]@{ 
    ResourceType = ""
    Region = ""
    Count = ""
    TotalSizeGB = ""
    TotalSizeTB = ""
    TotalSizeTiB = ""
}

# Add VM regional breakdown
if ($VMs.Count -and $Selected.VM) {
    # Add VM header with bracket formatting like Excel example
    $summaryRows += [PSCustomObject]@{ 
        ResourceType = "[ Azure VMs ]"
        Region = $null
        Count = $null
        TotalSizeGB = $null
        TotalSizeTB = $null
        TotalSizeTiB = $null
    }
    
    $vmRegionalSummary = $VMs | Group-Object Region | ForEach-Object {
        $totalDiskSize = ($_.Group | Measure-Object -Property VMDiskSizeGB -Sum).Sum
        if ($totalDiskSize -eq $null) { $totalDiskSize = 0 }
        $totalDiskSizeTB = [math]::Round($totalDiskSize / 1000, 4)
        $totalDiskSizeTiB = [math]::Round($totalDiskSize / 1024, 4)
        [PSCustomObject]@{
            ResourceType = ""
            Region = $_.Name
            Count = $_.Count
            TotalSizeGB = $totalDiskSize
            TotalSizeTB = $totalDiskSizeTB
            TotalSizeTiB = $totalDiskSizeTiB
        }
    } | Sort-Object Region
    
    $summaryRows += $vmRegionalSummary
    
    # Add gap after VM section
    $summaryRows += [PSCustomObject]@{ 
        ResourceType = ""
        Region = $null
        Count = $null
        TotalSizeGB = $null
        TotalSizeTB = $null
        TotalSizeTiB = $null
    }
}

# Add Storage Account regional breakdown if selected
if ($StorageAccounts.Count -and $Selected.STORAGE) {
    # Add Storage header with bracket formatting like Excel example
    $summaryRows += [PSCustomObject]@{ 
        ResourceType = "[ Azure Storage Accounts ]"
        Region = $null
        Count = $null
        TotalSizeGB = $null
        TotalSizeTB = $null
        TotalSizeTiB = $null
    }
    
    $storageRegionalSummary = $StorageAccounts | Group-Object Region | ForEach-Object {
        $totalCapacityBytes = ($_.Group | Measure-Object -Property UsedCapacityBytes -Sum).Sum
        if ($totalCapacityBytes -eq $null) { $totalCapacityBytes = 0 }
        $totalCapacityGB = [math]::round(($totalCapacityBytes / 1000000000), 2)  # Convert bytes to GB
        $totalCapacityTB = [math]::Round($totalCapacityBytes / 1000000000000, 4)  # Convert bytes to TB
        $totalCapacityTiB = [math]::Round($totalCapacityBytes / 1099511627776, 4)  # Convert bytes to TiB
        [PSCustomObject]@{
            ResourceType = ""
            Region = $_.Name
            Count = $_.Count
            TotalSizeGB = $totalCapacityGB
            TotalSizeTB = $totalCapacityTB
            TotalSizeTiB = $totalCapacityTiB
        }
    } | Sort-Object Region
    
    $summaryRows += $storageRegionalSummary
    
    # Add gap after Storage section
    $summaryRows += [PSCustomObject]@{ 
        ResourceType = ""
        Region = $null
        Count = $null
        TotalSizeGB = $null
        TotalSizeTB = $null
        TotalSizeTiB = $null
    }
}

if ($summaryRows.Count) {  
    Write-Progress -Id 4 -Activity "Generating Output Files" -Status "Writing comprehensive summary..." -PercentComplete 75
    $summaryRows | Export-Csv (Join-Path $outdir "azure_inventory_summary_$dateStr.csv") -NoTypeInformation  
    Write-Host "azure_inventory_summary_$dateStr.csv file has been written to $outdir" -ForegroundColor Cyan
}  

Write-Host "`n=== All Output Files Created Successfully ===" -ForegroundColor Green

Write-Progress -Id 4 -Activity "Generating Output Files" -Status "Creating ZIP archive..." -PercentComplete 90

Stop-Transcript
  
# Zip results  
$zipfile = Join-Path $PWD ("azure_sizing_" + $dateStr + ".zip")  
Add-Type -AssemblyName System.IO.Compression.FileSystem  
[IO.Compression.ZipFile]::CreateFromDirectory($outdir, $zipfile)  

# Complete all progress indicators
Write-Progress -Id 4 -Activity "Generating Output Files" -Completed

# Clean up - delete the output directory after ZIP creation
Write-Host "Cleaning up temporary files..." -ForegroundColor Yellow
Remove-Item -Path $outdir -Recurse -Force
Write-Host "Temporary directory removed: $outdir" -ForegroundColor Green
  
# Show final results on console
Write-Host "`nInventory complete. Results in $zipfile`n" -ForegroundColor Green
Write-Host "All output files have been compressed into the ZIP archive." -ForegroundColor Cyan
