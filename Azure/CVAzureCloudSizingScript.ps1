<#  
.SYNOPSIS  
    Azure Cloud Sizing Script - Comprehensive inventory and sizing analysis
.DESCRIPTION  
    Inventories Azure Virtual Machines, Storage Accounts, File Shares, and NetApp File Volumes across all or specified subscriptions.
    Calculates disk sizes for VMs, storage capacity utilization for Storage Accounts, capacity metrics for File Shares, and usage metrics for NetApp File Volumes.
    Generates detailed CSV reports with comprehensive sizing information in multiple units (GB, TB, TiB).
    Includes hierarchical progress tracking and comprehensive logging.
    Outputs timestamped CSV files and creates a ZIP archive of all results.

.PARAMETER Types
    Optional. Restrict inventory to specific resource types.
    Valid values: VM, Storage, FileShare, NetApp
    If not specified, all supported resource types will be inventoried.
    
.PARAMETER Subscriptions
    Optional. Target specific subscriptions by name or ID.
    If not specified, all accessible subscriptions will be processed.
    
.EXAMPLE  
    .\CVAzureCloudSizingScript.ps1  
    # Inventories all resources in all accessible subscriptions  
.EXAMPLE  
    .\CVAzureCloudSizingScript.ps1 -Types VM,Storage  
    # Explicitly inventories VMs and Storage Accounts in all subscriptions (same as default)
.EXAMPLE  
    .\CVAzureCloudSizingScript.ps1 -Types VM
    # Only inventories Virtual Machines in all subscriptions
.EXAMPLE  
    .\CVAzureCloudSizingScript.ps1 -Subscriptions "Production","Development"  
    # Inventories all resources in only the Production and Development subscriptions

.EXAMPLE  
    .\CVAzureCloudSizingScript.ps1 -Subscriptions "Dev Test","Production Environment"
    # Inventories all resources in subscriptions with spaces in names (always use quotes for names with spaces)

    # IMPORTANT: If you pass a subscription name that contains spaces WITHOUT quotes, PowerShell will treat the words as separate arguments
    # and the script will not match the subscription. Example of the problem and fixes:
    #   WRONG (will fail / be parsed incorrectly):
    #     .\CVAzureCloudSizingScript.ps1 -Subscriptions Dev Test
    #   CORRECT (use double quotes):
    #     .\CVAzureCloudSizingScript.ps1 -Subscriptions "Dev Test"
    #   ALTERNATIVE (use single quotes):
    #     .\CVAzureCloudSizingScript.ps1 -Subscriptions 'Dev Test'
    # You can also pass multiple quoted names separated by commas:
    #     .\CVAzureCloudSizingScript.ps1 -Subscriptions "Dev Test","Production Environment"

.EXAMPLE  
    .\CVAzureCloudSizingScript.ps1 -Subscriptions Production,Development
    # Inventories all resources in the subscriptions Production and Development (no spaces in names)

.EXAMPLE  
    .\CVAzureCloudSizingScript.ps1 -Types NetApp
    # Only inventories NetApp File volumes in all subscriptions
.EXAMPLE  
    .\CVAzureCloudSizingScript.ps1 -Types VM,Storage,NetApp -Subscriptions Production  
    # Inventories VMs, Storage Accounts, and NetApp File Volumes in only the Production subscription

.EXAMPLE  
    .\CVAzureCloudSizingScript.ps1 -Subscriptions xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx
    # Inventories all resources in the subscription with the specified Subscription ID

.OUTPUTS
    Creates timestamped output directory with the following files:
    - azure_vm_info_YYYY-MM-DD_HHMMSS.csv - VM inventory with disk sizing
    - azure_storage_accounts_info_YYYY-MM-DD_HHMMSS.csv - Storage Account inventory with capacity metrics
    - azure_file_shares_info_YYYY-MM-DD_HHMMSS.csv - File Share inventory with capacity metrics
    - azure_netapp_volumes_info_YYYY-MM-DD_HHMMSS.csv - NetApp Files volume inventory with capacity metrics
    - azure_inventory_summary_YYYY-MM-DD_HHMMSS.csv - Comprehensive summary with regional breakdowns
    - azure_sizing_script_output_YYYY-MM-DD_HHMMSS.log - Complete execution log
    - azure_sizing_YYYY-MM-DD_HHMMSS.zip - ZIP archive containing all output files
    
.NOTES
    Requires Azure PowerShell modules: Az.Accounts, Az.Compute, Az.Storage, Az.Monitor, Az.Resources, Az.NetAppFiles
    Script must be run by a user with appropriate Azure permissions to read VMs, Storage Accounts, File Shares, and NetApp Files
    VM disk sizing includes both OS disks and data disks with error handling for inaccessible disks
    Storage Account, File Share, and NetApp Files metrics are retrieved from Azure Monitor for the last 24 hours

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
       ./CVAzureCloudSizingScript.ps1 -Subscriptions Production,Development -Types VM,Storage

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
#>  
  
param(  
    [string[]]$Types, # Choices: VM, Storage, FileShare, NetApp  
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
    "FILESHARE"  = "FileShares"
    "NETAPP"     = "NetAppVolumes"
}  
  
# Normalize types  
if ($Types) {  
    $Types = $Types | ForEach-Object { $_.Trim().ToUpper() }  
    $Selected = @{}  
    foreach ($t in $Types) {  
        if ($ResourceTypeMap.ContainsKey($t)) { $Selected[$t] = $true }  
    }  
    if ($Selected.Count -eq 0) {  
        Write-Host "No valid -Types specified. Use: VM, Storage, FileShare, NetApp"  
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

# Helper function to determine if a storage account supports Azure Files
function Get-AzureFileSAs {
    param (
        [Parameter(Mandatory=$true)]
        [PSObject]$StorageAccount
    )

    return ($StorageAccount.Kind -in @('StorageV2', 'Storage') -and 
              $StorageAccount.Sku.Name -notin @('Premium_LRS', 'Premium_ZRS')) -or
              ($StorageAccount.Kind -eq 'FileStorage' -and 
              $StorageAccount.Sku.Name -in @('Premium_LRS', 'Premium_ZRS'))
}
if ($BlobLimit -gt 0) { Write-Host "  BlobLimit: $BlobLimit" -ForegroundColor Green }  
  
# Load modules  
$modules = @(  
    'Az.Accounts','Az.Compute','Az.Storage','Az.Monitor','Az.Resources','Az.NetAppFiles'  
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
    $notFoundSubs = @()
    
    foreach ($subFilter in $Subscriptions) {
        # Trim whitespace from the filter
        $cleanSubFilter = $subFilter.Trim()
        
        # Try exact match first (case-insensitive for names)
        $matchedSubs = $allSubs | Where-Object { 
            $_.Name.Trim() -eq $cleanSubFilter -or 
            $_.Id -eq $cleanSubFilter -or 
            $_.SubscriptionId -eq $cleanSubFilter 
        }
        
        # If no exact match, try case-insensitive name match
        if (-not $matchedSubs) {
            $matchedSubs = $allSubs | Where-Object { 
                $_.Name.Trim() -ieq $cleanSubFilter
            }
        }
        
        if ($matchedSubs) {
            $subs += $matchedSubs
            Write-Host "Found subscription: '$($matchedSubs.Name)' (ID: $($matchedSubs.Id))" -ForegroundColor Green
        } else {
            $notFoundSubs += $cleanSubFilter
            Write-Warning "Subscription '$cleanSubFilter' not found or not accessible"
        }
    }
    
    # Show available subscriptions if some weren't found
    if ($notFoundSubs.Count -gt 0) {
        Write-Host "`nAvailable subscriptions:" -ForegroundColor Yellow
        $allSubs | ForEach-Object { Write-Host "  - '$($_.Name)' (ID: $($_.Id))" -ForegroundColor Cyan }
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
$FileShares = @()  
$NetAppVolumes = @()  

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
    # Storage Accounts - Get all storage accounts once if either STORAGE or FILESHARE is selected
    if ($Selected.STORAGE -or $Selected.FILESHARE) {  
        try {
            $accounts = Get-AzStorageAccount  
            if ($accounts) {
                
                # Process Storage Account metrics if selected
                if ($Selected.STORAGE) {
                    $saCount = 0
                    foreach ($sa in $accounts) {  
                        $saCount++
                        $saPercentComplete = [math]::Round(($saCount / $accounts.Count) * 100, 1)
                        Write-Progress -Id 3 -ParentId 1 -Activity "Processing Storage Account Metrics" -Status "Processing Storage Account $saCount of $($accounts.Count) - $saPercentComplete% complete" -PercentComplete $saPercentComplete
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
                    Write-Progress -Id 3 -Activity "Processing Storage Account Metrics" -Completed
                }
                
                # Process File Shares if selected (separate progress tracking)
                if ($Selected.FILESHARE) {
                    $fileShareCount = 0
                    foreach ($sa in $accounts) {  
                        $fileShareCount++
                        $fsPercentComplete = [math]::Round(($fileShareCount / $accounts.Count) * 100, 1)
                        Write-Progress -Id 4 -ParentId 1 -Activity "Processing File Shares" -Status "Processing Storage Account $fileShareCount of $($accounts.Count) for File Shares - $fsPercentComplete% complete" -PercentComplete $fsPercentComplete
                        try {
                            # Check if this storage account supports Azure Files
                            if (Get-AzureFileSAs -StorageAccount $sa) {
                                $storageAccountFileShares = Get-AzRmStorageShare -StorageAccount $sa
                                $currentFileShareDetails = foreach ($fileShare in $storageAccountFileShares) {
                                    $storageAccountName = $fileShare.StorageAccountName
                                    $resourceGroupName = $fileShare.ResourceGroupName
                                    $shareName = $fileShare.Name
                                    Get-AzRmStorageShare -ResourceGroupName $resourceGroupName -StorageAccountName $storageAccountName -Name $shareName -GetShareUsage
                                }
                                
                                # Process each detailed file share from this storage account
                                foreach ($fileShareInfo in $currentFileShareDetails) {
                                    $fileShareObj = [ordered] @{}
                                    $fileShareObj.Add("Name", $fileShareInfo.Name)
                                    $fileShareObj.Add("StorageAccount", $sa.StorageAccountName)
                                    $fileShareObj.Add("StorageAccountType", $sa.Kind)
                                    $fileShareObj.Add("StorageAccountSkuName", $sa.Sku.Name)
                                    $fileShareObj.Add("StorageAccountAccessTier", $sa.AccessTier)
                                    # Determine file share-specific tier if available. If not present, set to 'Unknown' (do NOT fall back to storage account tier).
                                    $shareTier = $null
                                    if ($fileShareInfo -and $fileShareInfo.PSObject.Properties.Name -contains 'AccessTier') {
                                        $shareTier = $fileShareInfo.AccessTier
                                    } elseif ($fileShareInfo -and $fileShareInfo.Properties -and $fileShareInfo.Properties.AccessTier) {
                                        $shareTier = $fileShareInfo.Properties.AccessTier
                                    }
                                    if (-not $shareTier) { $shareTier = 'Unknown' }
                                    $fileShareObj.Add("ShareTier", $shareTier)
                                    $fileShareObj.Add("Subscription", $sub.Name)
                                    $fileShareObj.Add("Region", $sa.PrimaryLocation)
                                    $fileShareObj.Add("ResourceGroup", $sa.ResourceGroupName)
                                    $fileShareObj.Add("QuotaGiB", $fileShareInfo.QuotaGiB)
                                    $fileShareObj.Add("QuotaTiB", [math]::round(($fileShareInfo.QuotaGiB / 1024), 3))
                                    $fileShareObj.Add("UsedCapacityBytes", $fileShareInfo.ShareUsageBytes)
                                    $fileShareObj.Add("UsedCapacityGiB", [math]::round(($fileShareInfo.ShareUsageBytes / 1073741824), 0))
                                    $fileShareObj.Add("UsedCapacityTiB", [math]::round(($fileShareInfo.ShareUsageBytes / 1073741824 / 1024), 4))
                                    $fileShareObj.Add("UsedCapacityGB", [math]::round(($fileShareInfo.ShareUsageBytes / 1000000000), 3))
                                    $fileShareObj.Add("UsedCapacityTB", [math]::round(($fileShareInfo.ShareUsageBytes / 1000000000000), 4))

                                    $FileShares += New-Object -TypeName PSObject -Property $fileShareObj
                                }
                            } else {
                                Write-Verbose "Skipping File Share query for $($sa.StorageAccountName) because it does not support Azure Files."
                            }
                        } catch {
                            Write-Warning "Error getting Azure File Storage information from storage account $($sa.StorageAccountName) in subscription $($sub.Name): $_"
                        }
                    }
                    Write-Progress -Id 4 -Activity "Processing File Shares" -Completed
                }
            }
            Write-Progress -Id 3 -Activity "Processing Storage Accounts" -Completed
        } catch {
            Write-Warning "Error getting storage accounts: $_"
        }  
    }
    
    # NetApp Files
    if ($Selected.NETAPP) {
        try {
            # Time window for metric lookup
            $startTime = (Get-Date).AddHours(-1)
            $endTime = Get-Date
            $timeGrain = New-TimeSpan -Hours 1
            
            # Get all NetApp Files accounts in this subscription
            try {
                # First get all resource groups, then get NetApp accounts from each
                $resourceGroups = Get-AzResourceGroup
                $anfAccounts = @()
                
                foreach ($rg in $resourceGroups) {
                    try {
                        $rgNetAppAccounts = Get-AzNetAppFilesAccount -ResourceGroupName $rg.ResourceGroupName -ErrorAction SilentlyContinue
                        if ($rgNetAppAccounts) {
                            $anfAccounts += $rgNetAppAccounts
                        }
                    } catch {
                        # Continue silently if no NetApp accounts in this resource group
                    }
                }
            } catch {
                Write-Warning "Failed to get NetApp Files accounts in subscription $($sub.Name): $($_.Exception.Message)"
                $anfAccounts = $null
            }
            if ($anfAccounts) {
                $anfCount = 0
                $totalAnfAccounts = ($anfAccounts | Measure-Object).Count
                
                foreach ($account in $anfAccounts) {
                    $anfCount++
                    $anfPercentComplete = [math]::Round(($anfCount / $totalAnfAccounts) * 100, 1)
                    Write-Progress -Id 6 -ParentId 1 -Activity "Processing NetApp Files" -Status "Processing NetApp Account $anfCount of $totalAnfAccounts - $anfPercentComplete% complete" -PercentComplete $anfPercentComplete
        
                    try {
                        try {
                            $pools = Get-AzNetAppFilesPool -ResourceGroupName $account.ResourceGroupName -AccountName $account.Name
                        } catch {
                            Write-Warning "Failed to get capacity pools for NetApp account $($account.Name): $($_.Exception.Message)"
                            continue
                        }
                        
                        foreach ($pool in $pools) {
                            try {
                                # Extract just the pool name part (after the last '/' if present)
                                $poolName = if ($pool.Name -like '*/*') {
                                    ($pool.Name -split '/')[-1]
                                } else {
                                    $pool.Name
                                }
                                
                                $volumes = Get-AzNetAppFilesVolume -ResourceGroupName $account.ResourceGroupName -AccountName $account.Name -PoolName $poolName
                            } catch {
                                Write-Warning "Failed to get volumes for capacity pool $($pool.Name) in account $($account.Name): $($_.Exception.Message)"
                                continue
                            }
                            
                            foreach ($vol in $volumes) {
                                # Extract just the volume name part (after the last '/' if present)
                                $volumeName = if ($vol.Name -like '*/*') {
                                    ($vol.Name -split '/')[-1]
                                } else {
                                    $vol.Name
                                }
                                
                                try {
                                    # Get usage metric (LogicalSize)
                                    $usedBytes = 0
                                    try {
                                        # Use ResourceId approach
                                        $metric = Get-AzMetric -ResourceId $vol.Id -StartTime (Get-Date).AddDays(-1) -MetricName "VolumeLogicalSize" -AggregationType Average -WarningAction SilentlyContinue
                                        
                                        if ($metric.Data -and $metric.Data.Count -gt 0) {
                                            $usedBytes = $metric.Data[-1].Average
                                            if (-not $usedBytes) { $usedBytes = 0 }
                                        } else {
                                            Write-Warning "No usage metric data available for NetApp volume $volumeName"
                                        }
                                    } catch {
                                        Write-Warning "Could not get usage metrics for NetApp volume $volumeName - $($_.Exception.Message)"
                                    }
                                    
                                    $netAppObj = [ordered] @{}
                                    $netAppObj.Add("VolumeName", $volumeName)
                                    $netAppObj.Add("VolumeFullPath", $vol.Name)
                                    $netAppObj.Add("ResourceGroup", $account.ResourceGroupName)
                                    $netAppObj.Add("Subscription", $sub.Name)
                                    $netAppObj.Add("Region", $vol.Location)
                                    $netAppObj.Add("NetAppAccount", $account.Name)
                                    $netAppObj.Add("CapacityPool", $pool.Name)
                                    $netAppObj.Add("ProtocolType", ($vol.ProtocolTypes -join ", "))
                                    $netAppObj.Add("FilePath", $vol.CreationToken)
                                    $netAppObj.Add("ProvisionedGiB", [math]::Round($vol.UsageThreshold / 1GB, 2))
                                    $netAppObj.Add("ProvisionedTiB", [math]::Round($vol.UsageThreshold / 1TB, 4))
                                    $netAppObj.Add("ProvisionedGB", [math]::Round($vol.UsageThreshold / 1000000000, 2))
                                    $netAppObj.Add("ProvisionedTB", [math]::Round($vol.UsageThreshold / 1000000000000, 4))
                                    $netAppObj.Add("UsedCapacityBytes", $usedBytes)
                                    $netAppObj.Add("UsedCapacityGiB", [math]::Round($usedBytes / 1GB, 2))
                                    $netAppObj.Add("UsedCapacityTiB", [math]::Round($usedBytes / 1TB, 4))
                                    $netAppObj.Add("UsedCapacityGB", [math]::Round($usedBytes / 1000000000, 2))
                                    $netAppObj.Add("UsedCapacityTB", [math]::Round($usedBytes / 1000000000000, 4))
                                    $netAppObj.Add("ServiceLevel", $pool.ServiceLevel)
                                    $netAppObj.Add("PoolSizeGiB", [math]::Round($pool.Size / 1GB, 2))
                                    $netAppObj.Add("PoolSizeTiB", [math]::Round($pool.Size / 1TB, 4))
                                    
                                    $NetAppVolumes += New-Object -TypeName PSObject -Property $netAppObj
                                } catch {
                                    Write-Warning "Error processing NetApp volume $volumeName - $($_.Exception.Message)"
                                }
                            }
                        }
                    } catch {
                        Write-Warning "Error getting NetApp pools/volumes for account $($account.Name): $($_.Exception.Message)"
                    }
                }
                Write-Progress -Id 6 -Activity "Processing NetApp Files" -Completed
            } else {
                Write-Host "No NetApp Files accounts found in subscription $($sub.Name)" -ForegroundColor Yellow
            }
        } catch {
            Write-Warning "Error during NetApp Files processing in subscription $($sub.Name): $($_.Exception.Message)"
        }
    }
}  

# Complete subscription progress
Write-Progress -Id 1 -Activity "Processing Azure Subscriptions" -Completed  
Write-Host "`n=== All Subscriptions Processed Successfully ===" -ForegroundColor Green
if ($Selected.VM) { Write-Host "Total VMs found: $($VMs.Count)" -ForegroundColor Cyan }
if ($Selected.STORAGE) { Write-Host "Total Storage Accounts found: $($StorageAccounts.Count)" -ForegroundColor Cyan }
if ($Selected.FILESHARE) { Write-Host "Total File Shares found: $($FileShares.Count)" -ForegroundColor Cyan }
if ($Selected.NETAPP) { Write-Host "Total NetApp Volumes found: $($NetAppVolumes.Count)" -ForegroundColor Cyan }

# Write all resources to CSV files
Write-Progress -Id 5 -Activity "Generating Output Files" -Status "Writing CSV files..." -PercentComplete 0

if ($Selected.VM -and $VMs.Count) { 
    Write-Progress -Id 5 -Activity "Generating Output Files" -Status "Writing VMs CSV..." -PercentComplete 25
    $VMs | Export-Csv (Join-Path $outdir "azure_vm_info_$dateStr.csv") -NoTypeInformation
    Write-Host "azure_vm_info_$dateStr.csv file has been written to $outdir" -ForegroundColor Cyan
}
if ($Selected.STORAGE -and $StorageAccounts.Count) { 
    Write-Progress -Id 5 -Activity "Generating Output Files" -Status "Writing Storage Accounts CSV..." -PercentComplete 50
    $StorageAccounts | Export-Csv (Join-Path $outdir "azure_storage_accounts_info_$dateStr.csv") -NoTypeInformation 
    Write-Host "azure_storage_accounts_info_$dateStr.csv file has been written to $outdir" -ForegroundColor Cyan
}
if ($Selected.FILESHARE -and $FileShares.Count) { 
    Write-Progress -Id 5 -Activity "Generating Output Files" -Status "Writing File Shares CSV..." -PercentComplete 60
    $FileShares | Export-Csv (Join-Path $outdir "azure_file_shares_info_$dateStr.csv") -NoTypeInformation 
    Write-Host "azure_file_shares_info_$dateStr.csv file has been written to $outdir" -ForegroundColor Cyan
}
if ($Selected.NETAPP -and $NetAppVolumes.Count) { 
    Write-Progress -Id 5 -Activity "Generating Output Files" -Status "Writing NetApp Files CSV..." -PercentComplete 70
    $NetAppVolumes | Export-Csv (Join-Path $outdir "azure_netapp_volumes_info_$dateStr.csv") -NoTypeInformation 
    Write-Host "azure_netapp_volumes_info_$dateStr.csv file has been written to $outdir" -ForegroundColor Cyan
}

# Create comprehensive summary CSV  
$summaryRows = @()  

# Add overall resource type counts first
foreach ($k in $ResourceTypeMap.Keys) { 
    if ($Selected[$k]) {  
        # Calculate total disk size for VMs or storage capacity for Storage Accounts
        $Subscription = "All"
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
            $totalBlobCount = ($StorageAccounts | Measure-Object -Property BlobCount -Sum).Sum
            if ($totalBlobCount -eq $null) { $totalBlobCount = 0 }
            $ResourceType = "Storage Account (Total Blobs: $totalBlobCount)"
            $count = $StorageAccounts.Count
            $totalCapacityBytes = ($StorageAccounts | Measure-Object -Property UsedCapacityBytes -Sum).Sum
            if ($totalCapacityBytes -eq $null) { $totalCapacityBytes = 0 }
            $totalSize = [math]::round(($totalCapacityBytes / 1000000000), 2)  # Convert bytes to GB
            $totalSizeTB = [math]::Round($totalCapacityBytes / 1000000000000, 4)  # Convert bytes to TB
            $totalSizeTiB = [math]::Round($totalCapacityBytes / 1099511627776, 4)  # Convert bytes to TiB
        } elseif ($k -eq "FILESHARE" -and $FileShares.Count -gt 0) {
            $ResourceType = "File Share"
            $count = $FileShares.Count
            $totalCapacityBytes = ($FileShares | Measure-Object -Property UsedCapacityBytes -Sum).Sum
            if ($totalCapacityBytes -eq $null) { $totalCapacityBytes = 0 }
            $totalSize = [math]::round(($totalCapacityBytes / 1000000000), 2)  # Convert bytes to GB
            $totalSizeTB = [math]::Round($totalCapacityBytes / 1000000000000, 4)  # Convert bytes to TB
            $totalSizeTiB = [math]::Round($totalCapacityBytes / 1099511627776, 4)  # Convert bytes to TiB
        } elseif ($k -eq "NETAPP" -and $NetAppVolumes.Count -gt 0) {
            $ResourceType = "NetApp Files Volume"
            $count = $NetAppVolumes.Count
            $totalCapacityBytes = ($NetAppVolumes | Measure-Object -Property UsedCapacityBytes -Sum).Sum
            if ($totalCapacityBytes -eq $null) { $totalCapacityBytes = 0 }
            $totalSize = [math]::round(($totalCapacityBytes / 1000000000), 2)  # Convert bytes to GB
            $totalSizeTB = [math]::Round($totalCapacityBytes / 1000000000000, 4)  # Convert bytes to TB
            $totalSizeTiB = [math]::Round($totalCapacityBytes / 1099511627776, 4)  # Convert bytes to TiB
        }
        
        # Only add summary row if we have resources of this type
        if ($count -gt 0) {
            $summaryRows += [PSCustomObject]@{ 
                Subscription = $Subscription
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
    Subscription = ""
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
    
    $vmRegionalSummary = $VMs | Group-Object Region | ForEach-Object {
        $totalDiskSize = ($_.Group | Measure-Object -Property VMDiskSizeGB -Sum).Sum
        if ($totalDiskSize -eq $null) { $totalDiskSize = 0 }
        $totalDiskSizeTB = [math]::Round($totalDiskSize / 1000, 4)
        $totalDiskSizeTiB = [math]::Round($totalDiskSize / 1024, 4)
        [PSCustomObject]@{
            Subscription = "All"
            ResourceType = "VM"
            Region = $_.Name
            Count = $_.Count
            TotalSizeGB = $totalDiskSize
            TotalSizeTB = $totalDiskSizeTB
            TotalSizeTiB = $totalDiskSizeTiB
        }
    } | Sort-Object Region
    
    $summaryRows += $vmRegionalSummary
    
}

# Add Storage Account regional breakdown if selected
if ($StorageAccounts.Count -and $Selected.STORAGE) {
    # Add Storage header with bracket formatting like Excel example
    $storageRegionalSummary = $StorageAccounts | Group-Object Region | ForEach-Object {
        $totalCapacityBytes = ($_.Group | Measure-Object -Property UsedCapacityBytes -Sum).Sum
        if ($totalCapacityBytes -eq $null) { $totalCapacityBytes = 0 }
        $totalBlobCount = ($_.Group | Measure-Object -Property BlobCount -Sum).Sum
        if ($totalBlobCount -eq $null) { $totalBlobCount = 0 }
        $totalCapacityGB = [math]::round(($totalCapacityBytes / 1000000000), 2)  # Convert bytes to GB
        $totalCapacityTB = [math]::Round($totalCapacityBytes / 1000000000000, 4)  # Convert bytes to TB
        $totalCapacityTiB = [math]::Round($totalCapacityBytes / 1099511627776, 4)  # Convert bytes to TiB
        [PSCustomObject]@{
            Subscription = "All"
            ResourceType = "Storage Account (Total Blobs: $totalBlobCount)"
            Region = $_.Name
            Count = $_.Count
            TotalSizeGB = $totalCapacityGB
            TotalSizeTB = $totalCapacityTB
            TotalSizeTiB = $totalCapacityTiB
        }
    } | Sort-Object Region
    
    $summaryRows += $storageRegionalSummary  
}

# Add File Share regional breakdown if selected
if ($FileShares.Count -and $Selected.FILESHARE) {
    $fileShareRegionalSummary = $FileShares | Group-Object Region | ForEach-Object {
        $totalCapacityBytes = ($_.Group | Measure-Object -Property UsedCapacityBytes -Sum).Sum
        if ($totalCapacityBytes -eq $null) { $totalCapacityBytes = 0 }
        $totalCapacityGB = [math]::round(($totalCapacityBytes / 1000000000), 2)  # Convert bytes to GB
        $totalCapacityTB = [math]::Round($totalCapacityBytes / 1000000000000, 4)  # Convert bytes to TB
        $totalCapacityTiB = [math]::Round($totalCapacityBytes / 1099511627776, 4)  # Convert bytes to TiB
        [PSCustomObject]@{
            Subscription = "All"
            ResourceType = "File Share"
            Region = $_.Name
            Count = $_.Count
            TotalSizeGB = $totalCapacityGB
            TotalSizeTB = $totalCapacityTB
            TotalSizeTiB = $totalCapacityTiB
        }
    } | Sort-Object Region
    
    $summaryRows += $fileShareRegionalSummary  
}

# Add NetApp Files regional breakdown if selected
if ($NetAppVolumes.Count -and $Selected.NETAPP) {
    $netAppRegionalSummary = $NetAppVolumes | Group-Object Region | ForEach-Object {
        $totalCapacityBytes = ($_.Group | Measure-Object -Property UsedCapacityBytes -Sum).Sum
        if ($totalCapacityBytes -eq $null) { $totalCapacityBytes = 0 }
        $totalCapacityGB = [math]::round(($totalCapacityBytes / 1000000000), 2)  # Convert bytes to GB
        $totalCapacityTB = [math]::Round($totalCapacityBytes / 1000000000000, 4)  # Convert bytes to TB
        $totalCapacityTiB = [math]::Round($totalCapacityBytes / 1099511627776, 4)  # Convert bytes to TiB
        [PSCustomObject]@{
            Subscription = "All"
            ResourceType = "NetApp Files Volume"
            Region = $_.Name
            Count = $_.Count
            TotalSizeGB = $totalCapacityGB
            TotalSizeTB = $totalCapacityTB
            TotalSizeTiB = $totalCapacityTiB
        }
    } | Sort-Object Region
    
    $summaryRows += $netAppRegionalSummary  
}

# Add gap after overall summary
$summaryRows += [PSCustomObject]@{ 
    ResourceType = ""
    Region = $null
    Count = $null
    TotalSizeGB = $null
    TotalSizeTB = $null
    TotalSizeTiB = $null
}

# Add subscription-level summaries header
$summaryRows += [PSCustomObject]@{ 
    Subscription = "[ Subscription level Summary ]"
    ResourceType = ""
    Region = ""
    Count = ""
    TotalSizeGB = ""
    TotalSizeTB = ""
    TotalSizeTiB = ""
}

# Loop through each subscription we already processed
foreach ($sub in $subs) {
    $subscriptionName = $sub.Name
    # Add subscription header
    $summaryRows += [PSCustomObject]@{ 
        Subscription = $subscriptionName
        ResourceType = ""
        Region = ""
        Count = ""
        TotalSizeGB = ""
        TotalSizeTB = ""
        TotalSizeTiB = ""
    }
    
    # Process VMs for this subscription
    if ($Selected.VM) {
        $subscriptionVMs = $VMs | Where-Object { $_.Subscription -eq $subscriptionName }
        if ($subscriptionVMs.Count -gt 0) {
            # Add VM resource type total for this subscription
            $totalVMDiskSize = ($subscriptionVMs | Measure-Object -Property VMDiskSizeGB -Sum).Sum
            if ($totalVMDiskSize -eq $null) { $totalVMDiskSize = 0 }
            $totalVMDiskSizeTB = [math]::Round($totalVMDiskSize / 1000, 4)
            $totalVMDiskSizeTiB = [math]::Round($totalVMDiskSize / 1024, 4)
            
            $summaryRows += [PSCustomObject]@{
                Subscription = $subscriptionName
                ResourceType = "VM"
                Region = "All"
                Count = $subscriptionVMs.Count
                TotalSizeGB = $totalVMDiskSize
                TotalSizeTB = $totalVMDiskSizeTB
                TotalSizeTiB = $totalVMDiskSizeTiB
            }
            
            # Add regional breakdown for VMs in this subscription
            $vmRegionalBreakdown = $subscriptionVMs | Group-Object Region | ForEach-Object {
                $totalDiskSize = ($_.Group | Measure-Object -Property VMDiskSizeGB -Sum).Sum
                if ($totalDiskSize -eq $null) { $totalDiskSize = 0 }
                $totalDiskSizeTB = [math]::Round($totalDiskSize / 1000, 4)
                $totalDiskSizeTiB = [math]::Round($totalDiskSize / 1024, 4)
                [PSCustomObject]@{
                    Subscription = $subscriptionName
                    ResourceType = "VM"
                    Region = $_.Name
                    Count = $_.Count
                    TotalSizeGB = $totalDiskSize
                    TotalSizeTB = $totalDiskSizeTB
                    TotalSizeTiB = $totalDiskSizeTiB
                }
            } | Sort-Object Region
            
            $summaryRows += $vmRegionalBreakdown
        }
    }
    
    # Process Storage Accounts for this subscription
    if ($Selected.STORAGE) {
        $subscriptionStorage = $StorageAccounts | Where-Object { $_.Subscription -eq $subscriptionName }
        if ($subscriptionStorage.Count -gt 0) {
            # Add Storage resource type total for this subscription
            $totalStorageCapacityBytes = ($subscriptionStorage | Measure-Object -Property UsedCapacityBytes -Sum).Sum
            if ($totalStorageCapacityBytes -eq $null) { $totalStorageCapacityBytes = 0 }
            $totalBlobCount = ($subscriptionStorage | Measure-Object -Property BlobCount -Sum).Sum
            if ($totalBlobCount -eq $null) { $totalBlobCount = 0 }
            $totalStorageCapacityGB = [math]::round(($totalStorageCapacityBytes / 1000000000), 2)
            $totalStorageCapacityTB = [math]::Round($totalStorageCapacityBytes / 1000000000000, 4)
            $totalStorageCapacityTiB = [math]::Round($totalStorageCapacityBytes / 1099511627776, 4)
            
            $summaryRows += [PSCustomObject]@{
                Subscription = $subscriptionName
                ResourceType = "Storage Account (Total Blobs: $totalBlobCount)"
                Region = "All"
                Count = $subscriptionStorage.Count
                TotalSizeGB = $totalStorageCapacityGB
                TotalSizeTB = $totalStorageCapacityTB
                TotalSizeTiB = $totalStorageCapacityTiB
            }
            
            # Add regional breakdown for Storage Accounts in this subscription
            $storageRegionalBreakdown = $subscriptionStorage | Group-Object Region | ForEach-Object {
                $totalCapacityBytes = ($_.Group | Measure-Object -Property UsedCapacityBytes -Sum).Sum
                if ($totalCapacityBytes -eq $null) { $totalCapacityBytes = 0 }
                $totalBlobCount = ($_.Group | Measure-Object -Property BlobCount -Sum).Sum
                if ($totalBlobCount -eq $null) { $totalBlobCount = 0 }
                $totalCapacityGB = [math]::round(($totalCapacityBytes / 1000000000), 2)
                $totalCapacityTB = [math]::Round($totalCapacityBytes / 1000000000000, 4)
                $totalCapacityTiB = [math]::Round($totalCapacityBytes / 1099511627776, 4)
                [PSCustomObject]@{
                    Subscription = $subscriptionName
                    ResourceType = "Storage Account (Total Blobs: $totalBlobCount)"
                    Region = $_.Name
                    Count = $_.Count
                    TotalSizeGB = $totalCapacityGB
                    TotalSizeTB = $totalCapacityTB
                    TotalSizeTiB = $totalCapacityTiB
                }
            } | Sort-Object Region
            
            $summaryRows += $storageRegionalBreakdown
        }
    }
    
    # Process File Shares for this subscription
    if ($Selected.FILESHARE) {
        $subscriptionFileShares = $FileShares | Where-Object { $_.Subscription -eq $subscriptionName }
        if ($subscriptionFileShares.Count -gt 0) {
            # Add File Share resource type total for this subscription
            $totalFileShareCapacityBytes = ($subscriptionFileShares | Measure-Object -Property UsedCapacityBytes -Sum).Sum
            if ($totalFileShareCapacityBytes -eq $null) { $totalFileShareCapacityBytes = 0 }
            $totalFileShareCapacityGB = [math]::round(($totalFileShareCapacityBytes / 1000000000), 2)
            $totalFileShareCapacityTB = [math]::Round($totalFileShareCapacityBytes / 1000000000000, 4)
            $totalFileShareCapacityTiB = [math]::Round($totalFileShareCapacityBytes / 1099511627776, 4)
            
            $summaryRows += [PSCustomObject]@{
                Subscription = $subscriptionName
                ResourceType = "File Share"
                Region = "All"
                Count = $subscriptionFileShares.Count
                TotalSizeGB = $totalFileShareCapacityGB
                TotalSizeTB = $totalFileShareCapacityTB
                TotalSizeTiB = $totalFileShareCapacityTiB
            }
            
            # Add regional breakdown for File Shares in this subscription
            $fileShareRegionalBreakdown = $subscriptionFileShares | Group-Object Region | ForEach-Object {
                $totalCapacityBytes = ($_.Group | Measure-Object -Property UsedCapacityBytes -Sum).Sum
                if ($totalCapacityBytes -eq $null) { $totalCapacityBytes = 0 }
                $totalCapacityGB = [math]::round(($totalCapacityBytes / 1000000000), 2)
                $totalCapacityTB = [math]::Round($totalCapacityBytes / 1000000000000, 4)
                $totalCapacityTiB = [math]::Round($totalCapacityBytes / 1099511627776, 4)
                [PSCustomObject]@{
                    Subscription = $subscriptionName
                    ResourceType = "File Share"
                    Region = $_.Name
                    Count = $_.Count
                    TotalSizeGB = $totalCapacityGB
                    TotalSizeTB = $totalCapacityTB
                    TotalSizeTiB = $totalCapacityTiB
                }
            } | Sort-Object Region
            
            $summaryRows += $fileShareRegionalBreakdown
        }
    }
    
    # Process NetApp Files for this subscription
    if ($Selected.NETAPP) {
        $subscriptionNetAppVolumes = $NetAppVolumes | Where-Object { $_.Subscription -eq $subscriptionName }
        if ($subscriptionNetAppVolumes.Count -gt 0) {
            # Add NetApp Files resource type total for this subscription
            $totalNetAppCapacityBytes = ($subscriptionNetAppVolumes | Measure-Object -Property UsedCapacityBytes -Sum).Sum
            if ($totalNetAppCapacityBytes -eq $null) { $totalNetAppCapacityBytes = 0 }
            $totalNetAppCapacityGB = [math]::round(($totalNetAppCapacityBytes / 1000000000), 2)
            $totalNetAppCapacityTB = [math]::Round($totalNetAppCapacityBytes / 1000000000000, 4)
            $totalNetAppCapacityTiB = [math]::Round($totalNetAppCapacityBytes / 1099511627776, 4)
            
            $summaryRows += [PSCustomObject]@{
                Subscription = $subscriptionName
                ResourceType = "NetApp Files Volume"
                Region = "All"
                Count = $subscriptionNetAppVolumes.Count
                TotalSizeGB = $totalNetAppCapacityGB
                TotalSizeTB = $totalNetAppCapacityTB
                TotalSizeTiB = $totalNetAppCapacityTiB
            }
            
            # Add regional breakdown for NetApp Files in this subscription
            $netAppRegionalBreakdown = $subscriptionNetAppVolumes | Group-Object Region | ForEach-Object {
                $totalCapacityBytes = ($_.Group | Measure-Object -Property UsedCapacityBytes -Sum).Sum
                if ($totalCapacityBytes -eq $null) { $totalCapacityBytes = 0 }
                $totalCapacityGB = [math]::round(($totalCapacityBytes / 1000000000), 2)
                $totalCapacityTB = [math]::Round($totalCapacityBytes / 1000000000000, 4)
                $totalCapacityTiB = [math]::Round($totalCapacityBytes / 1099511627776, 4)
                [PSCustomObject]@{
                    Subscription = $subscriptionName
                    ResourceType = "NetApp Files Volume"
                    Region = $_.Name
                    Count = $_.Count
                    TotalSizeGB = $totalCapacityGB
                    TotalSizeTB = $totalCapacityTB
                    TotalSizeTiB = $totalCapacityTiB
                }
            } | Sort-Object Region
            
            $summaryRows += $netAppRegionalBreakdown
        }
    }
    
    # Add gap after each subscription
    $summaryRows += [PSCustomObject]@{ 
        Subscription = ""
        ResourceType = ""
        Region = ""
        Count = ""
        TotalSizeGB = ""
        TotalSizeTB = ""
        TotalSizeTiB = ""
    }
}


# Export summary if we have any rows
if ($summaryRows.Count) {  
    Write-Progress -Id 5 -Activity "Generating Output Files" -Status "Writing comprehensive summary..." -PercentComplete 75
    $summaryRows | Export-Csv (Join-Path $outdir "azure_inventory_summary_$dateStr.csv") -NoTypeInformation  
    Write-Host "azure_inventory_summary_$dateStr.csv file has been written to $outdir" -ForegroundColor Cyan
}  

Write-Host "`n=== All Output Files Created Successfully ===" -ForegroundColor Green

Write-Progress -Id 5 -Activity "Generating Output Files" -Status "Creating ZIP archive..." -PercentComplete 90

Stop-Transcript

# Zip results  
$zipfile = Join-Path $PWD ("azure_sizing_" + $dateStr + ".zip")  
Add-Type -AssemblyName System.IO.Compression.FileSystem  
[IO.Compression.ZipFile]::CreateFromDirectory($outdir, $zipfile)  

# Complete all progress indicators
Write-Progress -Id 5 -Activity "Generating Output Files" -Completed

# Clean up - delete the output directory after ZIP creation
Write-Host "Cleaning up temporary files..." -ForegroundColor Yellow
Remove-Item -Path $outdir -Recurse -Force
Write-Host "Temporary directory removed: $outdir" -ForegroundColor Green
  
# Show final results on console
Write-Host "`nInventory complete. Results in $zipfile`n" -ForegroundColor Green
Write-Host "All output files have been compressed into the ZIP archive. Please provide to Commvault representative." -ForegroundColor Cyan