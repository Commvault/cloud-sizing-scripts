<#  
.SYNOPSIS  
    Azure Cloud Sizing Script - Comprehensive inventory and sizing analysis
.DESCRIPTION  
    Inventories Azure Virtual Machines, Storage Accounts, File Shares, NetApp File Volumes, SQL Databases/Managed Instances, MySQL Servers, PostgreSQL Servers, and CosmosDB Accounts across all or specified subscriptions.
    Calculates disk sizes for VMs, storage capacity utilization for Storage Accounts, capacity metrics for File Shares, usage metrics for NetApp File Volumes, and database storage metrics for SQL, MySQL, PostgreSQL, and CosmosDB resources.
    Generates detailed CSV reports with comprehensive sizing information in multiple units (GB, TB, TiB).
    Includes hierarchical progress tracking and comprehensive logging.
    Outputs timestamped CSV files and creates a ZIP archive of all results.

.PARAMETER Types
    Optional. Restrict inventory to specific resource types.
    Valid values: VM, Storage, FileShare, NetApp, SQL, Cosmos
    If not specified, all supported resource types will be inventoried.
    Note: MySQL and PostgreSQL servers are included as part of SQL inventory.
    
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
    .\CVAzureCloudSizingScript.ps1 -Types Cosmos -Subscriptions "Development","Staging"  
    # Only inventories CosmosDB accounts in the Development and Staging subscriptions

.EXAMPLE  
    .\CVAzureCloudSizingScript.ps1 -Subscriptions xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx
    # Inventories all resources in the subscription with the specified Subscription ID

.OUTPUTS
    Creates timestamped output directory with the following files:
    - azure_vm_info_YYYY-MM-DD_HHMMSS.csv - VM inventory with disk sizing
    - azure_storage_accounts_info_YYYY-MM-DD_HHMMSS.csv - Storage Account inventory with capacity metrics
    - azure_file_shares_info_YYYY-MM-DD_HHMMSS.csv - File Share inventory with capacity metrics
    - azure_netapp_volumes_info_YYYY-MM-DD_HHMMSS.csv - NetApp Files volume inventory with capacity metrics
    - azure_sql_managed_instances_YYYY-MM-DD_HHMMSS.csv - SQL Managed Instances inventory
    - azure_sql_databases_inventory_YYYY-MM-DD_HHMMSS.csv - SQL Databases inventory  
    - azure_mysql_servers_YYYY-MM-DD_HHMMSS.csv - MySQL Servers inventory (included with SQL inventory)
    - azure_postgresql_servers_YYYY-MM-DD_HHMMSS.csv - PostgreSQL Servers inventory (included with SQL inventory)
    - azure_cosmosdb_accounts_YYYY-MM-DD_HHMMSS.csv - CosmosDB Accounts inventory with storage metrics
    - azure_inventory_summary_YYYY-MM-DD_HHMMSS.csv - Comprehensive summary with regional breakdowns
    - azure_sizing_script_output_YYYY-MM-DD_HHMMSS.log - Complete execution log
    - azure_sizing_YYYY-MM-DD_HHMMSS.zip - ZIP archive containing all output files
    
.NOTES
    Requires Azure PowerShell modules: Az.Accounts, Az.Compute, Az.Storage, Az.Monitor, Az.Resources, Az.NetAppFiles, Az.CosmosDB, Az.MySql, Az.PostgreSql
    Script must be run by a user with appropriate Azure permissions to read VMs, Storage Accounts, File Shares, NetApp Files, SQL resources, CosmosDB accounts, MySQL servers, and PostgreSQL servers
    VM disk sizing includes both OS disks and data disks with error handling for inaccessible disks
    Storage Account, File Share, NetApp Files, CosmosDB, MySQL, and PostgreSQL metrics are retrieved from Azure Monitor for the last 1 hour
    
    NOTE: All Azure Monitor metrics collected use Maximum aggregation over a 1-hour time period to capture current resource utilization values efficiently.
    
    IMPORTANT: If Azure Monitor metrics are reported inaccurately or are not available for certain resources, there may be discrepancies between the reported values in this script and actual resource utilization.
    This can occur due to Azure Monitor data delays, resource configuration issues, or temporary service interruptions.
    
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
    [string[]]$Types, # Choices: VM, Storage, FileShare, NetApp, SQL  
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
    "SQL"        = "SqlInventory"
    "COSMOS"     = "CosmosDBs"
}  
  
# Normalize types  
if ($Types) {  
    $Types = $Types | ForEach-Object { $_.Trim().ToUpper() }  
    $Selected = @{}  
    foreach ($t in $Types) {  
        if ($ResourceTypeMap.ContainsKey($t)) { $Selected[$t] = $true }  
    }  
    if ($Selected.Count -eq 0) {  
    Write-Host "No valid -Types specified. Use: VM, Storage, FileShare, NetApp, SQL, Cosmos"  
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
    'Az.Accounts','Az.Compute','Az.Storage','Az.Monitor','Az.Resources','Az.NetAppFiles','Az.CosmosDB','Az.MySql','Az.PostgreSql'  
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
# Renamed SQL collections to reduce similarity with original source
$SqlInstancesInventory = @()
$SqlDbInventory = @()
$SqlMIDbInventory = @()
$CosmosDBs = @()
$MySQLServers = @()
$PostgreSQLServers = @()

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
            Write-Host "Processing Virtual Machines in subscription $($sub.Name)" -ForegroundColor Green
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
                    Write-Host "Processing Storage Accounts in subscription $($sub.Name)" -ForegroundColor Green
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
                            $blobMetrics = Get-AzMetric -ResourceId "$resourceId/blobServices/default" -MetricName $metrics -AggregationType Maximum -StartTime (Get-Date).AddHours(-1) -WarningAction SilentlyContinue
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
                    Write-Host "Processing File Shares in subscription $($sub.Name)" -ForegroundColor Green
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
                                    
                                    if ($fileShareInfo.EnabledProtocols) {
                                        $fileShareObj.Add("ProtocolType", ($fileShareInfo.EnabledProtocols -join ", "))
                                    } elseif ($sa.Kind -eq 'StorageV2') {
                                        # StorageV2 accounts primarily support SMB protocol
                                        $fileShareObj.Add("ProtocolType", "SMB")
                                    } else {
                                        $fileShareObj.Add("ProtocolType", "")
                                    }
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
            Write-Host "Processing NetApp Volumes in subscription $($sub.Name)" -ForegroundColor Green
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
                    Write-Progress -Id 6 -ParentId 1 -Activity "Processing NetApp Volumes" -Status "Processing NetApp Account $anfCount of $totalAnfAccounts - $anfPercentComplete% complete" -PercentComplete $anfPercentComplete
        
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
                                        $metric = Get-AzMetric -ResourceId $vol.Id -StartTime (Get-Date).AddHours(-1) -MetricName "VolumeLogicalSize" -AggregationType Average -WarningAction SilentlyContinue
                                        
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
                Write-Progress -Id 6 -Activity "Processing NetApp Volumes" -Completed
            } else {
                Write-Host "No NetApp Files accounts found in subscription $($sub.Name)" -ForegroundColor Yellow
            }
        } catch {
            Write-Warning "Error during NetApp Files processing in subscription $($sub.Name): $($_.Exception.Message)"
        }
    }
    
    # SQL Managed Instances
    if ($Selected.SQL) {
        try {
            Write-Host "Processing SQL Managed Instances in subscription $($sub.Name)" -ForegroundColor Green
            $sqlMIList = Get-AzSqlInstance

            if ($sqlMIList) {
                $sqlMICount = 0
                foreach ($sqlMI in $sqlMIList) {
                    $sqlMICount++
                    $sqlMIPercentComplete = [math]::Round(($sqlMICount / $sqlMIList.Count) * 100, 1)
                    Write-Progress -Id 7 -ParentId 1 -Activity "Processing SQL Managed Instances" -Status "Processing SQL MI $sqlMICount of $($sqlMIList.Count) - $sqlMIPercentComplete% complete" -PercentComplete $sqlMIPercentComplete
                    
                    try {
                        # Get storage metrics for SQL MI
                        $storageUsedGB = 0
                        $storageAllocatedGB = 0
                        $storageUsedMB = 0
                        # Binary TiB values computed from MB (1024-based) for clarity
                        $storageUsedTiBBinary = 0
                        $storageAllocatedTiBBinary = 0
                        
                        try {
                            # Get storage usage metrics from Azure Monitor - using shorter time window for efficiency
                            $storageMetrics = Get-AzMetric -ResourceId $sqlMI.Id -StartTime (Get-Date).AddHours(-1) -MetricNames @("storage_space_used_mb") -AggregationType Maximum -WarningAction SilentlyContinue
                            
                            if ($storageMetrics) {
                                foreach ($metric in $storageMetrics) {
                                    if ($metric.Data -and $metric.Data.Count -gt 0) {
                                        $latestValue = $metric.Data[-1].Maximum
                                        if (-not $latestValue) { $latestValue = 0 }
                                        switch ($metric.Name.Value) {
                                            "storage_space_used_mb" {
                                                $valueMiB = [double]$latestValue
                                                $storageUsedMB = $valueMiB
                                                $valueBytes = $valueMiB * 1048576.0    # 1 MiB = 1,048,576 bytes (IEC)
                                                # Decimal GB (SI): 1 GB = 1,000,000,000 bytes
                                                $storageUsedGB = [math]::Round($valueBytes / 1000000000.0, 2)
                                                # Binary TiB (IEC): 1 TiB = 1024^4 bytes
                                                $storageUsedTiBBinary = [math]::Round($valueBytes / 1099511627776.0, 4)
                                                                    }
                                                                    "reserved_storage_mb" {
                                                $valueMiB = [double]$latestValue
                                                $valueBytes = $valueMiB * 1048576.0
                                                $storageAllocatedGB = [math]::Round($valueBytes / 1000000000.0, 2)
                                                $storageAllocatedTiBBinary = [math]::Round($valueBytes / 1099511627776.0, 4)
                                            }
                                        }
                                    }
                                }
                                Write-Verbose "Retrieved storage metrics for SQL MI $($sqlMI.ManagedInstanceName) - Used: $storageUsedGB GB, Allocated: $storageAllocatedGB GB"
                            }
                        } catch {
                            Write-Warning "Could not get storage metrics for SQL Managed Instance $($sqlMI.ManagedInstanceName): $($_.Exception.Message)"
                        }
                        
                        $sqlMIObj = [ordered] @{}
                        $sqlMIObj.Add("Subscription", $sub.Name)
                        $sqlMIObj.Add("ResourceGroup", $sqlMI.ResourceGroupName)
                        $sqlMIObj.Add("ManagedInstanceName", $sqlMI.ManagedInstanceName)
                        $sqlMIObj.Add("Region", $sqlMI.Location)
                        $sqlMIObj.Add("vCores", $sqlMI.VCores)
                        $sqlMIObj.Add("StorageSizeGB", $sqlMI.StorageSizeInGB)
                        # StorageUsedGB is decimal (SI). StorageUsedTiB is binary (IEC) computed directly from MB.
                        $sqlMIObj.Add("StorageUsedMB", $storageUsedMB)
                        $sqlMIObj.Add("StorageUsedGB", $storageUsedGB)
                        $sqlMIObj.Add("StorageUsedTiB", $storageUsedTiBBinary)
                        $sqlMIObj.Add("StorageUsedTB", [math]::Round($storageUsedGB / 1000, 4))
                        $sqlMIObj.Add("StorageAllocatedGB", $storageAllocatedGB)
                        $sqlMIObj.Add("StorageAllocatedTiB", $storageAllocatedTiBBinary)
                        $sqlMIObj.Add("StorageAllocatedTB", [math]::Round($storageAllocatedGB / 1000, 4))
                        $sqlMIObj.Add("LicenseType", $sqlMI.LicenseType)
                        $sqlMIObj.Add("State", $sqlMI.State)
                        $sqlMIObj.Add("SubnetId", $sqlMI.SubnetId)
                        
                        $SqlInstancesInventory += New-Object -TypeName PSObject -Property $sqlMIObj
                        Write-Verbose "Successfully processed SQL Managed Instance: $($sqlMI.ManagedInstanceName)"
                        
                        # Get databases for this SQL Managed Instance (not added to summary)
                        try {
                            Write-Verbose "Collecting databases for SQL MI: $($sqlMI.ManagedInstanceName)"
                            $miDatabases = Get-AzSqlInstanceDatabase -InstanceName $sqlMI.ManagedInstanceName -ResourceGroupName $sqlMI.ResourceGroupName -ErrorAction SilentlyContinue
                            
                            if ($miDatabases) {
                                foreach ($miDb in $miDatabases) {
                                    # Skip system databases
                                    if ($miDb.Name -in @('master', 'msdb', 'tempdb', 'model')) { continue }
                                    
                                    # Get retention policies for this database
                                    $ltrPolicyJson = $null
                                    $strPolicyJson = $null
                                    
                                    try {
                                        # Get Long-Term Retention policy
                                        $ltrPolicy = Get-AzSqlInstanceDatabaseBackupLongTermRetentionPolicy -InstanceName $sqlMI.ManagedInstanceName -DatabaseName $miDb.Name -ResourceGroupName $sqlMI.ResourceGroupName -ErrorAction SilentlyContinue
                                        if ($ltrPolicy) {
                                            $ltrPolicyJson = $ltrPolicy | ConvertTo-Json -Compress
                                        }
                                    } catch {
                                        Write-Warning "Could not get LTR policy for database $($miDb.Name): $($_.Exception.Message)"
                                    }
                                    
                                    try {
                                        # Get Short-Term Retention policy
                                        $strPolicy = Get-AzSqlInstanceDatabaseBackupShortTermRetentionPolicy -InstanceName $sqlMI.ManagedInstanceName -DatabaseName $miDb.Name -ResourceGroupName $sqlMI.ResourceGroupName -ErrorAction SilentlyContinue
                                        if ($strPolicy) {
                                            $strPolicyJson = $strPolicy | ConvertTo-Json -Compress
                                        }
                                    } catch {
                                        Write-Warning "Could not get STR policy for database $($miDb.Name): $($_.Exception.Message)"
                                    }
                                    
                                    $miDbObj = [ordered]@{
                                        "Subscription" = $sub.Name
                                        "ResourceGroup" = $sqlMI.ResourceGroupName
                                        "ManagedInstanceName" = $sqlMI.ManagedInstanceName
                                        "DatabaseName" = $miDb.Name
                                        "Status" = $miDb.Status
                                        "CreationDate" = $miDb.CreationDate
                                        "Collation" = $miDb.Collation
                                        "Region" = $sqlMI.Location
                                        "LongTermRetentionPolicyJson" = $ltrPolicyJson
                                        "ShortTermRetentionPolicyJson" = $strPolicyJson
                                    }
                                    
                                    $SqlMIDbInventory += New-Object -TypeName PSObject -Property $miDbObj
                                }
                                Write-Verbose "Collected $($miDatabases.Count) databases for SQL MI: $($sqlMI.ManagedInstanceName)"
                            }
                        } catch {
                            Write-Warning "Could not get databases for SQL Managed Instance $($sqlMI.ManagedInstanceName): $($_.Exception.Message)"
                        }
                    } catch {
                        Write-Warning "Error processing SQL Managed Instance $($sqlMI.ManagedInstanceName): $($_.Exception.Message)"
                    }
                }
                Write-Progress -Id 7 -Activity "Processing SQL Managed Instances" -Completed
            } else {
                Write-Host "No SQL Managed Instances found in subscription $($sub.Name)" -ForegroundColor Yellow
            }
        } catch {
            Write-Warning "Error during SQL Managed Instances processing in subscription $($sub.Name): $($_.Exception.Message)"
        }
    }

    # SQL Logical Servers & Databases
    if ($Selected.SQL) {
        try {
            Write-Host "Processing Azure SQL Servers and Databases in subscription $($sub.Name)" -ForegroundColor Green
            try {
                $sqlServers = Get-AzSqlServer -ErrorAction Stop
            } catch {
                Write-Warning "Unable to list SQL servers in subscription $($sub.Name): $($_.Exception.Message)"
                $sqlServers = $null
            }

            if ($sqlServers) {
                $sqlServerIdx = 0
                foreach ($sqlServer in $sqlServers) {
                    $sqlServerIdx++
                    Write-Progress -Id 8 -ParentId 1 -Activity "Processing SQL Servers" -Status "Server $sqlServerIdx of $($sqlServers.Count) - $($sqlServer.ServerName)" -PercentComplete ([math]::Round(($sqlServerIdx / $sqlServers.Count) * 100,1))
                    try {
                        $sqlDBs = Get-AzSqlDatabase -ServerName $sqlServer.ServerName -ResourceGroupName $sqlServer.ResourceGroupName -ErrorAction Stop
                    } catch {
                        Write-Warning "Could not enumerate databases for SQL server $($sqlServer.ServerName): $($_.Exception.Message)"
                        continue
                    }

                    foreach ($sqlDB in $sqlDBs) {
                        # Skip system DBs
                        if ($sqlDB.SkuName -eq 'System') { continue }

                        $sqlObj = [ordered]@{}
                        $sqlObj.Add("Subscription", $sub.Name)
                        $sqlObj.Add("ResourceGroup", $sqlServer.ResourceGroupName)
                        $sqlObj.Add("Server", $sqlServer.ServerName)
                        $sqlObj.Add("Database", $sqlDB.DatabaseName)
                        $sqlObj.Add("Edition", $sqlDB.Edition)
                        $sqlObj.Add("InstanceType", $sqlDB.SkuName)
                        $sqlObj.Add("MaxSizeGiB", [math]::Round(($sqlDB.MaxSizeBytes / 1073741824), 0))
                        $sqlObj.Add("MaxSizeGB", [math]::Round(($sqlDB.MaxSizeBytes / 1000000000), 3))
                        $sqlObj.Add("Region", $sqlDB.Location)
                        $sqlObj.Add("DatabaseId", $sqlDB.DatabaseId)
                        $sqlObj.Add("Status", $sqlDB.Status)

                        # Get allocated and used storage metrics (Maximum over last 24h)
                        $allocatedVal = $null
                        $usedVal = $null
                        try {
                            # Single call for both metrics reduces API calls and keeps timing/aggregation identical
                            $metrics = Get-AzMetric -ResourceId $sqlDB.ResourceId -MetricNames @("allocated_data_storage","storage") -AggregationType Maximum -StartTime (Get-Date).AddHours(-1) -WarningAction SilentlyContinue

                            # Map returned metrics back to the named variables for compatibility with existing logic
                            $allocatedMetric = $null
                            $usedMetric = $null
                            if ($metrics) {
                                $allocatedMetric = $metrics | Where-Object { $_.Name.Value -eq "allocated_data_storage" } | Select-Object -First 1
                                $usedMetric = $metrics | Where-Object { $_.Name.Value -eq "storage" } | Select-Object -First 1
                            }

                            if ($allocatedMetric -and $allocatedMetric.Data -and $allocatedMetric.Data.Count -gt 0) {
                                # pick the last non-null Maximum datapoint
                                $allocatedVal = ($allocatedMetric.Data | ForEach-Object { $_.Maximum } | Where-Object { $_ -ne $null } | Select-Object -Last 1)
                            }
                            if ($usedMetric -and $usedMetric.Data -and $usedMetric.Data.Count -gt 0) {
                                $usedVal = ($usedMetric.Data | ForEach-Object { $_.Maximum } | Where-Object { $_ -ne $null } | Select-Object -Last 1)
                            }
                        } catch {
                            # leave as $null if metrics fail
                            $allocatedVal = $null
                            $usedVal = $null
                        }

                        # Normalize numeric values to doubles for calculations
                        $allocatedBytes = if ($allocatedVal -ne $null) { [double]$allocatedVal } else { $null }
                        $usedBytes = if ($usedVal -ne $null) { [double]$usedVal } else { $null }

                        $sqlObj.Add("Allocated_Bytes", $allocatedBytes)
                        $sqlObj.Add("Utilized_Bytes", $usedBytes)
                        if ($allocatedBytes) { $sqlObj.Add("Allocated_GB", [math]::Round($allocatedBytes / 1000000000, 3)) } else { $sqlObj.Add("Allocated_GB", $null) }
                        if ($usedBytes) { $sqlObj.Add("Utilized_GB", [math]::Round($usedBytes / 1000000000, 3)) } else { $sqlObj.Add("Utilized_GB", $null) }

                        # Compute percent used. Prefer allocated metric; if missing, fall back to MaxSizeBytes from the DB resource
                        $percentUsed = $null
                        if ($usedBytes -ne $null -and $allocatedBytes -ne $null -and $allocatedBytes -gt 0) {
                            $percentUsed = [math]::Round((($usedBytes / $allocatedBytes) * 100), 2)
                        } elseif ($usedBytes -ne $null -and $sqlDB.MaxSizeBytes -ne $null -and $sqlDB.MaxSizeBytes -gt 0) {
                            $percentUsed = [math]::Round((($usedBytes / [double]$sqlDB.MaxSizeBytes) * 100), 2)
                        }
                        $sqlObj.Add("PercentUsed", $percentUsed)

                        # STR / LTR policies (best-effort; master & some instance types may not support)
                        try {
                            $ltr = Get-AzSqlDatabaseBackupLongTermRetentionPolicy -ServerName $sqlDB.ServerName -DatabaseName $sqlDB.DatabaseName -ResourceGroupName $sqlDB.ResourceGroupName -ErrorAction SilentlyContinue
                        } catch { $ltr = $null }
                        try {
                            $str = Get-AzSqlDatabaseBackupShortTermRetentionPolicy -ServerName $sqlDB.ServerName -DatabaseName $sqlDB.DatabaseName -ResourceGroupName $sqlDB.ResourceGroupName -ErrorAction SilentlyContinue
                        } catch { $str = $null }
                        $sqlObj.Add("LTRWeeklyRetention", ($ltr.WeeklyRetention -as [string]))
                        $sqlObj.Add("LTRMonthlyRetention", ($ltr.MonthlyRetention -as [string]))
                        $sqlObj.Add("PITR_Days", ($str.RetentionDays -as [string]))

                        $SqlDbInventory += New-Object -TypeName PSObject -Property $sqlObj
                    }
                }
                Write-Progress -Id 8 -Activity "Processing SQL Servers" -Completed
            }
        } catch {
            Write-Warning "Error processing SQL servers/databases in subscription $($sub.Name): $($_.Exception.Message)"
        }

        # MySQL Servers (part of SQL inventory)
        try {
            Write-Host "Processing MySQL servers in subscription $($sub.Name)" -ForegroundColor Green
            
            $allMySQLServers = @()
            
            try {
                # Get all MySQL flexible servers at subscription level
                Write-Progress -Id 10 -ParentId 1 -Activity "Discovering MySQL Servers" -Status "Getting MySQL Flexible Servers..." -PercentComplete 25
                $mysqlFlexibleServers = Get-AzMySqlFlexibleServer -ErrorAction SilentlyContinue
                if ($mysqlFlexibleServers) {
                    $allMySQLServers += $mysqlFlexibleServers | ForEach-Object {
                        $_ | Add-Member -NotePropertyName "ServerType" -NotePropertyValue "Flexible" -PassThru |
                             Add-Member -NotePropertyName "ResourceGroupName" -NotePropertyValue ($_.Id -split '/')[4] -PassThru
                    }
                }
                
                # Get all MySQL single servers at subscription level (deprecated but may still exist)
                Write-Progress -Id 10 -ParentId 1 -Activity "Discovering MySQL Servers" -Status "Getting MySQL Single Servers..." -PercentComplete 50
                $mysqlSingleServers = Get-AzMySqlServer -ErrorAction SilentlyContinue
                if ($mysqlSingleServers) {
                    $allMySQLServers += $mysqlSingleServers | ForEach-Object {
                        $_ | Add-Member -NotePropertyName "ServerType" -NotePropertyValue "Single" -PassThru |
                             Add-Member -NotePropertyName "ResourceGroupName" -NotePropertyValue ($_.Id -split '/')[4] -PassThru
                    }
                }
            } catch {
                Write-Warning "Unable to collect MySQL information from subscription: $($sub.Name). Error: $($_.Exception.Message)"
            }
            
            Write-Progress -Id 10 -Activity "Discovering MySQL Servers" -Status "Discovery Complete" -PercentComplete 100
            Start-Sleep -Milliseconds 100  # Brief pause to ensure progress bar displays properly
            Write-Progress -Id 10 -Activity "Discovering MySQL Servers" -Completed
            
            if ($allMySQLServers -and $allMySQLServers.Count -gt 0) {
                $mysqlCount = 0
                foreach ($mysqlServer in $allMySQLServers) {
                        $mysqlCount++
                        $mysqlPercentComplete = [math]::Round(($mysqlCount / $allMySQLServers.Count) * 100, 1)
                        Write-Progress -Id 11 -ParentId 1 -Activity "Processing MySQL Servers" -Status "Processing MySQL $mysqlCount of $($allMySQLServers.Count) - $mysqlPercentComplete% complete" -PercentComplete $mysqlPercentComplete
                        
                        try {
                            # Build MySQL object
                            $mysqlObj = [ordered]@{
                                "Id" = $mysqlServer.Id
                                "Subscription" = $sub.Name
                                "ResourceGroupName" = $mysqlServer.ResourceGroupName
                                "Name" = $mysqlServer.Name
                                "ServerType" = $mysqlServer.ServerType
                                "Location" = $mysqlServer.Location
                                "Region" = $mysqlServer.Location
                                "Version" = $mysqlServer.Version
                                "FullyQualifiedDomainName" = $mysqlServer.FullyQualifiedDomainName
                                "BackupRetentionDays" = $mysqlServer.BackupRetentionDays
                                "StorageSku" = $mysqlServer.StorageSku
                                "SkuName" = $mysqlServer.SkuName
                            }
                            
                            # Get storage size based on server type (different properties for different types)
                            if ($mysqlServer.ServerType -eq "Flexible") {
                                # Flexible servers typically have StorageInMb or StorageGB properties
                                if ($mysqlServer.StorageSizeGB) {
                                    $storageGB = $mysqlServer.StorageSizeGB
                                    $storageMB = [math]::Round($storageGB * 1024, 2)
                                    $mysqlObj.Add("StorageMB", $storageMB)
                                    $mysqlObj.Add("StorageGB", $storageGB)
                                } elseif ($mysqlServer.StorageInMb) {
                                    $storageMB = $mysqlServer.StorageInMb
                                    $storageGB = [math]::Round($storageMB / 1024, 2)
                                    $mysqlObj.Add("StorageMB", $storageMB)
                                    $mysqlObj.Add("StorageGB", $storageGB)
                                } else {
                                    $mysqlObj.Add("StorageMB", $null)
                                    $mysqlObj.Add("StorageGB", $null)
                                }
                            } elseif ($mysqlServer.ServerType -eq "Single") {
                                # Single servers have StorageProfile.StorageMb structure
                                if ($mysqlServer.StorageProfile -and $mysqlServer.StorageProfile.StorageMb) {
                                    $storageMB = $mysqlServer.StorageProfile.StorageMb
                                    $storageGB = [math]::Round($storageMB / 1024, 2)
                                    $mysqlObj.Add("StorageMB", $storageMB)
                                    $mysqlObj.Add("StorageGB", $storageGB)
                                } elseif ($mysqlServer.StorageSizeGB) {
                                    # Fallback to direct StorageGB if StorageProfile is not available
                                    $storageGB = $mysqlServer.StorageSizeGB
                                    $storageMB = [math]::Round($storageGB * 1024, 2)
                                    $mysqlObj.Add("StorageMB", $storageMB)
                                    $mysqlObj.Add("StorageGB", $storageGB)
                                } else {
                                    $mysqlObj.Add("StorageMB", $null)
                                    $mysqlObj.Add("StorageGB", $null)
                                }
                            } else {
                                # Unknown server type - try common properties
                                if ($mysqlServer.StorageGB) {
                                    $storageGB = $mysqlServer.StorageGB
                                    $storageMB = [math]::Round($storageGB * 1024, 2)
                                    $mysqlObj.Add("StorageMB", $storageMB)
                                    $mysqlObj.Add("StorageGB", $storageGB)
                                } else {
                                    $mysqlObj.Add("StorageMB", $null)
                                    $mysqlObj.Add("StorageGB", $null)
                                }
                            }
                            
                            # Get metrics for storage usage
                            $id = $mysqlServer.Id
                            
                            # Storage Used
                            try {
                                $storageUsedMetric = (Get-AzMetric -WarningAction SilentlyContinue -ResourceId $id -MetricName storage_used -AggregationType Maximum -StartTime (Get-Date).AddHours(-1)).Data.Maximum | Select-Object -Last 1
                                if ($storageUsedMetric) {
                                    $storageUsedMB = [math]::Round($storageUsedMetric / (1024 * 1024), 2)  # Convert bytes to MB
                                    $storageUsedGB = [math]::Round($storageUsedMetric / (1024 * 1024 * 1024), 4)  # Convert bytes to GB
                                } else {
                                    $storageUsedMB = $null
                                    $storageUsedGB = $null
                                }
                                $mysqlObj.Add("StorageUsedBytes", $storageUsedMetric)
                                $mysqlObj.Add("StorageUsedMB", $storageUsedMB)
                                $mysqlObj.Add("StorageUsedGB", $storageUsedGB)
                            } catch {
                                $mysqlObj.Add("StorageUsedBytes", $null)
                                $mysqlObj.Add("StorageUsedMB", $null)
                                $mysqlObj.Add("StorageUsedGB", $null)
                            }
                            
                            # Storage Percent
                            try {
                                $storagePercentMetric = (Get-AzMetric -WarningAction SilentlyContinue -ResourceId $id -MetricName storage_percent -AggregationType Maximum -StartTime (Get-Date).AddHours(-1)).Data.Maximum | Select-Object -Last 1
                                $mysqlObj.Add("StoragePercent", $storagePercentMetric)
                            } catch {
                                $mysqlObj.Add("StoragePercent", $null)
                            }

                            
                            $MySQLServers += New-Object -TypeName PSObject -Property $mysqlObj
                            Write-Verbose "Added MySQL server $($mysqlServer.Name)"
                        } catch {
                            Write-Warning "Error processing MySQL server $($mysqlServer.Name): $($_.Exception.Message)"
                            
                            # Create minimal object on error
                            $mysqlObj = [ordered]@{
                                "Id" = $mysqlServer.Id
                                "Name" = $mysqlServer.Name
                                "ServerType" = $mysqlServer.ServerType
                                "Error" = $_.Exception.Message
                            }
                            $MySQLServers += New-Object -TypeName PSObject -Property $mysqlObj
                    }
                }
                Write-Progress -Id 11 -ParentId 1 -Activity "Processing MySQL Servers" -Completed
            } else {
                Write-Host "No MySQL servers found in subscription $($sub.Name)" -ForegroundColor Yellow
            }
            
        } catch {
            Write-Warning "Error processing MySQL servers in subscription $($sub.Name): $($_.Exception.Message)"
        }

        # PostgreSQL Servers (part of SQL inventory)
        try {
            Write-Host "Processing PostgreSQL servers in subscription $($sub.Name)" -ForegroundColor Green
            
            $allPostgreSQLServers = @()
            
            try {
                # Get all PostgreSQL flexible servers at subscription level
                Write-Progress -Id 12 -ParentId 1 -Activity "Discovering PostgreSQL Servers" -Status "Getting PostgreSQL Flexible Servers..." -PercentComplete 25
                $postgresFlexibleServers = Get-AzPostgreSqlFlexibleServer -ErrorAction SilentlyContinue
                if ($postgresFlexibleServers) {
                    $allPostgreSQLServers += $postgresFlexibleServers | ForEach-Object {
                        $_ | Add-Member -NotePropertyName "ServerType" -NotePropertyValue "Flexible" -PassThru |
                             Add-Member -NotePropertyName "ResourceGroupName" -NotePropertyValue ($_.Id -split '/')[4] -PassThru
                    }
                }
                
                # Get all PostgreSQL single servers at subscription level (deprecated but may still exist)
                Write-Progress -Id 12 -ParentId 1 -Activity "Discovering PostgreSQL Servers" -Status "Getting PostgreSQL Single Servers..." -PercentComplete 50
                $postgresSingleServers = Get-AzPostgreSqlServer -ErrorAction SilentlyContinue
                if ($postgresSingleServers) {
                    $allPostgreSQLServers += $postgresSingleServers | ForEach-Object {
                        $_ | Add-Member -NotePropertyName "ServerType" -NotePropertyValue "Single" -PassThru |
                             Add-Member -NotePropertyName "ResourceGroupName" -NotePropertyValue ($_.Id -split '/')[4] -PassThru
                    }
                }
            } catch {
                Write-Warning "Unable to collect PostgreSQL information from subscription: $($sub.Name). Error: $($_.Exception.Message)"
            }
            
            Write-Progress -Id 12 -Activity "Discovering PostgreSQL Servers" -Status "Discovery Complete" -PercentComplete 100
            Start-Sleep -Milliseconds 100  # Brief pause to ensure progress bar displays properly
            Write-Progress -Id 12 -Activity "Discovering PostgreSQL Servers" -Completed
            
            if ($allPostgreSQLServers -and $allPostgreSQLServers.Count -gt 0) {
                $postgresCount = 0
                foreach ($postgresServer in $allPostgreSQLServers) {
                        $postgresCount++
                        $postgresPercentComplete = [math]::Round(($postgresCount / $allPostgreSQLServers.Count) * 100, 1)
                        Write-Progress -Id 13 -ParentId 1 -Activity "Processing PostgreSQL Servers" -Status "Processing PostgreSQL $postgresCount of $($allPostgreSQLServers.Count) - $postgresPercentComplete% complete" -PercentComplete $postgresPercentComplete
                        
                        try {
                            # Build PostgreSQL object
                            $postgresObj = [ordered]@{
                                "Subscription" = $sub.Name
                                "ResourceGroupName" = $postgresServer.ResourceGroupName
                                "Name" = $postgresServer.Name
                                "Id" = $postgresServer.Id
                                "ServerType" = $postgresServer.ServerType
                                "Location" = $postgresServer.Location
                                "Region" = $postgresServer.Location
                                "Version" = $postgresServer.Version
                                "FullyQualifiedDomainName" = $postgresServer.FullyQualifiedDomainName
                                "BackupRetentionDays" = if ($postgresServer.BackupRetentionDays) { $postgresServer.BackupRetentionDays } elseif ($postgresServer.BackupRetentionDay) { $postgresServer.BackupRetentionDay } elseif ($postgresServer.StorageProfileBackupRetentionDay) { $postgresServer.StorageProfileBackupRetentionDay } else { $null }
                                "SkuName" = $postgresServer.SkuName
                            }
                            
                            # Get storage size based on server type (different properties for different types)
                            if ($postgresServer.ServerType -eq "Flexible") {
                                # Flexible servers have different storage properties
                                if ($postgresServer.StorageSizeGB) {
                                    $storageGB = $postgresServer.StorageSizeGB
                                    $storageMB = $storageGB * 1024
                                    $postgresObj.Add("StorageMB", $storageMB)
                                    $postgresObj.Add("StorageGB", $storageGB)
                                } elseif ($postgresServer.StorageInMb) {
                                    $storageMB = $postgresServer.StorageInMb
                                    $storageGB = [math]::Round($storageMB / 1024, 2)
                                    $postgresObj.Add("StorageMB", $storageMB)
                                    $postgresObj.Add("StorageGB", $storageGB)
                                } else {
                                    $postgresObj.Add("StorageMB", $null)
                                    $postgresObj.Add("StorageGB", $null)
                                }
                            } elseif ($postgresServer.ServerType -eq "Single") {
                                # Single servers typically have StorageInMb property
                                if ($postgresServer.StorageInMb) {
                                    $storageMB = $postgresServer.StorageInMb
                                    $storageGB = [math]::Round($storageMB / 1024, 2)
                                    $postgresObj.Add("StorageMB", $storageMB)
                                    $postgresObj.Add("StorageGB", $storageGB)
                                } else {
                                    $postgresObj.Add("StorageMB", $null)
                                    $postgresObj.Add("StorageGB", $null)
                                }
                            } else {
                                $postgresObj.Add("StorageMB", $null)
                                $postgresObj.Add("StorageGB", $null)
                            }
                            
                            # Get metrics for storage usage
                            $id = $postgresServer.Id
                            
                            # Storage Used
                            try {
                                $storageUsedMetric = (Get-AzMetric -WarningAction SilentlyContinue -ResourceId $id -MetricName storage_used -AggregationType Maximum -StartTime (Get-Date).AddHours(-1)).Data.Maximum | Select-Object -Last 1
                                $storageUsedGB = if ($storageUsedMetric -ne $null) { [math]::Round($storageUsedMetric / 1000000000, 4) } else { $null }
                                $postgresObj.Add("StorageUsedGB", $storageUsedGB)
                            } catch {
                                $postgresObj.Add("StorageUsedGB", $null)
                            }
                            
                            # Storage Percent
                            try {
                                $storagePercentMetric = (Get-AzMetric -WarningAction SilentlyContinue -ResourceId $id -MetricName storage_percent -AggregationType Maximum -StartTime (Get-Date).AddHours(-1)).Data.Maximum | Select-Object -Last 1
                                $postgresObj.Add("StoragePercent", $storagePercentMetric)
                            } catch {
                                $postgresObj.Add("StoragePercent", $null)
                            }
                            
                            $PostgreSQLServers += New-Object -TypeName PSObject -Property $postgresObj
                            Write-Verbose "Added PostgreSQL server $($postgresServer.Name)"
                        } catch {
                            Write-Warning "Error processing PostgreSQL server $($postgresServer.Name): $($_.Exception.Message)"
                            
                            # Create minimal object on error
                            $postgresObj = [ordered]@{
                                "Id" = $postgresServer.Id
                                "Name" = $postgresServer.Name
                                "ServerType" = $postgresServer.ServerType
                                "Error" = $_.Exception.Message
                            }
                            $PostgreSQLServers += New-Object -TypeName PSObject -Property $postgresObj
                    }
                }
                Write-Progress -Id 13 -ParentId 1 -Activity "Processing PostgreSQL Servers" -Completed
            } else {
                Write-Host "No PostgreSQL servers found in subscription $($sub.Name)" -ForegroundColor Yellow
            }
        } catch {
            Write-Warning "Error processing PostgreSQL servers in subscription $($sub.Name): $($_.Exception.Message)"
        }
    }
    
    # CosmosDB Accounts
    if ($Selected.COSMOS) {
        try {
            Write-Host "Processing CosmosDB accounts in subscription $($sub.Name)" -ForegroundColor Green
            
            # Get all resource groups in the subscription
            $resourceGroups = Get-AzResourceGroup -ErrorAction SilentlyContinue
            if (-not $resourceGroups) {
                Write-Host "No resource groups found in subscription $($sub.Name)" -ForegroundColor Yellow
                continue
            }
            
            $allCosmosAccounts = @()
            $rgCount = 0
            
            # Iterate through each resource group to find CosmosDB accounts
            foreach ($rg in $resourceGroups) {
                $rgCount++
                Write-Progress -Id 8 -ParentId 1 -Activity "Scanning Resource Groups for CosmosDB" -Status "Resource Group $rgCount of $($resourceGroups.Count): $($rg.ResourceGroupName)" -PercentComplete ([math]::Round(($rgCount / $resourceGroups.Count) * 100, 1))
                
                try {
                    $cosmosAccountsInRG = Get-AzCosmosDBAccount -ResourceGroupName $rg.ResourceGroupName -ErrorAction SilentlyContinue
                    if ($cosmosAccountsInRG) {
                        $allCosmosAccounts += $cosmosAccountsInRG | Add-Member -NotePropertyName "ResourceGroupName" -NotePropertyValue $rg.ResourceGroupName -PassThru
                    }
                } catch {
                    Write-Error "Unable to collect CosmosDB information for Resource Group $($rg.ResourceGroupName) in subscription: $($sub.Name)"
                    Write-Error "Error: $_"
                    Continue
                }
            }
            
            Write-Progress -Id 8 -Activity "Scanning Resource Groups for CosmosDB" -Completed
            
            if ($allCosmosAccounts -and $allCosmosAccounts.Count -gt 0) {
                $cosmosCount = 0
                foreach ($cosmosAccount in $allCosmosAccounts) {
                    $cosmosCount++
                    $cosmosPercentComplete = [math]::Round(($cosmosCount / $allCosmosAccounts.Count) * 100, 1)
                    Write-Progress -Id 9 -ParentId 1 -Activity "Processing CosmosDB Accounts" -Status "Processing CosmosDB $cosmosCount of $($allCosmosAccounts.Count) - $cosmosPercentComplete% complete" -PercentComplete $cosmosPercentComplete
                    
                    try {
                        # Build CosmosDB object following the reference file pattern
                        $cosmosObj = [ordered]@{
                            "Subscription" = $sub.Name
                            "ResourceGroupName" = $cosmosAccount.ResourceGroupName
                            "Name" = $cosmosAccount.Name
                            "Location" = $cosmosAccount.Location
                            "Id" = $cosmosAccount.Id
                            "Kind" = $cosmosAccount.Kind
                            "InstanceId" = $cosmosAccount.InstanceId
                            "BackupPolicyBackupIntervalInMinutes" = $cosmosAccount.BackupPolicy.BackupIntervalInMinutes
                            "BackupPolicyBackupRetentionIntervalInHours" = $cosmosAccount.BackupPolicy.BackupRetentionIntervalInHours
                            "BackupPolicyBackupType" = $cosmosAccount.BackupPolicy.BackupType
                            "BackupPolicyBackupStorageRedundancy" = $cosmosAccount.BackupPolicy.BackupStorageRedundancy
                            "MinimalTlsVersion" = $cosmosAccount.MinimalTlsVersion
                        }
                        
                        # Get metrics following the reference file pattern
                        $id = $cosmosAccount.Id
                        
                        # Document Count
                        try {
                            $documentCountMetric = (Get-AzMetric -WarningAction SilentlyContinue -ResourceId $id -MetricName DocumentCount -AggregationType Maximum -StartTime (Get-Date).AddHours(-1)).Data.Maximum | Select-Object -Last 1
                            $cosmosObj.Add("DocumentCount", $documentCountMetric)
                        } catch {
                            $cosmosObj.Add("DocumentCount", $null)
                        }
                        
                        # Data Usage
                        try {
                            $dataUsageMetric = (Get-AzMetric -WarningAction SilentlyContinue -ResourceId $id -MetricName DataUsage -AggregationType Maximum -StartTime (Get-Date).AddHours(-1)).Data.Maximum | Select-Object -Last 1
                            $cosmosObj.Add("DataUsage", $dataUsageMetric)
                            # Convert DataUsage from bytes to GB
                            $dataUsageGB = if ($dataUsageMetric -ne $null -and $dataUsageMetric -ne '') { [math]::Round($dataUsageMetric / 1000000000, 4) } else { $null }
                            $cosmosObj.Add("DataUsageGB", $dataUsageGB)
                        } catch {
                            $cosmosObj.Add("DataUsage", $null)
                            $cosmosObj.Add("DataUsageGB", $null)
                        }
                        
                        # Physical Partition Size Info
                        try {
                            $partitionSizeMetric = (Get-AzMetric -WarningAction SilentlyContinue -ResourceId $id -MetricName PhysicalPartitionSizeInfo -AggregationType Maximum -StartTime (Get-Date).AddHours(-1)).Data.Maximum | Select-Object -Last 1
                            $cosmosObj.Add("PhysicalPartitionSizeInfo", $partitionSizeMetric)
                        } catch {
                            $cosmosObj.Add("PhysicalPartitionSizeInfo", $null)
                        }
                        
                        # Physical Partition Count
                        try {
                            $partitionCountMetric = (Get-AzMetric -WarningAction SilentlyContinue -ResourceId $id -MetricName PhysicalPartitionCount -AggregationType Maximum -StartTime (Get-Date).AddHours(-1)).Data.Maximum | Select-Object -Last 1
                            $cosmosObj.Add("PhysicalPartitionCount", $partitionCountMetric)
                        } catch {
                            $cosmosObj.Add("PhysicalPartitionCount", $null)
                        }
                        
                        # Index Usage
                        try {
                            $indexUsageMetric = (Get-AzMetric -WarningAction SilentlyContinue -ResourceId $id -MetricName IndexUsage -AggregationType Maximum -StartTime (Get-Date).AddHours(-1)).Data.Maximum | Select-Object -Last 1
                            $cosmosObj.Add("IndexUsage", $indexUsageMetric)
                        } catch {
                            $cosmosObj.Add("IndexUsage", $null)
                        }
                        
                        $CosmosDBs += New-Object -TypeName PSObject -Property $cosmosObj
                        Write-Verbose "Added CosmosDB account $($cosmosAccount.Name)"
                    } catch {
                        Write-Warning "Error processing CosmosDB account $($cosmosAccount.Name): $($_.Exception.Message)"
                        
                        # Create minimal object on error
                        $cosmosObj = [ordered]@{
                            "Subscription" = $sub.Name
                            "ResourceGroupName" = $cosmosAccount.ResourceGroupName
                            "Name" = $cosmosAccount.Name
                            "Location" = $cosmosAccount.Location
                            "Id" = $cosmosAccount.Id
                            "Kind" = $cosmosAccount.Kind
                            "Error" = $_.Exception.Message
                        }
                        $CosmosDBs += New-Object -TypeName PSObject -Property $cosmosObj
                    }
                }
                Write-Progress -Id 9 -ParentId 1 -Activity "Processing CosmosDB Accounts" -Completed
            } else {
                Write-Host "No CosmosDB accounts found in subscription $($sub.Name)" -ForegroundColor Yellow
            }
        } catch {
            Write-Warning "Error processing CosmosDB accounts in subscription $($sub.Name): $($_.Exception.Message)"
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
if ($Selected.SQL) { Write-Host "Total SQL Managed Instances found: $($SqlInstancesInventory.Count)" -ForegroundColor Cyan }
if ($Selected.SQL) { Write-Host "Total MySQL Servers found: $($MySQLServers.Count)" -ForegroundColor Cyan }
if ($Selected.SQL) { Write-Host "Total PostgreSQL Servers found: $($PostgreSQLServers.Count)" -ForegroundColor Cyan }
if ($Selected.SQL) { Write-Host "Total Azure SQL Servers found: $($SqlDbInventory | Select-Object -ExpandProperty Server | Sort-Object -Unique | Measure-Object).Count" -ForegroundColor Cyan }
if ($Selected.COSMOS) { Write-Host "Total CosmosDB Accounts found: $($CosmosDBs.Count)" -ForegroundColor Cyan }

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
if ($Selected.SQL -and $SqlInstancesInventory.Count) { 
    Write-Progress -Id 5 -Activity "Generating Output Files" -Status "Writing SQL Instances CSV..." -PercentComplete 75
    $SqlInstancesInventory | Export-Csv (Join-Path $outdir "azure_sql_managed_instances_$dateStr.csv") -NoTypeInformation 
    Write-Host "azure_sql_managed_instances_$dateStr.csv file has been written to $outdir" -ForegroundColor Cyan
}
if ($Selected.SQL -and $SqlDbInventory.Count) {
    Write-Progress -Id 5 -Activity "Generating Output Files" -Status "Writing SQL Databases CSV..." -PercentComplete 80
    $SqlDbInventory | Export-Csv (Join-Path $outdir "azure_sql_databases_inventory_$dateStr.csv") -NoTypeInformation
    Write-Host "azure_sql_databases_inventory_$dateStr.csv file has been written to $outdir" -ForegroundColor Cyan
}
if ($Selected.SQL -and $SqlMIDbInventory.Count) {
    Write-Progress -Id 5 -Activity "Generating Output Files" -Status "Writing SQL MI Databases CSV..." -PercentComplete 82
    $SqlMIDbInventory | Export-Csv (Join-Path $outdir "azure_sql_mi_databases_$dateStr.csv") -NoTypeInformation
    Write-Host "azure_sql_mi_databases_$dateStr.csv file has been written to $outdir" -ForegroundColor Cyan
}
if ($Selected.COSMOS -and $CosmosDBs.Count) {
    Write-Progress -Id 5 -Activity "Generating Output Files" -Status "Writing CosmosDB CSV..." -PercentComplete 85
    $CosmosDBs | Export-Csv (Join-Path $outdir "azure_cosmosdb_accounts_$dateStr.csv") -NoTypeInformation
    Write-Host "azure_cosmosdb_accounts_$dateStr.csv file has been written to $outdir" -ForegroundColor Cyan
}
if ($Selected.SQL -and $MySQLServers.Count) {
    Write-Progress -Id 5 -Activity "Generating Output Files" -Status "Writing MySQL CSV..." -PercentComplete 90
    $MySQLServers | Export-Csv (Join-Path $outdir "azure_mysql_servers_$dateStr.csv") -NoTypeInformation
    Write-Host "azure_mysql_servers_$dateStr.csv file has been written to $outdir" -ForegroundColor Cyan
}
if ($Selected.SQL -and $PostgreSQLServers.Count) {
    Write-Progress -Id 5 -Activity "Generating Output Files" -Status "Writing PostgreSQL CSV..." -PercentComplete 92
    $PostgreSQLServers | Export-Csv (Join-Path $outdir "azure_postgresql_servers_$dateStr.csv") -NoTypeInformation
    Write-Host "azure_postgresql_servers_$dateStr.csv file has been written to $outdir" -ForegroundColor Cyan
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
        } elseif ($k -eq "SQL") {
            # Handle SQL Managed Instances and SQL Databases separately
            
            # Add SQL Managed Instances if any exist
            if ($SqlInstancesInventory.Count -gt 0) {
                $ResourceType = "SQL Managed Instances"
                $count = $SqlInstancesInventory.Count
                $totalStorageUsed = ($SqlInstancesInventory | Measure-Object -Property StorageUsedGB -Sum).Sum
                if ($totalStorageUsed -eq $null) { $totalStorageUsed = 0 }
                $totalSize = [math]::Round($totalStorageUsed, 2)
                $totalSizeTB = [math]::Round($totalSize / 1000, 4)
                $totalSizeTiB = [math]::Round($totalSize / 1024, 4)
                
                # Add SQL Managed Instances summary row
                $summaryRows += [PSCustomObject]@{ 
                    Subscription = "All"
                    ResourceType = $ResourceType
                    Region = "All"
                    Count = $count
                    TotalSizeGB = $totalSize
                    TotalSizeTB = $totalSizeTB
                    TotalSizeTiB = $totalSizeTiB
                }
            }
            
            # Add SQL Databases if any exist
            if ($SqlDbInventory.Count -gt 0) {
                $ResourceType = "SQL Databases"
                $count = $SqlDbInventory.Count
                $totalDbMaxSize = ($SqlDbInventory | Measure-Object -Property MaxSizeGB -Sum).Sum
                if ($totalDbMaxSize -eq $null) { $totalDbMaxSize = 0 }
                $totalSize = [math]::Round($totalDbMaxSize, 2)
                $totalSizeTB = [math]::Round($totalSize / 1000, 4)
                $totalSizeTiB = [math]::Round($totalSize / 1024, 4)
                
                # Add SQL Databases summary row
                $summaryRows += [PSCustomObject]@{ 
                    Subscription = "All"
                    ResourceType = $ResourceType
                    Region = "All"
                    Count = $count
                    TotalSizeGB = $totalSize
                    TotalSizeTB = $totalSizeTB
                    TotalSizeTiB = $totalSizeTiB
                }
            }
            
            # Add MySQL Servers if any exist
            if ($MySQLServers.Count -gt 0) {
                $ResourceType = "MySQL Servers"
                $count = $MySQLServers.Count
                # Calculate total storage from StorageGB field
                $totalStorageGB = 0
                foreach ($mysql in $MySQLServers) {
                    if ($mysql.StorageUsedGB -ne $null -and $mysql.StorageUsedGB -ne '') {
                        $totalStorageGB += [double]$mysql.StorageUsedGB
                    }
                }
                $totalSize = [math]::Round($totalStorageGB, 2)
                $totalSizeTB = [math]::Round($totalSize / 1000, 4)
                $totalSizeTiB = [math]::Round($totalSize / 1024, 4)
                
                # Add MySQL Servers summary row
                $summaryRows += [PSCustomObject]@{ 
                    Subscription = "All"
                    ResourceType = $ResourceType
                    Region = "All"
                    Count = $count
                    TotalSizeGB = $totalSize
                    TotalSizeTB = $totalSizeTB
                    TotalSizeTiB = $totalSizeTiB
                }
            }
            
            # Add PostgreSQL Servers if any exist
            if ($PostgreSQLServers.Count -gt 0) {
                $ResourceType = "PostgreSQL Servers"
                $count = $PostgreSQLServers.Count
                # Calculate total storage from StorageGB field
                $totalStorageGB = 0
                foreach ($postgres in $PostgreSQLServers) {
                    if ($postgres.StorageUsedGB -ne $null -and $postgres.StorageUsedGB -ne '') {
                        $totalStorageGB += [double]$postgres.StorageUsedGB
                    }
                }
                $totalSize = [math]::Round($totalStorageGB, 2)
                $totalSizeTB = [math]::Round($totalSize / 1000, 4)
                $totalSizeTiB = [math]::Round($totalSize / 1024, 4)
                
                # Add PostgreSQL Servers summary row
                $summaryRows += [PSCustomObject]@{ 
                    Subscription = "All"
                    ResourceType = $ResourceType
                    Region = "All"
                    Count = $count
                    TotalSizeGB = $totalSize
                    TotalSizeTB = $totalSizeTB
                    TotalSizeTiB = $totalSizeTiB
                }
            }
            
            # Set count to 0 to prevent the normal summary row addition at the end
            $count = 0
        } elseif ($k -eq "COSMOS" -and $CosmosDBs.Count -gt 0) {
            $totalDatabases = ($CosmosDBs | Measure-Object -Property DatabaseCount -Sum).Sum
            if ($totalDatabases -eq $null) { $totalDatabases = 0 }
            $totalContainers = ($CosmosDBs | Measure-Object -Property ContainerCount -Sum).Sum
            if ($totalContainers -eq $null) { $totalContainers = 0 }
            $ResourceType = "CosmosDB"
            $count = $CosmosDBs.Count
            # Calculate total from DataUsage field (in bytes)
            $totalDataUsage = 0
            foreach ($cosmos in $CosmosDBs) {
                if ($cosmos.DataUsage -ne $null -and $cosmos.DataUsage -ne '') {
                    $totalDataUsage += [double]$cosmos.DataUsage
                }
            }
            $totalCapacityGB = if ($totalDataUsage -gt 0) { [math]::Round($totalDataUsage / 1000000000, 2) } else { 0 }
            $totalSize = $totalCapacityGB
            $totalSizeTB = if ($totalDataUsage -gt 0) { [math]::Round($totalDataUsage / 1000000000000, 4) } else { 0 }
            $totalSizeTiB = if ($totalDataUsage -gt 0) { [math]::Round($totalDataUsage / 1099511627776, 4) } else { 0 }
            $totalSizeTiB = if ($totalDataUsage -gt 0) { [math]::Round($totalDataUsage / 1099511627776, 4) } else { 0 }
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

# Add SQL Managed Instances regional breakdown if selected
if ($SqlInstancesInventory.Count -and $Selected.SQL) {
    $sqlMIRegionalSummary = $SqlInstancesInventory | Group-Object Region | ForEach-Object {
        $totalStorageUsed = ($_.Group | Measure-Object -Property StorageUsedGB -Sum).Sum
        if ($totalStorageUsed -eq $null) { $totalStorageUsed = 0 }
        $totalStorageTB = [math]::Round($totalStorageUsed / 1000, 4)
        $totalStorageTiB = [math]::Round($totalStorageUsed / 1024, 4)
        [PSCustomObject]@{
            Subscription = "All"
            ResourceType = "SQL Managed Instances"
            Region = $_.Name
            Count = $_.Count
            TotalSizeGB = [math]::Round($totalStorageUsed, 2)
            TotalSizeTB = $totalStorageTB
            TotalSizeTiB = $totalStorageTiB
        }
    } | Sort-Object Region

    $summaryRows += $sqlMIRegionalSummary
}

# Add SQL Databases regional breakdown if selected
if ($SqlDbInventory.Count -and $Selected.SQL) {
    $sqlDBRegionalSummary = $SqlDbInventory | Group-Object Region | ForEach-Object {
        $totalMaxSize = ($_.Group | Measure-Object -Property MaxSizeGB -Sum).Sum
        if ($totalMaxSize -eq $null) { $totalMaxSize = 0 }
        $totalSizeTB = [math]::Round($totalMaxSize / 1000, 4)
        $totalSizeTiB = [math]::Round($totalMaxSize / 1024, 4)
        [PSCustomObject]@{
            Subscription = "All"
            ResourceType = "SQL Databases"
            Region = $_.Name
            Count = $_.Count
            TotalSizeGB = [math]::Round($totalMaxSize, 2)
            TotalSizeTB = $totalSizeTB
            TotalSizeTiB = $totalSizeTiB
        }
    } | Sort-Object Region

    $summaryRows += $sqlDBRegionalSummary
}

# Add CosmosDB regional breakdown if selected
if ($CosmosDBs.Count -gt 0 -and $Selected.COSMOS) {
    $cosmosRegionalSummary = $CosmosDBs | Group-Object Location | ForEach-Object {
        $totalDataUsage = 0
        foreach ($cosmos in $_.Group) {
            if ($cosmos.DataUsage -ne $null -and $cosmos.DataUsage -ne '') {
                $totalDataUsage += [double]$cosmos.DataUsage
            }
        }
        $totalStorageGB = if ($totalDataUsage -gt 0) { [math]::Round($totalDataUsage / 1000000000, 4) } else { 0 }
        $totalStorageTB = if ($totalDataUsage -gt 0) { [math]::Round($totalDataUsage / 1000000000000, 4) } else { 0 }
        $totalStorageTiB = if ($totalDataUsage -gt 0) { [math]::Round($totalDataUsage / 1099511627776, 4) } else { 0 }
        [PSCustomObject]@{
            Subscription = "All"
            ResourceType = "CosmosDB"
            Region = $_.Name
            Count = $_.Count
            TotalSizeGB = $totalStorageGB
            TotalSizeTB = $totalStorageTB
            TotalSizeTiB = $totalStorageTiB
        }
    } | Sort-Object Region

    $summaryRows += $cosmosRegionalSummary
}

# Add MySQL regional breakdown if selected
if ($MySQLServers.Count -gt 0 -and $Selected.SQL) {
    $mysqlRegionalSummary = $MySQLServers | Group-Object Location | ForEach-Object {
        $totalStorageUsedGB = 0
        foreach ($mysql in $_.Group) {
            if ($mysql.StorageUsedGB -ne $null -and $mysql.StorageUsedGB -ne '') {
                $totalStorageUsedGB += [double]$mysql.StorageUsedGB
            }
        }
        $totalStorageTB = if ($totalStorageUsedGB -gt 0) { [math]::Round($totalStorageUsedGB / 1000, 4) } else { 0 }
        $totalStorageTiB = if ($totalStorageUsedGB -gt 0) { [math]::Round($totalStorageUsedGB / 1024, 4) } else { 0 }
        [PSCustomObject]@{
            Subscription = "All"
            ResourceType = "MySQL Servers"
            Region = $_.Name
            Count = $_.Count
            TotalSizeGB = [math]::Round($totalStorageUsedGB, 4)
            TotalSizeTB = $totalStorageTB
            TotalSizeTiB = $totalStorageTiB
        }
    } | Sort-Object Region

    $summaryRows += $mysqlRegionalSummary
}

# Add PostgreSQL regional breakdown if selected
if ($PostgreSQLServers.Count -gt 0 -and $Selected.SQL) {
    $postgresRegionalSummary = $PostgreSQLServers | Group-Object Location | ForEach-Object {
        $totalStorageUsedGB = 0
        foreach ($postgres in $_.Group) {
            if ($postgres.StorageUsedGB -ne $null -and $postgres.StorageUsedGB -ne '') {
                $totalStorageUsedGB += [double]$postgres.StorageUsedGB
            }
        }
        $totalStorageTB = if ($totalStorageUsedGB -gt 0) { [math]::Round($totalStorageUsedGB / 1000, 4) } else { 0 }
        $totalStorageTiB = if ($totalStorageUsedGB -gt 0) { [math]::Round($totalStorageUsedGB / 1024, 4) } else { 0 }
        [PSCustomObject]@{
            Subscription = "All"
            ResourceType = "PostgreSQL Servers"
            Region = $_.Name
            Count = $_.Count
            TotalSizeGB = [math]::Round($totalStorageUsedGB, 4)
            TotalSizeTB = $totalStorageTB
            TotalSizeTiB = $totalStorageTiB
        }
    } | Sort-Object Region

    $summaryRows += $postgresRegionalSummary
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
    
    # Process SQL for this subscription
    if ($Selected.SQL) {
        # SQL Managed Instances per-subscription summary
        $subscriptionSqlMI = $SqlInstancesInventory | Where-Object { $_.Subscription -eq $subscriptionName }
        if ($subscriptionSqlMI.Count -gt 0) {
            # Add SQL Managed Instances resource type total for this subscription
            $totalMIStorageUsed = ($subscriptionSqlMI | Measure-Object -Property StorageUsedGB -Sum).Sum
            if ($totalMIStorageUsed -eq $null) { $totalMIStorageUsed = 0 }
            $totalMIStorageTB = [math]::Round($totalMIStorageUsed / 1000, 4)
            $totalMIStorageTiB = [math]::Round($totalMIStorageUsed / 1024, 4)
            
            $summaryRows += [PSCustomObject]@{
                Subscription = $subscriptionName
                ResourceType = "SQL Managed Instances"
                Region = "All"
                Count = $subscriptionSqlMI.Count
                TotalSizeGB = [math]::Round($totalMIStorageUsed, 2)
                TotalSizeTB = $totalMIStorageTB
                TotalSizeTiB = $totalMIStorageTiB
            }
            
            # Add regional breakdown for SQL Managed Instances in this subscription
            $sqlMIRegionalBreakdown = $subscriptionSqlMI | Group-Object Region | ForEach-Object {
                $totalStorageUsed = ($_.Group | Measure-Object -Property StorageUsedGB -Sum).Sum
                if ($totalStorageUsed -eq $null) { $totalStorageUsed = 0 }
                $totalStorageTB = [math]::Round($totalStorageUsed / 1000, 4)
                $totalStorageTiB = [math]::Round($totalStorageUsed / 1024, 4)
                [PSCustomObject]@{
                    Subscription = $subscriptionName
                    ResourceType = "SQL Managed Instances"
                    Region = $_.Name
                    Count = $_.Count
                    TotalSizeGB = [math]::Round($totalStorageUsed, 2)
                    TotalSizeTB = $totalStorageTB
                    TotalSizeTiB = $totalStorageTiB
                }
            } | Sort-Object Region
            
            $summaryRows += $sqlMIRegionalBreakdown
        }
        
        # SQL Databases per-subscription summary
        $subscriptionSqlDB = $SqlDbInventory | Where-Object { $_.Subscription -eq $subscriptionName }
        if ($subscriptionSqlDB.Count -gt 0) {
            # Add SQL Databases resource type total for this subscription
            $totalDBMaxSize = ($subscriptionSqlDB | Measure-Object -Property MaxSizeGB -Sum).Sum
            if ($totalDBMaxSize -eq $null) { $totalDBMaxSize = 0 }
            $totalDBSizeTB = [math]::Round($totalDBMaxSize / 1000, 4)
            $totalDBSizeTiB = [math]::Round($totalDBMaxSize / 1024, 4)
            
            $summaryRows += [PSCustomObject]@{
                Subscription = $subscriptionName
                ResourceType = "SQL Databases"
                Region = "All"
                Count = $subscriptionSqlDB.Count
                TotalSizeGB = [math]::Round($totalDBMaxSize, 2)
                TotalSizeTB = $totalDBSizeTB
                TotalSizeTiB = $totalDBSizeTiB
            }
            
            # Add regional breakdown for SQL Databases in this subscription
            $sqlDBRegionalBreakdown = $subscriptionSqlDB | Group-Object Region | ForEach-Object {
                $totalMaxSize = ($_.Group | Measure-Object -Property MaxSizeGB -Sum).Sum
                if ($totalMaxSize -eq $null) { $totalMaxSize = 0 }
                $totalSizeTB = [math]::Round($totalMaxSize / 1000, 4)
                $totalSizeTiB = [math]::Round($totalMaxSize / 1024, 4)
                [PSCustomObject]@{
                    Subscription = $subscriptionName
                    ResourceType = "SQL Databases"
                    Region = $_.Name
                    Count = $_.Count
                    TotalSizeGB = [math]::Round($totalMaxSize, 2)
                    TotalSizeTB = $totalSizeTB
                    TotalSizeTiB = $totalSizeTiB
                }
            } | Sort-Object Region
            
            $summaryRows += $sqlDBRegionalBreakdown
        }
    }
    
    # CosmosDB per-subscription summary
    if ($Selected.COSMOS -and $CosmosDBs.Count -gt 0) {
        $subscriptionCosmosDBs = $CosmosDBs | Where-Object { $_.Subscription -eq $subscriptionName }
        
        if ($subscriptionCosmosDBs.Count -gt 0) {
            # Calculate totals
            $totalDataUsage = 0
            foreach ($cosmos in $subscriptionCosmosDBs) {
                if ($cosmos.DataUsage -ne $null -and $cosmos.DataUsage -ne '') {
                    $totalDataUsage += [double]$cosmos.DataUsage
                }
            }
            $totalStorageGB = if ($totalDataUsage -gt 0) { [math]::Round($totalDataUsage / 1000000000, 4) } else { 0 }
            $totalStorageTB = if ($totalDataUsage -gt 0) { [math]::Round($totalDataUsage / 1000000000000, 4) } else { 0 }
            $totalStorageTiB = if ($totalDataUsage -gt 0) { [math]::Round($totalDataUsage / 1099511627776, 4) } else { 0 }
            
            $summaryRows += [PSCustomObject]@{
                Subscription = $subscriptionName
                ResourceType = "CosmosDB"
                Region = "All"
                Count = $subscriptionCosmosDBs.Count
                TotalSizeGB = $totalStorageGB
                TotalSizeTB = $totalStorageTB
                TotalSizeTiB = $totalStorageTiB
            }
            
            # Add regional breakdown for CosmosDB in this subscription
            $cosmosRegionalBreakdown = $subscriptionCosmosDBs | Group-Object Location | ForEach-Object {
                $totalDataUsage = 0
                foreach ($cosmos in $_.Group) {
                    if ($cosmos.DataUsage -ne $null -and $cosmos.DataUsage -ne '') {
                        $totalDataUsage += [double]$cosmos.DataUsage
                    }
                }
                $totalStorageGB = if ($totalDataUsage -gt 0) { [math]::Round($totalDataUsage / 1000000000, 4) } else { 0 }
                $totalStorageTB = if ($totalDataUsage -gt 0) { [math]::Round($totalDataUsage / 1000000000000, 4) } else { 0 }
                $totalStorageTiB = if ($totalDataUsage -gt 0) { [math]::Round($totalDataUsage / 1099511627776, 4) } else { 0 }
                [PSCustomObject]@{
                    Subscription = $subscriptionName
                    ResourceType = "CosmosDB"
                    Region = $_.Name
                    Count = $_.Count
                    TotalSizeGB = $totalStorageGB
                    TotalSizeTB = $totalStorageTB
                    TotalSizeTiB = $totalStorageTiB
                }
            } | Sort-Object Region
            
            $summaryRows += $cosmosRegionalBreakdown
        }
    }
    
    # MySQL per-subscription summary
    if ($Selected.SQL -and $MySQLServers.Count -gt 0) {
        $subscriptionMySQLServers = $MySQLServers | Where-Object { $_.Subscription -eq $subscriptionName }
        
        if ($subscriptionMySQLServers.Count -gt 0) {
            # Calculate totals
            $totalStorageUsedGB = 0
            foreach ($mysql in $subscriptionMySQLServers) {
                if ($mysql.StorageUsedGB -ne $null -and $mysql.StorageUsedGB -ne '') {
                    $totalStorageUsedGB += [double]$mysql.StorageUsedGB
                }
            }
            $totalStorageTB = if ($totalStorageUsedGB -gt 0) { [math]::Round($totalStorageUsedGB / 1000, 4) } else { 0 }
            $totalStorageTiB = if ($totalStorageUsedGB -gt 0) { [math]::Round($totalStorageUsedGB / 1024, 4) } else { 0 }
            
            $summaryRows += [PSCustomObject]@{
                Subscription = $subscriptionName
                ResourceType = "MySQL Servers"
                Region = "All"
                Count = $subscriptionMySQLServers.Count
                TotalSizeGB = [math]::Round($totalStorageUsedGB, 4)
                TotalSizeTB = $totalStorageTB
                TotalSizeTiB = $totalStorageTiB
            }
            
            # Add regional breakdown for MySQL in this subscription
            $mysqlRegionalBreakdown = $subscriptionMySQLServers | Group-Object Location | ForEach-Object {
                $totalStorageUsedGB = 0
                foreach ($mysql in $_.Group) {
                    if ($mysql.StorageUsedGB -ne $null -and $mysql.StorageUsedGB -ne '') {
                        $totalStorageUsedGB += [double]$mysql.StorageUsedGB
                    }
                }
                $totalStorageTB = if ($totalStorageUsedGB -gt 0) { [math]::Round($totalStorageUsedGB / 1000, 4) } else { 0 }
                $totalStorageTiB = if ($totalStorageUsedGB -gt 0) { [math]::Round($totalStorageUsedGB / 1024, 4) } else { 0 }
                [PSCustomObject]@{
                    Subscription = $subscriptionName
                    ResourceType = "MySQL Servers"
                    Region = $_.Name
                    Count = $_.Count
                    TotalSizeGB = [math]::Round($totalStorageUsedGB, 4)
                    TotalSizeTB = $totalStorageTB
                    TotalSizeTiB = $totalStorageTiB
                }
            } | Sort-Object Region
            
            $summaryRows += $mysqlRegionalBreakdown
        }
    }
    
    # PostgreSQL per-subscription summary
    if ($Selected.SQL -and $PostgreSQLServers.Count -gt 0) {
        $subscriptionPostgreSQLServers = $PostgreSQLServers | Where-Object { $_.Subscription -eq $subscriptionName }
        
        if ($subscriptionPostgreSQLServers.Count -gt 0) {
            # Calculate totals
            $totalStorageUsedGB = 0
            foreach ($postgres in $subscriptionPostgreSQLServers) {
                if ($postgres.StorageUsedGB -ne $null -and $postgres.StorageUsedGB -ne '') {
                    $totalStorageUsedGB += [double]$postgres.StorageUsedGB
                }
            }
            $totalStorageTB = if ($totalStorageUsedGB -gt 0) { [math]::Round($totalStorageUsedGB / 1000, 4) } else { 0 }
            $totalStorageTiB = if ($totalStorageUsedGB -gt 0) { [math]::Round($totalStorageUsedGB / 1024, 4) } else { 0 }
            
            $summaryRows += [PSCustomObject]@{
                Subscription = $subscriptionName
                ResourceType = "PostgreSQL Servers"
                Region = "All"
                Count = $subscriptionPostgreSQLServers.Count
                TotalSizeGB = [math]::Round($totalStorageUsedGB, 4)
                TotalSizeTB = $totalStorageTB
                TotalSizeTiB = $totalStorageTiB
            }
            
            # Add regional breakdown for PostgreSQL in this subscription
            $postgresRegionalBreakdown = $subscriptionPostgreSQLServers | Group-Object Location | ForEach-Object {
                $totalStorageUsedGB = 0
                foreach ($postgres in $_.Group) {
                    if ($postgres.StorageUsedGB -ne $null -and $postgres.StorageUsedGB -ne '') {
                        $totalStorageUsedGB += [double]$postgres.StorageUsedGB
                    }
                }
                $totalStorageTB = if ($totalStorageUsedGB -gt 0) { [math]::Round($totalStorageUsedGB / 1000, 4) } else { 0 }
                $totalStorageTiB = if ($totalStorageUsedGB -gt 0) { [math]::Round($totalStorageUsedGB / 1024, 4) } else { 0 }
                [PSCustomObject]@{
                    Subscription = $subscriptionName
                    ResourceType = "PostgreSQL Servers"
                    Region = $_.Name
                    Count = $_.Count
                    TotalSizeGB = [math]::Round($totalStorageUsedGB, 4)
                    TotalSizeTB = $totalStorageTB
                    TotalSizeTiB = $totalStorageTiB
                }
            } | Sort-Object Region
            
            $summaryRows += $postgresRegionalBreakdown
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