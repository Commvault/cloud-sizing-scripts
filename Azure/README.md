### Azure - Execution Instructions

This PowerShell script inventories Azure resources across subscriptions to assist Commvault representatives in gathering information about cloud resources that may need protection and helps representatives in estimating the cost of protecting these resources.

#### Method 1 (Recommended) – Run in Azure Cloud Shell

1. Learn about Azure Cloud Shell:
   https://docs.microsoft.com/en-us/azure/cloud-shell/overview

2. Verify Azure permissions:
   Ensure your Azure AD account has "Reader" role on target subscriptions
   Additional "Reader and Data Access" role may be needed for storage metrics

3. Access Azure Cloud Shell:
   - Login to Azure Portal with verified account
   - Open Azure Cloud Shell (PowerShell mode)

4. Upload this script:
   Use the Cloud Shell file upload feature to upload `CVAzureCloudSizingScript.ps1`

5. Run the script (examples below). With no parameters it scans all accessible subscriptions and all supported resource types.

#### Method 2 (Alternative) – Run Locally with PowerShell 7

1. Install PowerShell 7:
   https://github.com/PowerShell/PowerShell/releases

2. Install required Azure PowerShell modules:
   ```powershell
   Install-Module Az.Accounts,Az.Compute,Az.Storage,Az.Monitor,Az.Resources,Az.NetAppFiles,Az.CosmosDB,Az.MySql,Az.PostgreSql -Force
   ```

3. Connect to Azure:
   ```powershell
   Connect-AzAccount
   ```

4. Verify permissions:
   Ensure your Azure AD account has "Reader" role on target subscriptions

5. Change to the script directory (where this repo was cloned/unzipped):
   ```powershell
   cd ./Azure
   ```

6. (Windows only, first run) If script execution is blocked you may need (in an elevated PowerShell):
   ```powershell
   Set-ExecutionPolicy -Scope CurrentUser RemoteSigned
   ```

7. Run the script (same parameter syntax as Cloud Shell examples below).

#### Common Parameters
* `-Subscriptions`  Comma-separated list of subscription names or IDs. Omit to include all accessible subscriptions.
* `-Types`          Comma-separated list of resource types to limit discovery (e.g. `VM,Storage,NetApp,SQL,Cosmos`). Omit for all supported types.

#### Supported Resource Types
* `VM` - Virtual Machines with disk sizing
* `Storage` - Storage Accounts with capacity metrics
* `FileShare` - Azure File Shares with usage metrics
* `NetApp` - NetApp Files volumes with capacity metrics
* `SQL` - SQL Managed Instances, SQL Databases, MySQL Servers, PostgreSQL Servers
* `Cosmos` - CosmosDB Accounts with storage metrics

#### Example Invocations
```powershell
# All resources in all accessible subscriptions
./CVAzureCloudSizingScript.ps1

# Only VMs and Storage Accounts in all subscriptions
./CVAzureCloudSizingScript.ps1 -Types VM,Storage

# All resources in specific subscriptions
./CVAzureCloudSizingScript.ps1 -Subscriptions "Production","Development"

# Only VMs in specific subscriptions
./CVAzureCloudSizingScript.ps1 -Types VM -Subscriptions "Production","Development"

# SQL and CosmosDB resources in specific subscriptions
./CVAzureCloudSizingScript.ps1 -Types SQL,Cosmos -Subscriptions "Database-Prod","Analytics-Prod"

# NetApp Files volumes across all subscriptions
./CVAzureCloudSizingScript.ps1 -Types NetApp
```

#### Important Notes for Subscription Names
* **Always use quotes** for subscription names that contain spaces:
  ```powershell
  # CORRECT
  ./CVAzureCloudSizingScript.ps1 -Subscriptions "Dev Test","Production Environment"
  
  # WRONG - will fail
  ./CVAzureCloudSizingScript.ps1 -Subscriptions Dev Test
  ```
* You can use subscription IDs instead of names to avoid spacing issues
* The script will show available subscriptions if specified ones are not found

#### Results & Output
The script creates a timestamped output directory with the following files:
- `azure_vm_info_YYYY-MM-DD_HHMMSS.csv` - VM inventory with disk sizing
- `azure_storage_accounts_info_YYYY-MM-DD_HHMMSS.csv` - Storage Account inventory with capacity metrics
- `azure_file_shares_info_YYYY-MM-DD_HHMMSS.csv` - File Share inventory with capacity metrics
- `azure_netapp_volumes_info_YYYY-MM-DD_HHMMSS.csv` - NetApp Files volume inventory with capacity metrics
- `azure_sql_managed_instances_YYYY-MM-DD_HHMMSS.csv` - SQL Managed Instances inventory
- `azure_sql_databases_inventory_YYYY-MM-DD_HHMMSS.csv` - SQL Databases inventory
- `azure_mysql_servers_YYYY-MM-DD_HHMMSS.csv` - MySQL Servers inventory
- `azure_postgresql_servers_YYYY-MM-DD_HHMMSS.csv` - PostgreSQL Servers inventory
- `azure_cosmosdb_accounts_YYYY-MM-DD_HHMMSS.csv` - CosmosDB Accounts inventory with storage metrics
- `azure_inventory_summary_YYYY-MM-DD_HHMMSS.csv` - Comprehensive summary with regional breakdowns
- `azure_sizing_script_output_YYYY-MM-DD_HHMMSS.log` - Complete execution log
- `azure_sizing_YYYY-MM-DD_HHMMSS.zip` - ZIP archive containing all output files

The script automatically creates a ZIP archive of all results. In Azure Cloud Shell you can download this via the built-in file browser; locally you will find it in the same folder you executed the script from. Share the ZIP file with your Commvault representative for sizing analysis.

#### Performance Notes
- Azure Monitor metrics are collected using Maximum aggregation over a 1-hour time period for efficient data retrieval