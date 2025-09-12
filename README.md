# Cloud Sizing Scripts

This repository contains PowerShell scripts for cloud resource discovery. These scripts are designed to assist Commvault representatives in gathering information about cloud resources that may need protection, and help representatives in estimating the cost of protecting these resources. For setup instructions and steps to run, refer to the individual script files in each cloud provider folder.

# AWS 
Script: `AWS/CVAWSCloudSizingScript.ps1` Discovers AWS resources and produce sizing/inventory reports to support Commvault protection planning.

**Supported workloads:** `EC2, S3, EFS, FSx, RDS, DynamoDB, Redshift`

Execution Instructions
----------------------

Two ways to run the AWS sizing script — CloudShell, Local PowerShell.

Method 1 — Run in AWS CloudShell 
1. Sign in to the AWS Console and open CloudShell.
2. Enter PowerShell:
   ```powershell
   pwsh
   ```
3. (Install ImportExcel in CloudShell if Excel output is required)
   ```powershell
   Install-Module -Name ImportExcel -Scope CurrentUser -Force
   ```
4. Upload `CVAWSCloudSizingScript.ps1` to CloudShell and run:
   ```powershell
   ./CVAWSCloudSizingScript.ps1 -DefaultProfile -Regions "us-east-1"
   ```
5. (Optional) Make executable:
   ```bash
   chmod +x CVAWSCloudSizingScript.ps1
   ```

Method 2 — Run locally 
1. Install PowerShell 7:
   https://github.com/PowerShell/PowerShell/releases
3. Install required modules (example consolidated command):
   ```powershell
   # remove any loaded AWSTools modules first (optional)
   Get-Module AWS.Tools.* | Remove-Module -Force

   # install ImportExcel and AWSTools installer then required AWSTools modules
   Install-Module -Name ImportExcel -Scope CurrentUser -Force -Confirm:$false
   Install-Module -Name AWS.Tools.Installer -Scope CurrentUser -Force -Confirm:$false

   Install-AWSToolsModule -Name AWS.Tools.Common,AWS.Tools.EC2,AWS.Tools.S3,AWS.Tools.SecurityToken,AWS.Tools.IdentityManagement,AWS.Tools.CloudWatch,AWS.Tools.RDS,AWS.Tools.DynamoDBv2,AWS.Tools.Redshift,AWS.Tools.FSx,AWS.Tools.ElasticFileSystem -Scope CurrentUser -CleanUp -Force -Confirm:$false
   ```
   (Add modules if you require additional AWS services.)
4. Run the script with desired parameters:
   ```powershell
   ./CVAWSCloudSizingScript.ps1 -DefaultProfile -Regions "us-west-2"
   ```

Common script parameters
- -DefaultProfile — use default AWS CLI profile / CloudShell role.
- -UserSpecifiedProfileNames "Profile1,Profile2" — comma-separated local profiles.
- -AllLocalProfiles — process all local profiles.
- -ProfileLocation "<path>" — shared credentials file path.
- -CrossAccountRoleName "<RoleName>" — role to assume in target accounts.
- -UserSpecifiedAccounts "123456789012,098765432112" — comma-separated account IDs.
- -Regions "us-east-1,us-west-2" — comma-separated regions to query.

Example invocations
```powershell
# CloudShell using CloudShell role (default IAM role)
./CVAWSCloudSizingScript.ps1 -DefaultProfile -Regions "us-east-1"

# CloudShell using uploaded credentials file
./CVAWSCloudSizingScript.ps1 -UserSpecifiedProfileNames "Profile1" -ProfileLocation "./Creds.txt" -Regions "us-east-1"

# Local, using specific credential file and profiles
./CVAWSCloudSizingScript.ps1 -UserSpecifiedProfileNames "prod,dev" -ProfileLocation "./Creds.txt" -Regions "us-east-1,us-west-2"

# Cross-account role using file with account IDs [CloudShell]
./CVAWSCloudSizingScript.ps1 -CrossAccountRoleName "InventoryRole" -UserSpecifiedAccounts "123456789012" -Regions "us-east-1"
```

Outputs
-------
Files are written to the working directory with timestamps:
- `<AccountId>_summary_YYYY-MM-DD_HHMMSS.xlsx` — per-account Excel summary & detail sheets
- `comprehensive_all_aws_accounts_summary_YYYY-MM-DD_HHMMSS.xlsx` — consolidated workbook
- `aws_sizing_script_output_YYYY-MM-DD_HHMMSS.log` — execution log
- `aws_sizing_results_YYYY-MM-DD_HHMMSS.zip` — ZIP archive 

**Note:** Ensure the executing user has all necessary AWS permissions. The required IAM permissions are included in the script header.


## Azure
Script: `Azure/CVAzureCloudSizingScript.ps1`
Discovers Azure cloud resources to assist with Commvault protection planning.

## Google Cloud
Script: `GoogleCloud/CVGoogleCloudSizingScript.ps1`
Discovers Google Cloud resources to assist with Commvault protection planning.

### Execution Instructions

Below are two ways to run the Google Cloud sizing script. Method 1 (Local Powershell) is the recommended as Google Cloud Shell timeouts long processes leading to non graceful terminations of the execution.

#### Method 1 (Recommendation: **Medium to Large Scale GC Environments**) – Run Locally with PowerShell 7

1. Install PowerShell 7:
    https://github.com/PowerShell/PowerShell/releases

2. Install Google Cloud SDK:
    https://cloud.google.com/sdk/docs/install

3. Authenticate:
    ```powershell
    gcloud auth login
    ```

4. Verify permissions:
    Ensure the authenticated account has Viewer (or higher) on each project you want to include.

5. Change to the script directory (where this repo was cloned/unzipped):
    ```powershell
    cd ./GoogleCloud
    ```

6. (Windows only, first run) If script execution is blocked you may need (in an elevated PowerShell):
    ```powershell
    Set-ExecutionPolicy -Scope CurrentUser RemoteSigned
    ```

7. Run the script (same parameter syntax as Cloud Shell examples below).

#### Method 2 (Recommendedation: **Small Scale GC Environments**) – Run in Google Cloud Shell

1. (Optional) Review Cloud Shell basics:
    https://cloud.google.com/shell/docs

2. Confirm permissions:
    Your identity must have at least the Viewer role (or equivalent list/get permissions) on each target project.

3. Launch Cloud Shell:
    - Sign in to the Google Cloud Console.
    - Click the Cloud Shell (terminal) icon.

4. Upload the script:
    - Use the upload button to add `CVGoogleCloudSizingScript.ps1` (found in `GoogleCloud/`).
    - Enter PowerShell:
      ```bash
      pwsh
      ```
    - (Optional) Make executable (mainly if you switched shells first):
      ```bash
      chmod +x CVGoogleCloudSizingScript.ps1
      ```

5. Run the script (examples below). With no parameters it scans all accessible projects and all supported workload types.

#### Common Parameters
* `-Projects`  Comma‑separated list of GCP project IDs. Omit to include all projects visible to your credentials.
* `-Types`     Comma‑separated list of workload types to limit discovery (e.g. `VM,Storage,Fileshare`). Omit for all supported types.
* (Review the script header for any advanced/optional parameters.)

#### Example Invocations
```powershell
# All workloads in all accessible projects
./CVGoogleCloudSizingScript.ps1

# Only VM and Storage workloads in all accessible projects
./CVGoogleCloudSizingScript.ps1 -Types VM,Storage

# All workloads in specific projects
./CVGoogleCloudSizingScript.ps1 -Projects my-gcp-project-1,my-gcp-project-2

# Only VMs in specific projects
./CVGoogleCloudSizingScript.ps1 -Types VM -Projects my-gcp-project-1,my-gcp-project-2
```

#### Results & Output
The script writes logs, CSV summaries, and any tree/structure reports to the working directory with timestamped filenames (often later bundled into a ZIP). In Cloud Shell you can download these via the built‑in file browser; locally you will find them in the same folder you executed the script from. Share the ZIP or individual CSVs with the team as needed.

