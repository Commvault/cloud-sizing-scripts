# AWS 
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

   Install-AWSToolsModule -Name AWS.Tools.Common,AWS.Tools.EC2,AWS.Tools.S3,AWS.Tools.SecurityToken,AWS.Tools.IdentityManagement,AWS.Tools.CloudWatch,AWS.Tools.RDS,AWS.Tools.DynamoDBv2,AWS.Tools.Redshift,AWS.Tools.FSx,AWS.Tools.ElasticFileSystem,AWS.Tools.EKS -Scope CurrentUser -CleanUp -Force -Confirm:$false
   ```
4. Run the script with desired parameters:
   ```powershell
   ./CVAWSCloudSizingScript.ps1 -DefaultProfile -Regions "us-west-2"
   ```

Common script parameters
- -DefaultProfile — Uses default AWS CLI profile / CloudShell role.
- -UserSpecifiedProfileNames "Profile1,Profile2" — comma-separated local profiles.
- -AllLocalProfiles — process all local profiles given in Credential File.
- -ProfileLocation "<path>" — shared Credentials file path.
- -CrossAccountRoleName "<RoleName>" — role to assume in target accounts.
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
- `<AccountId>_summary_YYYY-MM-DD_HHMMSS.xlsx` — per-account Excel summary & detail sheets(EC2, S3, RDS, FSx, EFS, DynamoDB, Redshift, EKS)
- `comprehensive_all_aws_accounts_summary_YYYY-MM-DD_HHMMSS.xlsx` — consolidated workbook
- `aws_sizing_script_output_YYYY-MM-DD_HHMMSS.log` — execution log
- `aws_sizing_results_YYYY-MM-DD_HHMMSS.zip` — ZIP archive 

**Note:** Ensure the executing user has all necessary AWS permissions. The required IAM permissions are included in the script header.
