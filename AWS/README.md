# AWS 

## Overview

This PowerShell Script inventories AWS workloads across one or multiple accounts and regions.  
It provides a unified view of key AWS resources, their configurations, and capacity metrics to assist with cost analysis, right-sizing, and capacity planning.

The script collects information for the following services:
- **EC2** — Instances with attached/unattached EBS volumes  
- **S3** — Buckets and storage metrics  
- **EFS** — Elastic File Systems  
- **FSx** — File systems (ONTAP/SVX, Windows, Lustre, etc.)  
- **RDS** — Database Instances  
- **DynamoDB** — Tables with provisioned throughput details  
- **DocumentDB** — Clusters  
- **Redshift** — Clusters  
- **EKS** — Clusters and Associated PVCs

The inventory captures **provisioned size details** for the above services, along with other configuration and metadata information.  
Results are exported as timestamped Excel workbooks and consolidated ZIP archives for easy sharing and reporting.

## Requirements
- **PowerShell 7**  
  Download: https://github.com/PowerShell/PowerShell/releases  
- **AWS CLI** (for local powershell runs)  
  Install guide: https://docs.aws.amazon.com/cli/latest/userguide/getting-started-install.html  
- **PowerShell Modules**  
  - ImportExcel  
  - AWS.Tools.Common  
  - AWS.Tools.EC2  
  - AWS.Tools.S3  
  - AWS.Tools.SecurityToken  
  - AWS.Tools.IdentityManagement  
  - AWS.Tools.CloudWatch  
  - AWS.Tools.RDS  
  - AWS.Tools.DynamoDBv2  
  - AWS.Tools.Redshift  
  - AWS.Tools.FSx  
  - AWS.Tools.ElasticFileSystem  
  - AWS.Tools.EKS  
  - AWS.Tools.DocDB
  
Execution Instructions
----------------------

Two ways to run the AWS sizing script — CloudShell, Local PowerShell.

Method 1 — Run in AWS CloudShell 
1. Sign in to the AWS Console and open CloudShell.
2. Enter PowerShell:
   ```powershell
   pwsh
   ```
3. (Install ImportExcel in CloudShell)
   ```powershell
   Install-Module -Name ImportExcel -Scope CurrentUser -Force
   ```
   Note: AWS.Tools modules are pre-installed in CloudShell.
4. (Optional) Make executable:
   ```bash
   chmod +x CVAWSCloudSizingScript.ps1
   ```
5. Run using the default IAM role:
   ```powershell
   ./CVAWSCloudSizingScript.ps1 -DefaultProfile -Regions "us-east-1"
   ```
6. Run using uploaded Creds.txt file with profiles:
   ```powershell
   ./CVAWSCloudSizingScript.ps1 -UserSpecifiedProfileNames "Profile1,Profile2" -ProfileLocation "./Creds.txt" -Regions "us-east-1,us-west-2"
   ```


Method 2 — Run locally
1. Install PowerShell 7:
   https://github.com/PowerShell/PowerShell/releases
2. Install AWS CLI:
   https://docs.aws.amazon.com/cli/latest/userguide/getting-started-install.html
3. Install required modules (example consolidated command):
   ```powershell
   # remove any loaded AWSTools modules first (optional)
   Get-Module AWS.Tools.* | Remove-Module -Force

   # install ImportExcel and AWSTools installer then required AWSTools modules
   Install-Module -Name ImportExcel -Scope CurrentUser -Force -Confirm:$false
   Install-Module -Name AWS.Tools.Installer -Scope CurrentUser -Force -Confirm:$false

   Install-AWSToolsModule -Name AWS.Tools.Common,AWS.Tools.EC2,AWS.Tools.S3,AWS.Tools.SecurityToken,AWS.Tools.IdentityManagement,AWS.Tools.CloudWatch,AWS.Tools.RDS,AWS.Tools.DynamoDBv2,AWS.Tools.Redshift,AWS.Tools.FSx,AWS.Tools.ElasticFileSystem,AWS.Tools.EKS,AWS.Tools.DocDB -Scope CurrentUser -CleanUp -Force -Confirm:$false
   ```
4. Verify required modules are installed:
   ```powershell
   Get-Module -ListAvailable AWS.Tools.* , ImportExcel | Select-Object Name, Version, Path
   ```
5. Fix unsigned script error:
   If you encounter the following error when running the script on Windows:
   .\CVAWSCloudSizingScript.ps1 cannot be loaded because it is not digitally signed.
   This is due to PowerShell's execution policy restricting unsigned scripts. To temporarily allow the script to run in the current session, execute the following command in PowerShell (run as Administrator if necessary):

   ```powershell
   Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
   ```
6. Run the script with desired parameters:
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


Credential Files:
- Creds.txt (AWS shared credentials format)
```ini
[Profile1]
aws_access_key_id = <AccessKey1>
aws_secret_access_key = <SecretKey1>

[Profile2]
aws_access_key_id = <AccessKey2>
aws_secret_access_key = <SecretKey2>
```

- Accounts.txt (one AWS account ID per line, no commas):
```text
123456789012
987654321098
555555555555
```

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


Required IAM Permissions
-------
The executing user/role must have the following IAM permissions for the script to run successfully.

```json
{
    "Version": "2012-10-17",
    "Statement": [
        {
            "Effect": "Allow",
            "Action": [
                "sts:GetCallerIdentity",
                "sts:AssumeRole",
                "iam:ListAccountAliases",
                "ec2:DescribeRegions",
                "ec2:DescribeInstances",
                "ec2:DescribeVolumes",
                "ec2:DescribeTags",
                "s3:GetBucketLocation",
                "s3:ListAllMyBuckets",
                "s3:GetBucketTagging",
                "s3:ListBucket",
                "cloudwatch:GetMetricStatistics",
                "cloudwatch:ListMetrics",
                "elasticfilesystem:DescribeFileSystems",
                "elasticfilesystem:ListTagsForResource",
                "elasticfilesystem:DescribeTags",
                "fsx:DescribeFileSystems",
                "fsx:DescribeVolumes",
                "fsx:ListTagsForResource",
                "fsx:DescribeStorageVirtualMachines",
                "rds:DescribeDBInstances",
                "rds:DescribeDBClusters",
                "rds:ListTagsForResource",
                "dynamodb:ListTables",
                "dynamodb:DescribeTable",
                "dynamodb:ListTagsOfResource",
                "redshift:DescribeClusters",
                "redshift:DescribeTags",
                "eks:ListClusters",
                "eks:DescribeCluster",
                "eks:ListNodegroups",
                "eks:ListTagsForResource"
            ],
            "Resource": "*"
        }
    ]
}
```
**Important: EKS (Kubernetes Workload)**
- To collect in-cluster workload details (such as PVCs, Nodes), the executing IAM User/Role must be added to the EKS cluster’s `aws-auth` ConfigMap with appropriate Kubernetes Role-Based Access Control permissions. Otherwise, only basic cluster metadata will be available.

Example `aws-auth` ConfigMap
```yaml
apiVersion: v1
kind: ConfigMap
metadata:
  name: aws-auth
  namespace: kube-system
data:
  mapUsers: '[{
      "userarn": "arn:aws:iam::123456789012:user/ScriptUser1",
      "username": "scriptuser1",
      "groups": ["system:masters"]
    },
    {
      "userarn": "arn:aws:iam::123456789012:user/ScriptUser2",
      "username": "scriptuser2",
      "groups": ["system:masters"]
    }]'
  mapRoles: |
    - rolearn: arn:aws:iam::123456789012:role/ScriptUserRole
      username: system:node:{{EC2PrivateDNSName}}
      groups:
        - system:bootstrappers
        - system:nodes
```        
- mapUsers — Maps IAM users to Kubernetes groups.
- mapRoles — Maps IAM roles (e.g., node instance roles) to Kubernetes groups.
- system:masters — Grants full cluster-admin permissions.

