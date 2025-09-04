#requires -Version 7.0
#requires -Modules ImportExcel, AWS.Tools.Common, AWS.Tools.EC2, AWS.Tools.S3, AWS.Tools.RDS, AWS.Tools.SecurityToken, AWS.Tools.Organizations, AWS.Tools.IdentityManagement, AWS.Tools.CloudWatch, AWS.Tools.ElasticFileSystem, AWS.Tools.SSO, AWS.Tools.SSOOIDC, AWS.Tools.FSx, AWS.Tools.Backup, AWS.Tools.CostExplorer, AWS.Tools.DynamoDBv2, AWS.Tools.SQS, AWS.Tools.SecretsManager, AWS.Tools.KeyManagementService, AWS.Tools.EKS
<#
.SYNOPSIS
    AWS Cloud Sizing Script – Comprehensive EC2 and Storage inventory and sizing analysis.

.DESCRIPTION
    Inventories AWS EC2 instances, volumes, and S3 storage across one or multiple accounts and regions.
    Supports multiple authentication methods including:
      - Default AWS CLI profile
      - User-specified profiles
      - All locally configured profiles
      - Cross-account role assumption
    Generates detailed CSV reports with sizing information (GB, TB, TiB).
    Includes hierarchical progress tracking and detailed logging.
    Outputs timestamped CSV files and creates a ZIP archive of all results.

.PARAMETER DefaultProfile
    Use the default AWS CLI profile for authentication.

.PARAMETER UserSpecifiedProfileNames
    Comma-separated list of AWS CLI profiles to use.

.PARAMETER AllLocalProfiles
    Use all locally configured AWS CLI profiles.

.PARAMETER ProfileLocation
    Path to a shared credentials file (e.g., .\Creds.txt).

.PARAMETER CrossAccountRoleName
    Name of the IAM role to assume for cross-account access.

.PARAMETER UserSpecifiedAccounts
    Comma-separated list of AWS account IDs to use with the specified cross-account role.

.PARAMETER UserSpecifiedAccountsFile
    Path to a file containing AWS account IDs (one per line) to use with the specified cross-account role.

.PARAMETER Regions
    Comma-separated list of AWS regions to query.

.PARAMETER ExternalId
    External ID required for cross-account role assumption (Optional).

.EXAMPLES
    # Use the default AWS CLI profile
    .\CVAWSCloudSizingScript.ps1 -DefaultProfile -Regions "us-west-1"

    # Use specific profiles with a credentials file having multiple accounts
    .\CVAWSCloudSizingScript.ps1 -UserSpecifiedProfileNames "Profile1,Profile2" -ProfileLocation '.\Creds.txt' -Regions "us-west-1"

    # Use all local profiles with a credentials file
    .\CVAWSCloudSizingScript.ps1 -AllLocalProfiles -Regions "us-west-1" -ProfileLocation '.\Creds.txt'

    # Assume a cross-account role for an account
    .\CVAWSCloudSizingScript.ps1 -CrossAccountRoleName "DRVSA_TenantRole" -UserSpecifiedAccounts "123456789012" -Regions "us-west-2" -ExternalId "###"

    # Assume a cross-account role for an account from a file
    .\CVAWSCloudSizingScript.ps1 -CrossAccountRoleName "InventoryRole" -UserSpecifiedAccountsFile ".\Accounts.txt" -Regions "us-east-1,us-west-2"

.OUTPUTS
    Creates a timestamped output directory with the following files:
    - aws_ec2_info_YYYY-MM-DD_HHMMSS.csv   – EC2 instance inventory with disk sizing
    - aws_s3_info_YYYY-MM-DD_HHMMSS.csv    – S3 bucket inventory with storage metrics
    - aws_inventory_summary_YYYY-MM-DD_HHMMSS.csv – Comprehensive summary with regional breakdowns
    - aws_sizing_script_output_YYYY-MM-DD_HHMMSS.log – Complete execution log
    - aws_sizing_YYYY-MM-DD_HHMMSS.zip     – ZIP archive containing all output files

.NOTES
    Requirements:
    - PowerShell 7 and the AWS.Tools.* modules (see #requires directive at the top).
    - AWS CLI profiles or credentials must be configured prior to running the script.

    IAM Permissions:
    The IAM role or user running this script must have the following permissions:
    {
        "Version": "2012-10-17",
        "Statement": [
            {
                "Effect": "Allow",
                "Action": [
                    "ec2:DescribeInstances",
                    "ec2:DescribeVolumes",
                    "ec2:DescribeRegions",
                    "s3:ListAllMyBuckets",
                    "s3:GetBucketLocation",
                    "s3:ListBucket",
                    "s3:GetBucketTagging",
                    "cloudwatch:GetMetricData",
                    "cloudwatch:GetMetricStatistics",
                    "cloudwatch:ListMetrics",
                    "sts:GetCallerIdentity",
                    "iam:ListAccountAliases"
                ],
                "Resource": "*"
            }
        ]
    }

    Credential Files:
    - Creds.txt (AWS shared credentials format):
        [Profile1]
        aws_access_key_id = <AccessKey1>
        aws_secret_access_key = <SecretKey1>

        [Profile2]
        aws_access_key_id = <AccessKey2>
        aws_secret_access_key = <SecretKey2>

    - Accounts.txt (one AWS account ID per line, no commas):
        123456789012
        987654321098
        555555555555

    EC2 disk sizing includes both root volumes and attached EBS volumes with error handling.
    S3 storage metrics are retrieved via AWS APIs and may require CloudWatch metrics or S3 Storage Lens.

    SETUP INSTRUCTIONS FOR LOCAL SYSTEM (Recommended):

    1. Install PowerShell 7:
       https://github.com/PowerShell/PowerShell/releases

    2. Install AWS Tools for PowerShell (as listed in #requires):
       Install-Module -Name AWSPowerShell.NetCore -Force

    3. Configure AWS CLI profiles:
       aws configure --profile MyProfile

    4. Verify AWS IAM permissions:
       Ensure your IAM user/role has the required permissions listed above.

    5. Run the script:
       .\CVAWSCloudSizingScript.ps1 -DefaultProfile -Regions "us-west-1"
       .\CVAWSCloudSizingScript.ps1 -CrossAccountRoleName "InventoryRole" -UserSpecifiedAccountsFile ".\Accounts.txt" -Regions "us-east-1"

    SETUP INSTRUCTIONS FOR AWS CLOUDSHELL (Alternative):

    1. Learn about AWS CloudShell:
       https://docs.aws.amazon.com/cloudshell/

    2. Verify AWS IAM permissions:
       Ensure your IAM user or role has read permissions for EC2 and S3.

    3. Access AWS CloudShell:
       - Login to AWS Console
       - Open CloudShell in your target region

    4. Upload this script:
       Use the CloudShell file upload feature to upload CVAWSCloudSizingScript.ps1

    5. Run the script:
       ./CVAWSCloudSizingScript.ps1
       ./CVAWSCloudSizingScript.ps1 -UserSpecifiedProfileNames "Profile1,Profile2" -Regions "us-east-1,us-west-2"
#>

[CmdletBinding(DefaultParameterSetName = 'DefaultProfile')]
param (
  [Parameter(ParameterSetName='AllLocalProfiles',Mandatory=$true)]
  [ValidateNotNullOrEmpty()][switch]$AllLocalProfiles,

  [Parameter(ParameterSetName='CrossAccountRole',Mandatory=$true)]
  [ValidateNotNullOrEmpty()][string]$CrossAccountRoleName,
  [string]$ExternalId,

  [Parameter(ParameterSetName='DefaultProfile')][switch]$DefaultProfile,

  [Parameter(ParameterSetName='UserSpecifiedProfiles',Mandatory=$true)]
  [ValidateNotNullOrEmpty()][string]$UserSpecifiedProfileNames,

  [Parameter(ParameterSetName='CrossAccountRole')]
  [ValidateNotNullOrEmpty()][string]$UserSpecifiedAccounts,

  [Parameter(ParameterSetName='CrossAccountRole')]
  [ValidateNotNullOrEmpty()][string]$UserSpecifiedAccountsFile,

  [ValidateSet("GovCloud","")][string]$Partition,
  [string]$ProfileLocation,
  [string]$Regions,
  [string]$RegionToQuery,
  [switch]$SkipBucketTags,
  [switch]$DebugBucketTags
)

# Set culture to en-US for consistent date and time formatting
$CurrentCulture = [System.Globalization.CultureInfo]::CurrentCulture
[System.Threading.Thread]::CurrentThread.CurrentCulture = 'en-US'
[System.Threading.Thread]::CurrentThread.CurrentUICulture = 'en-US'

# Default API query regions
$defaultQueryRegion = "us-east-1"
$defaultGovCloudQueryRegion = "us-gov-east-1"

# Timestamp for output file names
$date = Get-Date
$date_string = $date.ToString("yyyy-MM-dd_HHmmss")

# CloudWatch metric timeframe (last 7 days)
$utcEndTime = $date.ToUniversalTime()
$utcStartTime = $utcEndTime.AddDays(-7)

# Set up logging
$output_log = "output_aws_$date_string.log"
if (Test-Path "./$output_log") { Remove-Item -Path "./$output_log" }
Start-Transcript -Path "./$output_log"

# Display script parameters
Write-Host "Arguments passed:" -ForegroundColor Green
$PSBoundParameters | Format-Table | Out-String

# Handle custom profile location if specified
$profileLocationOpt = @{}
if ($ProfileLocation) {
  $profileLocationOpt = @{ProfileLocation = $ProfileLocation}
  Write-Host "Using Profile Location: $ProfileLocation"
}

# Output file base names
$baseOutputEc2Instance = "aws_ec2_instance_info"
$baseOutputEc2UnattachedVolume = "aws_ec2_unattached_volume_info"
$baseOutputS3 = "aws_s3_info"
$archiveFile = "aws_sizing_results_$date_string.zip"

# Collections for multi-account processing
$ec2ListByAccount = @{}                   # Dictionary to store EC2 instances by account
$ec2UnattachedVolListByAccount = @{}      # Dictionary to store unattached EBS volumes by account
$s3ListByAccount = @{}                    # Dictionary to store S3 buckets by account
$accountsProcessed = [System.Collections.ArrayList]::new()  # List of processed accounts
$allOutputFiles = [System.Collections.ArrayList]::new()     # List of generated output files

# Main function to collect AWS data for a given credential/account
function getAWSData($cred) {
  # Determine which regions to inventory
  if ($Regions -and $Regions.Trim() -ne '') {
    # Use user-specified regions
    [string[]]$awsRegions = $Regions.Split(',') | ForEach-Object { $_.Trim() } | Where-Object { $_ }
  } else {
    # Query all available regions
    try {
      $awsRegions = Get-EC2Region @profileLocationOpt -Region $queryRegion -Credential $cred | Select-Object -ExpandProperty RegionName
    } catch {
      Write-Host "Failed to list EC2 regions (query region $queryRegion)" -ForegroundColor Red
      Write-Host $_
      return
    }
  }

  # Get current AWS account info
  Write-Host "Current identity:" -ForegroundColor Green
  try { 
    $awsAccountInfo = Get-STSCallerIdentity -Credential $cred -Region $queryRegion -ErrorAction Stop 
  } catch { 
    Write-Host $_ -ForegroundColor Red
    return 
  }
  $awsAccountInfo | Format-Table | Out-String

  # Try to get account alias if available
  try { 
    $awsAccountAlias = Get-IAMAccountAlias -Credential $cred -Region $queryRegion -ErrorAction Stop 
  } catch { 
    $awsAccountAlias = $null 
  }

  # Initialize collections for this account if not already done
  $accountId = $awsAccountInfo.Account
  if (-not $ec2ListByAccount.ContainsKey($accountId)) {
    $ec2ListByAccount[$accountId] = [System.Collections.ArrayList]::new()
    $ec2UnattachedVolListByAccount[$accountId] = [System.Collections.ArrayList]::new()
    $s3ListByAccount[$accountId] = [System.Collections.ArrayList]::new()
    $accountsProcessed.Add(@{
      AccountId = $accountId
      AccountAlias = $awsAccountAlias
    }) | Out-Null
  }

  # Get reference to account-specific collections
  $ec2List = $ec2ListByAccount[$accountId]
  $ec2UnattachedVolList = $ec2UnattachedVolListByAccount[$accountId]
  $s3List = $s3ListByAccount[$accountId]

  # Process each AWS region
  $awsRegionCounter = 1
  foreach ($awsRegion in $awsRegions) {
    Write-Progress -ID 2 -Activity "Region $awsRegion" -Status "Region $awsRegionCounter of $($awsRegions.Count)" -PercentComplete (($awsRegionCounter / $awsRegions.Count)*100)
    $awsRegionCounter++

    #-----------------------------------------
    # S3 Buckets Inventory
    #-----------------------------------------
    try {
      $s3Buckets = (Get-S3Bucket -Credential $cred -Region $awsRegion -BucketRegion $awsRegion -ErrorAction Stop).BucketName
    } catch {
      Write-Host "Failed to get S3 buckets for region $awsRegion (Acct $($awsAccountInfo.Account))" -ForegroundColor Red
      if ($DebugBucketTags) { Write-Host $_ -ForegroundColor DarkYellow }
      $s3Buckets = @()
    }

    $bCount = 1
    foreach ($s3Bucket in $s3Buckets) {
      Write-Progress -ID 3 -Activity "Bucket $s3Bucket" -Status "$bCount / $($s3Buckets.Count)" -PercentComplete (($bCount / $s3Buckets.Count)*100)
      $bCount++

      # CloudWatch metrics enumeration for bucket size and object count
      $filter = [Amazon.CloudWatch.Model.DimensionFilter]::new()
      $filter.Name = 'BucketName'
      $filter.Value = $s3Bucket

      try {
        # Get the storage types for this bucket (Standard, StandardIA, etc.)
        $bytesStorageTypes = (Get-CWMetricList -Dimension $filter -Credential $cred -Region $awsRegion -ErrorAction Stop |
          Where-Object MetricName -eq 'BucketSizeBytes' |
          Select-Object -ExpandProperty Dimensions |
          Where-Object Name -eq StorageType).Value
        $numObjStorageTypes = (Get-CWMetricList -Dimension $filter -Credential $cred -Region $awsRegion -ErrorAction Stop |
          Where-Object MetricName -eq 'NumberOfObjects' |
          Select-Object -ExpandProperty Dimensions |
          Where-Object Name -eq StorageType).Value
      } catch {
        Write-Host ("Metric enumeration failed for bucket {0} ({1})" -f $s3Bucket,$awsRegion) -ForegroundColor Yellow
        if ($DebugBucketTags) { Write-Host $_ -ForegroundColor DarkYellow }
      }

      # Common dimension for bucket metrics
      $bucketNameDim = [Amazon.CloudWatch.Model.Dimension]::new()
      $bucketNameDim.Name  = "BucketName"
      $bucketNameDim.Value = $s3Bucket

      # Get storage sizes for each storage type
      $bytesStorages = @{}
      foreach ($bytesStorageType in $bytesStorageTypes) {
        $bucketBytesStorageDim = [Amazon.CloudWatch.Model.Dimension]::new()
        $bucketBytesStorageDim.Name = "StorageType"; $bucketBytesStorageDim.Value = $bytesStorageType
        try {
          $maxBucketSizes = (Get-CWMetricStatistic -Statistic Maximum `
              -Namespace AWS/S3 -MetricName BucketSizeBytes `
              -StartTime $utcStartTime.AddDays(-1).ToString("yyyy-MM-ddTHH:mm:ssZ") `
              -EndTime $utcEndTime.ToString("yyyy-MM-ddTHH:mm:ssZ") `
              -Period 86400 -Credential $cred -Region $awsRegion `
              -Dimension $bucketNameDim,$bucketBytesStorageDim -ErrorAction Stop |
              Select-Object -ExpandProperty Datapoints).Maximum
        } catch {
          if ($DebugBucketTags) {
            Write-Host ("BucketSizeBytes metric failure {0}/{1}" -f $s3Bucket,$bytesStorageType) -ForegroundColor DarkYellow
            Write-Host $_ -ForegroundColor DarkYellow
          }
        }
        $bytesStorages[$bytesStorageType] = ($maxBucketSizes | Measure-Object -Maximum).Maximum
      }

      # Get object counts for each storage type
      $numObjStorages = @{}
      foreach ($numObjStorageType in $numObjStorageTypes) {
        $bucketNumObjStorageDim = [Amazon.CloudWatch.Model.Dimension]::new()
        $bucketNumObjStorageDim.Name = "StorageType"; $bucketNumObjStorageDim.Value = $numObjStorageType
        try {
          $maxBucketObjects = (Get-CWMetricStatistic -Statistic Maximum `
              -Namespace AWS/S3 -MetricName NumberOfObjects `
              -StartTime $utcStartTime.AddDays(-1).ToString("yyyy-MM-ddTHH:mm:ssZ") `
              -EndTime $utcEndTime.ToString("yyyy-MM-ddTHH:mm:ssZ") `
              -Period 86400 -Credential $cred -Region $awsRegion `
              -Dimension $bucketNameDim,$bucketNumObjStorageDim -ErrorAction Stop |
              Select-Object -ExpandProperty Datapoints).Maximum
        } catch {
          if ($DebugBucketTags) {
            Write-Host ("NumberOfObjects metric failure {0}/{1}" -f $s3Bucket,$numObjStorageType) -ForegroundColor DarkYellow
            Write-Host $_ -ForegroundColor DarkYellow
          }
        }
        $numObjStorages[$numObjStorageType] = ($maxBucketObjects | Measure-Object -Maximum).Maximum
      }

      # Get bucket tags if not skipped
      if ($SkipBucketTags) {
        $bucketTags = @()
      } else {
        try {
          $bucketTags = Get-S3BucketTagging -BucketName $s3Bucket -Credential $cred -Region $awsRegion -ErrorAction Stop
        } catch {
          # Silent when TagSet absent
            if ($_.Exception.Message -match 'TagSet does not exist') {
              if ($DebugBucketTags) { Write-Host ("No tag set for bucket {0}" -f $s3Bucket) -ForegroundColor DarkGray }
              $bucketTags = @()
            } else {
              Write-Host ("Tag retrieval issue for {0}: {1}" -f $s3Bucket, $_.Exception.Message) -ForegroundColor Yellow
              if ($DebugBucketTags) { Write-Host $_ -ForegroundColor DarkYellow }
              $bucketTags = @()
            }
        }
      }

      # Create S3 bucket object with base properties
      $s3obj = [PSCustomObject]@{
        AwsAccountId    = $awsAccountInfo.Account
        AwsAccountAlias = $awsAccountAlias
        BucketName      = $s3Bucket
        Region          = $awsRegion
        BackupPlans     = ""
        InBackupPlan    = $false
      }

      # Add storage size properties for each storage type in multiple units
      foreach ($bytesStorage in $bytesStorages.GetEnumerator()) {
        $val = $bytesStorage.Value
        if ($null -eq $val) {
          $sizeBytes = 0
          $s3SizeGB=0;$s3SizeTB=0;$s3SizeGiB=0;$s3SizeTiB=0
        } else {
          $sizeBytes = $val
          $s3SizeGB = $val / 1073741824
          $s3SizeTB = $s3SizeGB / 1000
          $s3SizeGiB = $s3SizeGB / 1.073741824
          $s3SizeTiB = $s3SizeGiB / 1024
        }
        Add-Member -InputObject $s3obj -NotePropertyName ($bytesStorage.Name + "_SizeGB")    -NotePropertyValue ([math]::round($s3SizeGB,3))
        Add-Member -InputObject $s3obj -NotePropertyName ($bytesStorage.Name + "_SizeTB")    -NotePropertyValue ([math]::round($s3SizeTB,4))
        Add-Member -InputObject $s3obj -NotePropertyName ($bytesStorage.Name + "_SizeGiB")   -NotePropertyValue ([math]::round($s3SizeGiB,3))
        Add-Member -InputObject $s3obj -NotePropertyName ($bytesStorage.Name + "_SizeTiB")   -NotePropertyValue ([math]::round($s3SizeTiB,4))
        Add-Member -InputObject $s3obj -NotePropertyName ($bytesStorage.Name + "_SizeBytes") -NotePropertyValue $sizeBytes
      }

      # Add object count properties for each storage type
      foreach ($numObjStorage in $numObjStorages.GetEnumerator()) {
        $count = if ($null -eq $numObjStorage.Value) {0}else{$numObjStorage.Value}
        Add-Member -InputObject $s3obj -MemberType NoteProperty -Name ("NumberOfObjects-" + $numObjStorage.Name) -Value $count
      }

      # Add tag properties
      foreach ($tag in $bucketTags) {
        $key = $tag.Key -replace '[^a-zA-Z0-9]','_'
        Add-Member -InputObject $s3obj -MemberType NoteProperty -Name ("Tag: $key") -Value $tag.Value -Force
      }

      # Add S3 bucket to the account's collection
      $s3List.Add($s3obj) | Out-Null
    }
    Write-Progress -ID 3 -Activity "Buckets done" -Completed

    #-----------------------------------------
    # EC2 Instances Inventory
    #-----------------------------------------
    try {
      $ec2Instances = (Get-EC2Instance -Credential $cred -Region $awsRegion -ErrorAction Stop).Instances
    } catch {
      Write-Host "Failed to get EC2 instances for $awsRegion" -ForegroundColor Red
      $ec2Instances = @()
    }

    $i=1
    foreach ($ec2 in $ec2Instances) {
      Write-Progress -ID 4 -Activity "EC2 $($ec2.InstanceId)" -Status "$i / $($ec2Instances.Count)" -PercentComplete (($i / $ec2Instances.Count)*100)
      $i++
      
      # Calculate total volume size for this instance
      $volSize=0
      $volumes = $ec2.BlockDeviceMappings.Ebs
      foreach ($vol in $volumes) {
        try { 
          $volSize += (Get-EC2Volume -VolumeId $vol.VolumeId -Credential $cred -Region $awsRegion -ErrorAction Stop).Size 
        } catch {}
      }
      
      # Create EC2 instance object with computed sizing in multiple units
      $ec2obj = [PSCustomObject]@{
        AwsAccountId    = $awsAccountInfo.Account
        AwsAccountAlias = $awsAccountAlias
        InstanceId      = $ec2.InstanceId
        Name            = ($ec2.Tags | ForEach-Object { if ($_.Key -ceq 'Name'){ $_.Value } })
        Volumes         = $volumes.Count
        SizeGiB         = $volSize
        SizeTiB         = [math]::Round(($volSize/1024),4)
        SizeGB          = [math]::Round(($volSize * 1.073741824),3)
        SizeTB          = [math]::Round(($volSize * 0.001073741824),4)
        Region          = $awsRegion
        InstanceType    = $ec2.InstanceType
        Platform        = $ec2.Platform
        ProductCode     = $ec2.ProductCodes.ProductCodeType
        BackupPlans     = ""
        InBackupPlan    = $false
      }
      
      # Add tags as properties
      foreach ($tag in $ec2.Tags) {
        $key = $tag.Key -replace '[^a-zA-Z0-9]','_'
        if ($key -ne 'Name') {
          $ec2obj | Add-Member -MemberType NoteProperty -Name "Tag: $key" -Value $tag.Value -Force
        }
      }
      
      # Add EC2 instance to the account's collection
      $ec2List.Add($ec2obj) | Out-Null
    }
    Write-Progress -ID 4 -Activity "EC2 done" -Completed

    #-----------------------------------------
    # Unattached EBS Volumes Inventory
    #-----------------------------------------
    try {
      $ec2UnattachedVolumes = Get-EC2Volume -Credential $cred -Region $awsRegion -Filter @{Name='status';Values='available'} -ErrorAction Stop
    } catch {
      Write-Host "Failed to get unattached volumes for $awsRegion" -ForegroundColor Red
      $ec2UnattachedVolumes = @()
    }
    
    $u=1
    foreach ($uv in $ec2UnattachedVolumes) {
      Write-Progress -ID 5 -Activity "Unattached $($uv.VolumeId)" -Status "$u / $($ec2UnattachedVolumes.Count)" -PercentComplete (($u / $ec2UnattachedVolumes.Count)*100)
      $u++
      
      # Create unattached volume object with computed sizing in multiple units
      $obj = [PSCustomObject]@{
        AwsAccountId    = $awsAccountInfo.Account
        AwsAccountAlias = $awsAccountAlias
        VolumeId        = $uv.VolumeId
        Name            = ($uv.Tags | ForEach-Object { if ($_.Key -ceq 'Name'){ $_.Value } })
        SizeGiB         = $uv.Size
        SizeTiB         = [math]::Round(($uv.Size/1024),4)
        SizeGB          = [math]::Round(($uv.Size * 1.073741824),3)
        SizeTB          = [math]::Round(($uv.Size * 0.001073741824),4)
        Region          = $awsRegion
        VolumeType      = $uv.VolumeType
        BackupPlans     = ""
        InBackupPlan    = $false
      }
      
      # Add tags as properties
      foreach ($tag in $uv.Tags) {
        $key = $tag.Key -replace '[^a-zA-Z0-9]','_'
        if ($key -ne 'Name') {
          $obj | Add-Member -MemberType NoteProperty -Name "Tag: $key" -Value $tag.Value -Force
        }
      }
      
      # Add unattached volume to the account's collection
      $ec2UnattachedVolList.Add($obj) | Out-Null
    }
    Write-Progress -ID 5 -Activity "Unattached done" -Completed
  }
  Write-Progress -ID 2 -Activity "Regions done" -Completed
}

# Process account-specific data and create output files
function processAccountData($accountId, $accountAlias) {
  # Get account-specific data collections
  $ec2List = $ec2ListByAccount[$accountId]
  $ec2UnattachedVolList = $ec2UnattachedVolListByAccount[$accountId]
  $s3List = $s3ListByAccount[$accountId]

  # Create account-specific output filenames
  $acctOutputEc2Instance = "${baseOutputEc2Instance}-${accountId}-$date_string.csv"
  $acctOutputEc2UnattachedVolume = "${baseOutputEc2UnattachedVolume}-${accountId}-$date_string.csv"
  $acctOutputS3 = "${baseOutputS3}-${accountId}-$date_string.csv"
  
  # Process EC2 instances
  if ($ec2List.Count -gt 0) {
    addTagsToAllObjectsInList $ec2List
    $ec2List | Export-Csv -Path $acctOutputEc2Instance -NoTypeInformation
    Write-Host "CSV for account ${accountId} (${accountAlias}): ${acctOutputEc2Instance}" -ForegroundColor Green
    $allOutputFiles.Add($acctOutputEc2Instance) | Out-Null
  }
  
  # Process unattached volumes
  if ($ec2UnattachedVolList.Count -gt 0) {
    addTagsToAllObjectsInList $ec2UnattachedVolList
    $ec2UnattachedVolList | Export-Csv -Path $acctOutputEc2UnattachedVolume -NoTypeInformation
    Write-Host "CSV for account ${accountId} (${accountAlias}): ${acctOutputEc2UnattachedVolume}" -ForegroundColor Green
    $allOutputFiles.Add($acctOutputEc2UnattachedVolume) | Out-Null
  }
  
  # Process S3 buckets
  if ($s3List.Count -gt 0) {
    # S3 normalization - ensure consistent properties across all bucket objects
    $s3Props = $s3List.ForEach{ $_.PSObject.Properties.Name } | Select-Object -Unique
    $s3PropsOrdered = $s3Props | Where-Object { $_ -notmatch '^Tag:\s*' }
    $s3PropsOrdered += $s3Props | Where-Object { $_ -match '^Tag:\s*' } | Sort-Object -Unique
    
    # Create a properly typed ArrayList for the normalized S3 list
    $s3ListNormalized = [System.Collections.ArrayList]::new()
    
    # Add each item individually to ensure proper typing
    foreach($item in ($s3List | Select-Object $s3PropsOrdered)) {
      foreach($p in $s3PropsOrdered){
        if(($p -like "*_Size*" -or $p -like "NumberOfObjects*") -and [string]::IsNullOrEmpty($item.$p)){
          $item.$p = 0
        }
      }
      $s3ListNormalized.Add($item) | Out-Null
    }
    
    $s3ListNormalized | Export-Csv -Path $acctOutputS3 -NoTypeInformation
    Write-Host "CSV for account ${accountId} (${accountAlias}): ${acctOutputS3}" -ForegroundColor Green
    $allOutputFiles.Add($acctOutputS3) | Out-Null
    
    return $s3ListNormalized
  }
  
  return $null
}

# Display account summary with calculated totals
function displayAccountSummary($accountId, $accountAlias) {
  $ec2List = $ec2ListByAccount[$accountId]
  $ec2UnattachedVolList = $ec2UnattachedVolListByAccount[$accountId]
  $s3List = $s3ListByAccount[$accountId]
  
  # Calculate EC2 instance totals in multiple units
  $ec2TotalGiB = ($ec2List.SizeGiB | Measure-Object -Sum).Sum
  $ec2TotalTiB = ($ec2List.SizeTiB | Measure-Object -Sum).Sum
  $ec2TotalGB  = ($ec2List.SizeGB  | Measure-Object -Sum).Sum
  $ec2TotalTB  = ($ec2List.SizeTB  | Measure-Object -Sum).Sum

  # Calculate unattached EBS volume totals in multiple units
  $ec2UnVolTotalGiB = ($ec2UnattachedVolList.SizeGiB | Measure-Object -Sum).Sum
  $ec2UnVolTotalTiB = ($ec2UnattachedVolList.SizeTiB | Measure-Object -Sum).Sum
  $ec2UnVolTotalGB  = ($ec2UnattachedVolList.SizeGB  | Measure-Object -Sum).Sum
  $ec2UnVolTotalTB  = ($ec2UnattachedVolList.SizeTB  | Measure-Object -Sum).Sum
  
  # S3 normalization for consistent property access
  $s3Props = $s3List.ForEach{ $_.PSObject.Properties.Name } | Select-Object -Unique
  $s3PropsOrdered = $s3Props | Where-Object { $_ -notmatch '^Tag:\s*' }
  $s3PropsOrdered += $s3Props | Where-Object { $_ -match '^Tag:\s*' } | Sort-Object -Unique
  $s3ListNormalized = [System.Collections.ArrayList]@($s3List | Select-Object $s3PropsOrdered)
  
  foreach($row in $s3ListNormalized){
    foreach($p in $s3PropsOrdered){
      if(($p -like "*_Size*" -or $p -like "NumberOfObjects*") -and [string]::IsNullOrEmpty($row.$p)){
        $row.$p = 0
      }
    }
  }
  
  # Calculate S3 totals by storage type
  $s3TBProps = $s3PropsOrdered | Select-String -Pattern "_SizeTB"
  $s3ListAg = $s3ListNormalized | Select-Object $s3PropsOrdered
  $s3TotalTBs=@{}
  foreach($p in $s3TBProps){ $s3TotalTBs[$p] = ($s3ListAg.$p | Measure-Object -Sum).Sum }
  $s3TotalTBsFormatted = $s3TotalTBs.GetEnumerator() | ForEach-Object {
    [PSCustomObject]@{ StorageType = $_.Key; Size_TB = "{0:n7}" -f $_.Value }
  }
  
  # Display account summary
  Write-Host
  Write-Host "===== ACCOUNT: $accountId ($accountAlias) =====" -ForegroundColor Cyan
  Write-Host "EC2 instances: $($ec2List.Count)" -ForegroundColor Green
  Write-Host "Attached volumes: $(($ec2List.Volumes | Measure-Object -Sum).Sum)" -ForegroundColor Green
  Write-Host ("Attached capacity: {0} GiB | {1} GB | {2} TiB | {3} TB" -f $ec2TotalGiB,$ec2TotalGB,$ec2TotalTiB,$ec2TotalTB) -ForegroundColor Green
  Write-Host
  Write-Host "Unattached volumes: $($ec2UnattachedVolList.Count)" -ForegroundColor Green
  Write-Host ("Unattached capacity: {0} GiB | {1} GB | {2} TiB | {3} TB" -f $ec2UnVolTotalGiB,$ec2UnVolTotalGB,$ec2UnVolTotalTiB,$ec2UnVolTotalTB) -ForegroundColor Green
  Write-Host
  Write-Host "S3 buckets: $($s3List.Count)" -ForegroundColor Green
  Write-Host "S3 total (by StorageType, TB):" -ForegroundColor Green
  if($s3TotalTBsFormatted.Count -gt 0){
    $s3TotalTBsFormatted | ForEach-Object { Write-Host ("  {0} = {1}" -f $_.StorageType,$_.Size_TB) -ForegroundColor Green }
  } else {
    Write-Host "  (none)" -ForegroundColor Green
  }

  # Region breakdown
  Write-Host
  Write-Host "Region breakdown for account ${accountId}:" -ForegroundColor Cyan
  
  # Create unique list of regions from all resource types
  $allRegions = [System.Collections.ArrayList]::new()
  
  # Process EC2 regions
  if ($ec2List -and $ec2List.Count -gt 0) {
    $ec2List.Region | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | ForEach-Object {
      if (-not $allRegions.Contains($_)) { $allRegions.Add($_) | Out-Null }
    }
  }
  
  # Process unattached volume regions
  if ($ec2UnattachedVolList -and $ec2UnattachedVolList.Count -gt 0) {
    $ec2UnattachedVolList.Region | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | ForEach-Object {
      if (-not $allRegions.Contains($_)) { $allRegions.Add($_) | Out-Null }
    }
  }
  
  # Process S3 regions
  if ($s3ListAg -and $s3ListAg.Count -gt 0) {
    $s3ListAg.Region | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | ForEach-Object {
      if (-not $allRegions.Contains($_)) { $allRegions.Add($_) | Out-Null }
    }
  }
  
  # Sort the regions
  $allRegions = $allRegions | Sort-Object
                
  # Display per-region breakdown
  foreach ($r in $allRegions){
    $ec2R = $ec2List | Where-Object Region -eq $r
    $unR  = $ec2UnattachedVolList | Where-Object Region -eq $r
    $s3R  = $s3ListAg | Where-Object Region -eq $r
    
    # Calculate EC2 totals for this region
    $ec2R_GiB = ($ec2R.SizeGiB | Measure-Object -Sum).Sum
    $ec2R_GB  = ($ec2R.SizeGB  | Measure-Object -Sum).Sum
    $ec2R_TiB = ($ec2R.SizeTiB | Measure-Object -Sum).Sum
    $ec2R_TB  = ($ec2R.SizeTB  | Measure-Object -Sum).Sum
    
    # Calculate unattached volume totals for this region
    $un_GiB = ($unR.SizeGiB | Measure-Object -Sum).Sum
    $un_GB  = ($unR.SizeGB  | Measure-Object -Sum).Sum
    $un_TiB = ($unR.SizeTiB | Measure-Object -Sum).Sum
    $un_TB  = ($unR.SizeTB  | Measure-Object -Sum).Sum
    
    # Calculate S3 totals for this region
    $s3R_TotalTiB = 0
    if ($s3R -and ($s3R -is [array] -or $s3R -is [System.Collections.ArrayList]) -and $s3R.Count -gt 0) {
      # Get the first object to examine its properties
      $firstS3Object = $s3R[0]
      
      # Find all TiB properties by examining property names directly
      $sizeProperties = $firstS3Object.PSObject.Properties.Name | Where-Object { $_ -like '*_SizeTiB' }
      
      # Sum all the TiB sizes
      foreach ($propName in $sizeProperties) {
        $s3R | ForEach-Object {
          if ($_.$propName) {
            $s3R_TotalTiB += $_.$propName
          }
        }
      }
    }
    
    # Display per-region totals
    Write-Host ("Region: {0}" -f $r) -ForegroundColor Green
    Write-Host ("  EC2 Instances: {0}" -f $ec2R.Count)
    Write-Host ("  Attached Volumes: {0}" -f (($ec2R.Volumes | Measure-Object -Sum).Sum))
    Write-Host ("  Attached Capacity: {0} GiB | {1} GB | {2:n4} TiB | {3:n4} TB" -f $ec2R_GiB,$ec2R_GB,$ec2R_TiB,$ec2R_TB)
    Write-Host ("  Unattached Volumes: {0}" -f $unR.Count)
    Write-Host ("  Unattached Capacity: {0} GiB | {1} GB | {2:n4} TiB | {3:n4} TB" -f $un_GiB,$un_GB,$un_TiB,$un_TB)
    Write-Host ("  S3 Buckets: {0}" -f ($s3R | Measure-Object).Count)
    Write-Host ("  S3 Total TiB: {0:n4}" -f $s3R_TotalTiB)
    Write-Host
  }
  
  return $s3ListAg
}

# Helper function to ensure all objects in a list have the same properties
function addTagsToAllObjectsInList($list){
  $allKeys=@{}
  foreach($o in $list){
    foreach($p in $o.PSObject.Properties){
      if(-not $allKeys.ContainsKey($p.Name)){ $allKeys[$p.Name]=$true }
    }
  }
  foreach($o in $list){
    foreach($k in $allKeys.Keys){
      if(-not $o.PSObject.Properties.Name.Contains($k)){
        $o | Add-Member -MemberType NoteProperty -Name $k -Value $null -Force
      }
    }
  }
}

# Main execution block with error handling
try {
  # Set the query region based on parameters and ensure partitionId is always set
  if ($RegionToQuery) { 
    $queryRegion = $RegionToQuery
    $partitionId = if ($Partition -eq 'GovCloud') { 'aws-us-gov' } else { 'aws' }
  }
  elseif ($Partition -eq 'GovCloud') { 
    $queryRegion = $defaultGovCloudQueryRegion
    $partitionId = 'aws-us-gov' 
  }
  else { 
    $queryRegion = $defaultQueryRegion
    $partitionId = 'aws' 
  }

  Write-Host "Using partition: $partitionId" -ForegroundColor Cyan

  # Process based on the parameter set specified by the user
  switch ($PSCmdlet.ParameterSetName) {
    'DefaultProfile' {
      # Use the default AWS credential profile
      try { 
        (Get-STSCallerIdentity @profileLocationOpt -Region $queryRegion) | Out-Null 
      } catch {
        Write-Error "Default credential/profile not set. Run Set-AWSCredential."
        break
      }
      $cred = Get-AWSCredential @profileLocationOpt
      getAWSData $cred
    }
    'UserSpecifiedProfiles' {
      # Process specific named AWS profiles
      $profiles = $UserSpecifiedProfileNames.Split(',') | ForEach-Object { $_.Trim() } | Where-Object { $_ }
      $idx=1
      foreach ($p in $profiles){
        Write-Progress -ID 1 -Activity "Profile $p" -Status "$idx / $($profiles.Count)" -PercentComplete (($idx/$profiles.Count)*100)
        $idx++
        try { 
          $cred = Get-AWSCredential @profileLocationOpt -ProfileName $p 
        } catch { 
          Write-Host "Skip $p (cred error)" -ForegroundColor Yellow
          continue 
        }
        getAWSData $cred
      }
      Write-Progress -ID 1 -Activity "Profiles done" -Completed
    }
    'AllLocalProfiles' {
      # Process all locally configured AWS profiles
      $profiles = (Get-AWSCredential @profileLocationOpt -ListProfileDetail).ProfileName
      $idx=1
      foreach ($p in $profiles){
        Write-Progress -ID 1 -Activity "Profile $p" -Status "$idx / $($profiles.Count)" -PercentComplete (($idx/$profiles.Count)*100)
        $idx++
        try { 
          $cred = Get-AWSCredential @profileLocationOpt -ProfileName $p 
        } catch { 
          continue 
        }
        getAWSData $cred
      }
      Write-Progress -ID 1 -Activity "Profiles done" -Completed
    }
    'CrossAccountRole' {
      # Process using cross-account IAM roles
      try { 
        Get-STSCallerIdentity @profileLocationOpt -Region $queryRegion | Out-Null 
      } catch { 
        Write-Error "Source credential not set."
        break 
      }
      
      if ($UserSpecifiedAccounts -and $UserSpecifiedAccountsFile){ 
        Write-Error "Only one of -UserSpecifiedAccounts / -UserSpecifiedAccountsFile required."
        break 
      }
      
      if ($UserSpecifiedAccountsFile) { 
        $acctIds = Get-Content $UserSpecifiedAccountsFile 
      } elseif ($UserSpecifiedAccounts){ 
        $acctIds = $UserSpecifiedAccounts.Split(',') 
      } else { 
        Write-Error "-UserSpecifiedAccounts or -UserSpecifiedAccountsFile required."
        break 
      }
      
      # Debugging: Check the accounts to be processed
      Write-Host "Accounts to process: $($acctIds -join ', ')" -ForegroundColor Cyan

      $idx=1
      foreach ($acct in $acctIds){
        Write-Progress -ID 1 -Activity "Acct $acct" -Status "$idx / $($acctIds.Count)" -PercentComplete (($idx/$acctIds.Count)*100)
        $idx++
        
        # Construct the ARN properly, using explicit string format to ensure variables are properly expanded
        $roleArn = "arn:$partitionId`:iam::$acct`:role/$CrossAccountRoleName"
        Write-Host "Attempting to assume role with ARN: $roleArn" -ForegroundColor Cyan
        
        try { 
          if ($ExternalId) {
            $cred = (Use-STSRole @profileLocationOpt -RoleArn $roleArn -RoleSessionName $MyInvocation.MyCommand.Name -ExternalId $ExternalId -Region $queryRegion).Credentials
          } else {
            $cred = (Use-STSRole @profileLocationOpt -RoleArn $roleArn -RoleSessionName $MyInvocation.MyCommand.Name -Region $queryRegion).Credentials
          }
        } catch { 
          Write-Host "Failed to assume role for account $acct" -ForegroundColor Red
          Write-Host "Error: $_" -ForegroundColor Red
          continue 
        }
        getAWSData $cred
      }
      Write-Progress -ID 1 -Activity "Cross-account done" -Completed
    }
  }

  # Process collected data for each account
  $combinedS3Lists = [System.Collections.ArrayList]::new()
  Write-Host "Processing data for $($accountsProcessed.Count) accounts..." -ForegroundColor Cyan
  
  foreach ($acctInfo in $accountsProcessed) {
    $accountId = $acctInfo.AccountId
    $accountAlias = $acctInfo.AccountAlias
    
    Write-Host "Processing account: $accountId ($accountAlias)" -ForegroundColor Cyan
    
    # Export files per account
    $s3NormalizedList = processAccountData $accountId $accountAlias
    
    # Generate and display summary for this account
    $s3ListAg = displayAccountSummary $accountId $accountAlias
    
    # Store for combined reports
    if ($s3NormalizedList -and $s3NormalizedList.Count -gt 0) {
      foreach ($item in $s3NormalizedList) {
        $combinedS3Lists.Add($item) | Out-Null
      }
    }
  }
  
  # Create combined reports if multiple accounts processed
  if ($accountsProcessed.Count -gt 1) {
    Write-Host "Creating combined reports for all accounts..." -ForegroundColor Cyan
    
    # Combined EC2 instances
    $combinedEc2List = [System.Collections.ArrayList]::new()
    foreach ($acctInfo in $accountsProcessed) {
      $accountId = $acctInfo.AccountId
      if ($ec2ListByAccount.ContainsKey($accountId)) {
        foreach ($item in $ec2ListByAccount[$accountId]) {
          $combinedEc2List.Add($item) | Out-Null
        }
      }
    }
    
    if ($combinedEc2List.Count -gt 0) {
      $combinedOutputEc2Instance = "${baseOutputEc2Instance}-combined-$date_string.csv"
      addTagsToAllObjectsInList $combinedEc2List
      $combinedEc2List | Export-Csv -Path $combinedOutputEc2Instance -NoTypeInformation
      Write-Host "Combined EC2 CSV: $combinedOutputEc2Instance" -ForegroundColor Green
      $allOutputFiles.Add($combinedOutputEc2Instance) | Out-Null
    }
    
    # Combined unattached volumes
    $combinedEc2UnattachedVolList = [System.Collections.ArrayList]::new()
    foreach ($acctInfo in $accountsProcessed) {
      $accountId = $acctInfo.AccountId
      if ($ec2UnattachedVolListByAccount.ContainsKey($accountId)) {
        foreach ($item in $ec2UnattachedVolListByAccount[$accountId]) {
          $combinedEc2UnattachedVolList.Add($item) | Out-Null
        }
      }
    }
    
    if ($combinedEc2UnattachedVolList.Count -gt 0) {
      $combinedOutputEc2UnattachedVolume = "${baseOutputEc2UnattachedVolume}-combined-$date_string.csv"
      addTagsToAllObjectsInList $combinedEc2UnattachedVolList
      $combinedEc2UnattachedVolList | Export-Csv -Path $combinedOutputEc2UnattachedVolume -NoTypeInformation
      Write-Host "Combined unattached volumes CSV: $combinedOutputEc2UnattachedVolume" -ForegroundColor Green
      $allOutputFiles.Add($combinedOutputEc2UnattachedVolume) | Out-Null
    }
    
    # Combined S3 buckets
    if ($combinedS3Lists.Count -gt 0) {
      $combinedOutputS3 = "${baseOutputS3}-combined-$date_string.csv"
      $combinedS3Lists | Export-Csv -Path $combinedOutputS3 -NoTypeInformation
      Write-Host "Combined S3 CSV: $combinedOutputS3" -ForegroundColor Green
      $allOutputFiles.Add($combinedOutputS3) | Out-Null
    }
  }

  # Add log file to output files list
  $allOutputFiles.Add($output_log) | Out-Null
  try {
      if (Get-Module -ListAvailable -Name ImportExcel | Where-Object { $_ }) {
          Write-Host "ImportExcel module found. Proceeding with Excel summary creation..." -ForegroundColor Green

          # Force importing the module
          Import-Module ImportExcel -Force -ErrorAction Stop

          Write-Host "Creating Excel summary files..." -ForegroundColor Green

          # Create per-account Excel summaries
          foreach ($acctInfo in $accountsProcessed) {
              $accountId = $acctInfo.AccountId
              $accountAlias = $acctInfo.AccountAlias

              $ec2List = $ec2ListByAccount[$accountId]
              $ec2UnattachedVolList = $ec2UnattachedVolListByAccount[$accountId]
              $s3List = $s3ListByAccount[$accountId]

              $summaryXlsx = "${accountId}_summary_$date_string.xlsx"
              if (Test-Path $summaryXlsx) { Remove-Item $summaryXlsx -Force }

              # Calculate EC2 total metrics
              $ec2TotalCount = $ec2List.Count
              $ec2TotalGiB = ($ec2List.SizeGiB | Measure-Object -Sum).Sum
              $ec2TotalGB = ($ec2List.SizeGB | Measure-Object -Sum).Sum
              $ec2TotalTiB = ($ec2List.SizeTiB | Measure-Object -Sum).Sum
              $ec2TotalTB = ($ec2List.SizeTB | Measure-Object -Sum).Sum

              # Calculate S3 total metrics for all units
              $s3TotalCount = $s3List.Count
              $s3TotalTiB = 0
              $s3TotalGiB = 0
              $s3TotalGB = 0
              $s3TotalTB = 0
            
              if ($s3List.Count -gt 0) {
                  # Calculate S3 TiB totals
                  ($s3List | Get-Member -MemberType NoteProperty | Where-Object Name -like "*_SizeTiB") | ForEach-Object {
                      $s3TotalTiB += (($s3List.$($_.Name) | Measure-Object -Sum).Sum)
                  }
                  
                  # Calculate S3 GiB totals
                  ($s3List | Get-Member -MemberType NoteProperty | Where-Object Name -like "*_SizeGiB") | ForEach-Object {
                      $s3TotalGiB += (($s3List.$($_.Name) | Measure-Object -Sum).Sum)
                  }
                  
                  # Calculate S3 GB totals
                  ($s3List | Get-Member -MemberType NoteProperty | Where-Object Name -like "*_SizeGB") | ForEach-Object {
                      $s3TotalGB += (($s3List.$($_.Name) | Measure-Object -Sum).Sum)
                  }
                  
                  # Calculate S3 TB totals
                  ($s3List | Get-Member -MemberType NoteProperty | Where-Object Name -like "*_SizeTB") | ForEach-Object {
                      $s3TotalTB += (($s3List.$($_.Name) | Measure-Object -Sum).Sum)
                  }
              }

              # Get all regions from resources
              $allRegions = [System.Collections.ArrayList]::new()
              
              # Add regions from EC2 instances
              if ($ec2List -and $ec2List.Count -gt 0) {
                  $ec2List.Region | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | ForEach-Object {
                      if (-not $allRegions.Contains($_)) { $allRegions.Add($_) | Out-Null }
                  }
              }
              
              # Add regions from S3 buckets
              if ($s3List -and $s3List.Count -gt 0) {
                  $s3List.Region | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | ForEach-Object {
                      if (-not $allRegions.Contains($_)) { $allRegions.Add($_) | Out-Null }
                  }
              }
              
              # Sort the regions
              $allRegions = $allRegions | Sort-Object

              #---------- EC2 SUMMARY SECTION ----------
              # EC2 Overall Summary
              $ec2Summary = [PSCustomObject]@{
                  "ResourceType" = "EC2 Instances"
                  "Region"       = "All"
                  "Count"        = $ec2TotalCount
                  "Total Size (GiB)" = $ec2TotalGiB
                  "Total Size (GB)"  = $ec2TotalGB
                  "Total Size (TiB)" = $ec2TotalTiB
                  "Total Size (TB)"  = $ec2TotalTB
              }

              # EC2 Regional Breakdown
              $ec2RegionalSummary = @()
              foreach ($region in $allRegions) {
                  $ec2InRegion = $ec2List | Where-Object Region -eq $region
                  $regionEc2Count = $ec2InRegion.Count
                  if ($regionEc2Count -gt 0) {
                      $regionEc2GiB = ($ec2InRegion.SizeGiB | Measure-Object -Sum).Sum
                      $regionEc2GB = ($ec2InRegion.SizeGB | Measure-Object -Sum).Sum
                      $regionEc2TiB = ($ec2InRegion.SizeTiB | Measure-Object -Sum).Sum
                      $regionEc2TB = ($ec2InRegion.SizeTB | Measure-Object -Sum).Sum
                      
                      $ec2RegionalSummary += [PSCustomObject]@{
                          "Region"       = $region
                          "Count"        = $regionEc2Count
                          "Total Size (GiB)" = $regionEc2GiB
                          "Total Size (GB)"  = $regionEc2GB
                          "Total Size (TiB)" = $regionEc2TiB
                          "Total Size (TB)"  = $regionEc2TB
                      }
                  }
              }

              # Export EC2 Summary to Excel - both overall and regional breakdown
              $ec2Summary | Export-Excel -Path $summaryXlsx -WorksheetName "EC2 Summary" -AutoSize -FreezeTopRow -BoldTopRow
              
              # Add regional breakdown after the summary (on same sheet)
              if ($ec2RegionalSummary.Count -gt 0) {
                  $ec2RegionalSummary | Export-Excel -Path $summaryXlsx -WorksheetName "EC2 Summary" -AutoSize -FreezeTopRow -BoldTopRow -StartRow 4
              }

              #---------- S3 SUMMARY SECTION ----------
              # S3 Overall Summary
              $s3Summary = [PSCustomObject]@{
                  "ResourceType" = "S3 Buckets"
                  "Region"       = "All"
                  "Count"        = $s3TotalCount
                  "Total Size (GiB)" = [math]::Round($s3TotalGiB, 3)
                  "Total Size (GB)"  = [math]::Round($s3TotalGB, 3)
                  "Total Size (TiB)" = [math]::Round($s3TotalTiB, 4)
                  "Total Size (TB)"  = [math]::Round($s3TotalTB, 4)
              }

              # S3 Regional Breakdown
              $s3RegionalSummary = @()
              foreach ($region in $allRegions) {
                  $s3InRegion = $s3List | Where-Object Region -eq $region
                  $regionS3Count = ($s3InRegion | Measure-Object).Count
                  
                  if ($regionS3Count -gt 0) {
                      # Calculate total S3 size in all units for this region
                      $regionS3GiB = 0
                      $regionS3GB = 0
                      $regionS3TiB = 0
                      $regionS3TB = 0
                      
                      if ($s3InRegion -and ($s3InRegion | Measure-Object).Count -gt 0) {
                          # Find all size properties by examining property names
                          $giBProperties = $s3InRegion[0].PSObject.Properties.Name | Where-Object { $_ -like '*_SizeGiB' }
                          $gbProperties = $s3InRegion[0].PSObject.Properties.Name | Where-Object { $_ -like '*_SizeGB' }
                          $tiBProperties = $s3InRegion[0].PSObject.Properties.Name | Where-Object { $_ -like '*_SizeTiB' }
                          $tbProperties = $s3InRegion[0].PSObject.Properties.Name | Where-Object { $_ -like '*_SizeTB' }
                          
                          # Sum all the GiB sizes
                          foreach ($propName in $giBProperties) {
                              $s3InRegion | ForEach-Object {
                                  if ($_.$propName) {
                                      $regionS3GiB += $_.$propName
                                  }
                              }
                          }
                          
                          # Sum all the GB sizes
                          foreach ($propName in $gbProperties) {
                              $s3InRegion | ForEach-Object {
                                  if ($_.$propName) {
                                      $regionS3GB += $_.$propName
                                  }
                              }
                          }
                          
                          # Sum all the TiB sizes
                          foreach ($propName in $tiBProperties) {
                              $s3InRegion | ForEach-Object {
                                  if ($_.$propName) {
                                      $regionS3TiB += $_.$propName
                                  }
                              }
                          }
                          
                          # Sum all the TB sizes
                          foreach ($propName in $tbProperties) {
                              $s3InRegion | ForEach-Object {
                                  if ($_.$propName) {
                                      $regionS3TB += $_.$propName
                                  }
                              }
                          }
                      }
                      
                      $s3RegionalSummary += [PSCustomObject]@{
                          "Region"           = $region
                          "Count"            = $regionS3Count
                          "Total Size (GiB)" = [math]::Round($regionS3GiB, 3)
                          "Total Size (GB)"  = [math]::Round($regionS3GB, 3)
                          "Total Size (TiB)" = [math]::Round($regionS3TiB, 4)
                          "Total Size (TB)"  = [math]::Round($regionS3TB, 4)
                      }
                  }
              }

              # Export S3 Summary to Excel - both overall and regional breakdown
              $s3Summary | Export-Excel -Path $summaryXlsx -WorksheetName "S3 Summary" -AutoSize -FreezeTopRow -BoldTopRow
              
              # Add regional breakdown after the summary (on same sheet)
              if ($s3RegionalSummary.Count -gt 0) {
                  $s3RegionalSummary | Export-Excel -Path $summaryXlsx -WorksheetName "S3 Summary" -AutoSize -FreezeTopRow -BoldTopRow -StartRow 4
              }

              # Add detailed resource information worksheets
              if ($ec2List.Count -gt 0) {
                  $ec2List | Export-Excel -Path $summaryXlsx -WorksheetName "EC2 Details" -AutoSize -FreezeTopRow -BoldTopRow
              }
              
              if ($s3List.Count -gt 0) {
                  $s3List | Export-Excel -Path $summaryXlsx -WorksheetName "S3 Details" -AutoSize -FreezeTopRow -BoldTopRow
              }

              # Add unattached volumes worksheet if applicable
              if ($ec2UnattachedVolList.Count -gt 0) {
                  $ec2UnattachedVolList | Export-Excel -Path $summaryXlsx -WorksheetName "Unattached Volumes" -AutoSize -FreezeTopRow -BoldTopRow
              }

              Write-Host "Excel summary for account ${accountId}: ${summaryXlsx}" -ForegroundColor Green
              
              # Add to list of output files
              $allOutputFiles.Add($summaryXlsx) | Out-Null
          }

          # Create combined Excel summary if multiple accounts processed
          if ($accountsProcessed.Count -gt 1) {
              $combinedSummaryXlsx = "combined_summary_$date_string.xlsx"
              if (Test-Path $combinedSummaryXlsx) { Remove-Item $combinedSummaryXlsx -Force }
              
              # Combine all EC2 instances
              $combinedEc2List = [System.Collections.ArrayList]::new()
              foreach ($acctInfo in $accountsProcessed) {
                  $accountId = $acctInfo.AccountId
                  if ($ec2ListByAccount.ContainsKey($accountId)) {
                      foreach ($item in $ec2ListByAccount[$accountId]) {
                          $combinedEc2List.Add($item) | Out-Null
                      }
                  }
              }
              
              # Combine all unattached volumes
              $combinedEc2UnattachedVolList = [System.Collections.ArrayList]::new()
              foreach ($acctInfo in $accountsProcessed) {
                  $accountId = $acctInfo.AccountId
                  if ($ec2UnattachedVolListByAccount.ContainsKey($accountId)) {
                      foreach ($item in $ec2UnattachedVolListByAccount[$accountId]) {
                          $combinedEc2UnattachedVolList.Add($item) | Out-Null
                      }
                  }
              }
              
              # Combined S3 buckets
              $combinedS3List = [System.Collections.ArrayList]::new()
              foreach ($acctInfo in $accountsProcessed) {
                  $accountId = $acctInfo.AccountId
                  if ($s3ListByAccount.ContainsKey($accountId)) {
                      foreach ($item in $s3ListByAccount[$accountId]) {
                          $combinedS3List.Add($item) | Out-Null
                      }
                  }
              }
              
              # Export all combined data to Excel
              if ($combinedEc2List.Count -gt 0) {
                  addTagsToAllObjectsInList $combinedEc2List
                  $combinedEc2List | Export-Excel -Path $combinedSummaryXlsx -WorksheetName "EC2 Instances" -AutoSize -FreezeTopRow -BoldTopRow
              }
              
              if ($combinedEc2UnattachedVolList.Count -gt 0) {
                  addTagsToAllObjectsInList $combinedEc2UnattachedVolList
                  $combinedEc2UnattachedVolList | Export-Excel -Path $combinedSummaryXlsx -WorksheetName "Unattached Volumes" -AutoSize -FreezeTopRow -BoldTopRow
              }
              
              if ($combinedS3List.Count -gt 0) {
                  $combinedS3List | Export-Excel -Path $combinedSummaryXlsx -WorksheetName "S3 Buckets" -AutoSize -FreezeTopRow -BoldTopRow
              }
              
              Write-Host "Combined Excel summary: $combinedSummaryXlsx" -ForegroundColor Green
              $allOutputFiles.Add($combinedSummaryXlsx) | Out-Null
          }
      } else {
          Write-Host "ImportExcel module not installed; skipping Excel summary." -ForegroundColor Yellow
      }
  } catch {
      Write-Host "Excel summary creation failed: $_" -ForegroundColor Yellow
  }

} catch {
  # Handle any unhandled exceptions
  Write-Error "Script error:"
  Write-Error $_
  Write-Error $_.ScriptStackTrace
} finally {
  # Ensure transcript is stopped even if script fails
  Stop-Transcript
}

# Archive all result files into a single ZIP file
$existingFiles = $allOutputFiles | Where-Object { Test-Path $_ }
if ($existingFiles.Count -gt 0) {
  Compress-Archive -Path $existingFiles -DestinationPath $archiveFile -Force
  foreach($f in $allOutputFiles){ Remove-Item -Path $f -ErrorAction SilentlyContinue }
  Write-Host "Results compressed: $archiveFile" -ForegroundColor Green
} else {
  Write-Host "No output files to archive." -ForegroundColor Yellow
}

[System.Threading.Thread]::CurrentThread.CurrentCulture = $CurrentCulture
[System.Threading.Thread]::CurrentThread.CurrentUICulture = $CurrentCulture
Write-Host "`nInventory complete. Results in $archiveFile." -ForegroundColor Cyan
Write-Host "All output files have been compressed into the ZIP archive." -ForegroundColor Cyan
