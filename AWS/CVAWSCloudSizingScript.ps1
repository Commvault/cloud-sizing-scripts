#requires -Version 7.0

<#
.SYNOPSIS
    AWS Cloud Sizing Script – Comprehensive inventory and sizing analysis for compute, storage, and databases.

.DESCRIPTION
    Inventories AWS workloads across one or multiple accounts and regions:
      - EC2 instances (with attached/unattached EBS volumes)
      - S3 buckets and storage metrics
      - Elastic File System (EFS)
      - FSx file systems (ONTAP/SVX, Windows, Lustre, etc.)
      - RDS database instances
      - DynamoDB tables
      - Redshift clusters
      - EKS Clusters and PVCs associated with the cluster
    Supports multiple authentication methods including:
      - Access/Secret Key authentication (via Creds.txt or ProfileLocation)
      - Default AWS CLI profile (IAM role assumed automatically in CloudShell/EC2)
      - Cross-account role assumption
    Generates detailed Excel reports with both workload summaries and details,
    at the account level and across all accounts.
    Includes hierarchical progress tracking and detailed logging.
    Outputs timestamped Excel files and creates a ZIP archive of all results.

.PARAMETER DefaultProfile
    Use the default AWS CLI profile (IAM role) for authentication.
    Typically used when running inside AWS CloudShell or an EC2 instance.

.PARAMETER UserSpecifiedProfileNames
    Comma-separated list of AWS CLI profile names defined inside a credentials file (Creds.txt).

.PARAMETER AllLocalProfiles
    Use all profiles defined inside a credentials file.

.PARAMETER ProfileLocation
    Path to a shared credentials file (e.g., ./Creds.txt).
    Required for access/secret key authentication outside AWS (e.g., Laptop/Desktop).

.PARAMETER CrossAccountRoleName
    Name of the IAM role to assume for cross-account access.
    Must be executed from AWS environments (EC2/CloudShell) or with valid base IAM role.

.PARAMETER UserSpecifiedAccounts
    Comma-separated list of AWS account IDs to use with the specified cross-account role.

.PARAMETER UserSpecifiedAccountsFile
    Path to a file containing AWS account IDs (one per line) to use with the specified cross-account role.

.PARAMETER Regions
    Comma-separated list of AWS regions to query.

.PARAMETER ExternalId
    External ID required for cross-account role assumption (Optional).

.OUTPUTS
    Creates a timestamped output directory with the following files:

    - <AccountID>_summary_YYYY-MM-DD_HHMMSS.xlsx
        – Account-level report containing:
            • Workload-specific summaries (EC2, S3, RDS, EFS, FSx, DynamoDB, Redshift)
            • Workload-specific detailed sheets for each service

    - comprehensive_all_aws_accounts_summary_YYYY-MM-DD_HHMMSS.xlsx
        – Consolidated report across all AWS accounts with the same summary + details structure

    - aws_sizing_script_output_YYYY-MM-DD_HHMMSS.log
        – Complete execution log

    - aws_sizing_results_YYYY-MM-DD_HHMMSS.zip
        – ZIP archive containing all per-account summary files, the comprehensive report, and the log
#>


[CmdletBinding(DefaultParameterSetName = 'DefaultProfile')]
param (
    [Parameter(ParameterSetName='AllLocalProfiles',Mandatory=$true)]
    [ValidateNotNullOrEmpty()][switch]$AllLocalProfiles,

    [Parameter(ParameterSetName='CrossAccountRole',Mandatory=$true)]
    [ValidateNotNullOrEmpty()][string]$CrossAccountRoleName,

    [Parameter(ParameterSetName='CrossAccountRole')]
    [string]$CrossAccountRoleSessionName = "CVAWS-Cost-Sizing",

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
    [switch]$DebugBucketTags,
    [switch]$SelectiveZipping = $true
)


[System.Threading.Thread]::CurrentThread.CurrentCulture = 'en-US'
[System.Threading.Thread]::CurrentThread.CurrentUICulture = 'en-US'

Add-Type -AssemblyName System.IO.Compression.FileSystem

$script:startTime = Get-Date
$script:currentActivity = ""
$script:moduleImportError = $false
$script:KubectlInstalled = $false

# Service configuration registry
$script:ServiceRegistry = @{
    EC2 = @{
        DisplayName = "Instances"
        GetCommand = "Get-EC2Instance"
        PropertyPath = "Instances"
        SizeProperty = "SizeGiB"
        IdProperty = "InstanceId"
        StorageType = "EBS"
        ProcessorFunction = "Process-EC2Instance"
        RequiresTagProcessing = $true
    }
    S3 = @{
        DisplayName = "Buckets"
        GetCommand = "Get-S3Bucket"
        PropertyPath = $null
        SizeProperty = "Custom"
        IdProperty = "BucketName"
        StorageType = "Object"
        ProcessorFunction = "Process-S3Bucket"
        RequiresTagProcessing = $true
        SpecialHandling = $true
    }
    EFS = @{
        DisplayName = "File Systems"
        GetCommand = "Get-EFSFileSystem"
        PropertyPath = $null
        SizeProperty = "SizeGiB"
        IdProperty = "FileSystemId"
        StorageType = "NFS"
        ProcessorFunction = "Process-EFSFileSystem"
        RequiresTagProcessing = $true
    }
    FSX = @{
        DisplayName = "File Systems"
        GetCommand = "Get-FSXFileSystem"
        PropertyPath = $null
        SizeProperty = "StorageCapacityGiB"
        IdProperty = "FileSystemId"
        StorageType = "Managed"
        ProcessorFunction = "Process-FSXFileSystem"
        RequiresTagProcessing = $true
        HasTypeBreakdown = $true
    }
    FSX_SVM = @{
        DisplayName = "ONTAP SVMs"
        GetCommand = "Get-FSXStorageVirtualMachine"
        PropertyPath = $null
        SizeProperty = "Custom"
        IdProperty = "StorageVirtualMachineId"
        StorageType = "Metadata"
        ProcessorFunction = "Process-FSXStorageVirtualMachine"
        RequiresTagProcessing = $true
        SpecialHandling = $true
    }
EKS = @{
    DisplayName = "Kubernetes"
    GetCommand = "Get-AllEKSClusters"
    CandidateGetCommands = @("Get-AllEKSClusters","Get-EKSClusterList")
    PropertyPath = $null
    SizeProperty = "Custom"
    IdProperty = "Name"
    StorageType = "Block"
    ProcessorFunction = "Process-EKSCluster"
    RequiresTagProcessing = $false
    SpecialHandling = $true
}

    UnattachedVolumes = @{
        DisplayName = "Volumes"
        GetCommand = "Get-EC2Volume"
        PropertyPath = $null
        SizeProperty = "SizeGiB"
        IdProperty = "VolumeId"
        StorageType = "EBS"
        ProcessorFunction = "Process-UnattachedVolume"
        RequiresTagProcessing = $true
        Filter = @{Name='status';Values='available'}
    }
    RDS = @{
        DisplayName = "Databases"
        GetCommand = "Get-RDSDBInstance"
        CandidateGetCommands = @("Get-RDSDBInstance","Get-RDSInstance","Get-RDSDBInstances","Get-RDSDBInstanceList")
        PropertyPath = "DBInstances"
        SizeProperty = "AllocatedStorageGB"
        IdProperty = "DBInstanceIdentifier"
        StorageType = "Block"
        ProcessorFunction = "Process-RDSInstance"
        RequiresTagProcessing = $true
    }
    DynamoDB = @{
        DisplayName = "Tables"
        GetCommand = "Get-DDBTableList"
        CandidateGetCommands = @("Get-DDBTableList","Get-DDBTable","Get-DDBTables","Get-DDBTableNames")
        PropertyPath = "TableNames"
        SizeProperty = "TableSizeBytes"
        IdProperty = "TableName"
        StorageType = "NoSQL"
        ProcessorFunction = "Process-DynamoDBTable"
        RequiresTagProcessing = $true
    }
    Redshift = @{
        DisplayName = "Clusters"
        GetCommand = "Get-RSCluster"
        PropertyPath = "Clusters"
        SizeProperty = "Custom"
        IdProperty = "ClusterIdentifier"
        StorageType = "Columnar"
        ProcessorFunction = "Process-RedshiftCluster"
        RequiresTagProcessing = $true
        SpecialHandling = $true
    }
}

$script:Config = @{
    DefaultQueryRegion = if ($RegionToQuery) { $RegionToQuery } else { "us-east-1" }
    DefaultGovCloudQueryRegion = "us-gov-west-1"
    Partition = if ($Partition) { $Partition } else { "Standard" }
    ProfileLocation = if ($ProfileLocation) { @{ProfileLocation = $ProfileLocation} } else { @{} }
    OutputPath = (Get-Location).Path
    SkipBucketTags = $SkipBucketTags.IsPresent
    DebugBucketTags = $DebugBucketTags.IsPresent
}

$date = Get-Date
$date_string = $date.ToString("yyyy-MM-dd_HHmmss")

$script:LogFile = "aws_sizing_script_output_$date_string.log"
$script:LogPath = Join-Path $script:Config.OutputPath $script:LogFile

$utcEndTime = $date.ToUniversalTime()
$utcStartTime = $utcEndTime.AddDays(-7)

$script:ServiceDataByAccount = @{}
$script:AccountsProcessed = @()
$script:AllOutputFiles = [System.Collections.ArrayList]::new()

$baseOutputEc2Instance = "aws_ec2_instance_info"
$baseOutputEc2UnattachedVolume = "aws_ec2_unattached_volume_info"
$baseOutputS3 = "aws_s3_info"
$baseOutputEFS = "aws_efs_info"
$baseOutputFSx = "aws_fsx_info"
$baseOutputFSxSVM = "aws_fsx_ontap_svms_info"
$baseOutputRDS = "aws_rds_info"
$baseOutputDynamoDB = "aws_dynamodb_info"
$baseOutputRedshift = "aws_redshift_info"
$baseOutputEKS = "aws_eks_info"
$archiveFile = "aws_sizing_results_$date_string.zip"

function Show-ScriptProgress {
    param(
        [string]$Activity,
        [string]$Status = "",
        [int]$PercentComplete = 0
    )

    if (-not $script:LastProgressUpdate) { $script:LastProgressUpdate = Get-Date }
    $now = Get-Date
    $elapsed = $now - $script:LastProgressUpdate

    if ($elapsed.TotalMilliseconds -lt 300 -and $PercentComplete -lt 100) {
        return
    }
    $script:LastProgressUpdate = $now

    $script:currentActivity = $Activity
    Write-Progress -Id 1 -Activity "AWS Cost Sizing Analysis" -Status "$Activity - $Status" -PercentComplete $PercentComplete
}

function Get-SafeCWMetricStatistic {
    param(
        [string]$Namespace,
        [string]$MetricName,
        [object]$Dimensions,
        [int]$Period = 300,
        [string[]]$Statistics = "Average",
        [datetime]$StartTime,
        [datetime]$EndTime,
        [object]$Credential,
        [string]$Region
    )

    if (-not $EndTime)   { $EndTime = if ($script:utcEndTime) { [datetime]$script:utcEndTime } else { (Get-Date).ToUniversalTime() } }
    if (-not $StartTime) { $StartTime = if ($script:utcStartTime) { [datetime]$script:utcStartTime } else { $EndTime.AddDays(-7) } }

    $utcStart = $StartTime.ToUniversalTime()
    $utcEnd   = $EndTime.ToUniversalTime()

    if ($Dimensions -and -not ($Dimensions -is [System.Collections.IEnumerable] -and -not ($Dimensions -is [string]))) {
        $Dimensions = @($Dimensions)
    }

    $cwParams = @{
        Namespace      = $Namespace
        MetricName     = $MetricName
        Dimension      = $Dimensions
        StartTime      = $StartTime
        EndTime        = $EndTime
        UtcStartTime   = $utcStart
        UtcEndTime     = $utcEnd
        Period         = $Period
        Statistic      = $Statistics
        Credential     = $Credential
        Region         = $Region
        ErrorAction    = 'SilentlyContinue'
        WarningAction  = 'SilentlyContinue'

    }

    return Get-CWMetricStatistic @cwParams
}

function Write-ScriptOutput {
    param(
        [string]$Message,
        [ValidateSet("Info", "Warning", "Error", "Success")]
        [string]$Level = "Info"
    )

    $colors = @{
        Info = "White"
        Warning = "Yellow"
        Error = "Red"
        Success = "Green"
    }

    $timestamp = (Get-Date).ToString("HH:mm:ss")
    $logTimestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    $consoleMessage = "[$timestamp] $Message"
    $logMessage = "[$logTimestamp] [$Level] $Message"

    Write-Host $consoleMessage -ForegroundColor $colors[$Level]

    try {
        $logMessage | Out-File -FilePath $script:LogPath -Append -Encoding UTF8
    }
    catch {
        Write-Host "[$timestamp] Warning: Failed to write to log file: $_" -ForegroundColor Yellow
    }
}

function Convert-BytesToSizes {
    param([double]$Bytes)

    if ($Bytes -eq 0) {
        return @{
            SizeGiB = 0; SizeTiB = 0; SizeGB = 0; SizeTB = 0
        }
    }

    return @{
        SizeGiB = $Bytes / 1GB
        SizeTiB = $Bytes / 1TB
        SizeGB = $Bytes / 1000000000
        SizeTB = $Bytes / 1000000000000
    }
}

function Add-TagProperties {
    param([object]$Object, [array]$Tags)

    if ($Tags) {
        foreach ($tag in $Tags) {
            $tagName = "Tag_$($tag.Key -replace '[^a-zA-Z0-9_]', '_')"
            Add-Member -InputObject $Object -NotePropertyName $tagName -NotePropertyValue $tag.Value -Force
        }
    }
}

function Initialize-ServiceCollections {
    param([string]$AccountId)

    if (-not $script:ServiceDataByAccount.ContainsKey($AccountId)) {
        $script:ServiceDataByAccount[$AccountId] = @{}

        foreach ($serviceName in $script:ServiceRegistry.Keys) {
            $script:ServiceDataByAccount[$AccountId][$serviceName] = [System.Collections.ArrayList]@()
        }
    }
}

function Invoke-ServiceInventory {
    param(
        [string]$ServiceName,
        [object]$Credential,
        [string]$Region,
        [object]$AccountInfo,
        [string]$AccountAlias
    )

    $serviceConfig = $script:ServiceRegistry[$ServiceName]
    if (-not $serviceConfig) {
        Write-ScriptOutput "Unknown service: $ServiceName" -Level Warning
        return
    }

    try {
        Write-ScriptOutput "$ServiceName inventory started for region $Region" -Level Info

        if (-not (Get-Command $serviceConfig.GetCommand -ErrorAction SilentlyContinue)) {
            Write-ScriptOutput "Command $($serviceConfig.GetCommand) not available. Install the appropriate AWS Tools module for $ServiceName." -Level Warning
            return
        }

        $getParams = @{
            Credential = $Credential
            Region = $Region
            ErrorAction = 'Stop'
        }

        if ($ServiceName -eq "S3") { $getParams.BucketRegion = $Region }
        if ($ServiceName -eq "UnattachedVolumes" -and $serviceConfig.Filter) {
            $getParams.Filter = $serviceConfig.Filter
        }

        $resolvedCommand = $null
        if ($serviceConfig.CandidateGetCommands) {
            foreach ($cand in $serviceConfig.CandidateGetCommands) {
                if (Get-Command -Name $cand -ErrorAction SilentlyContinue) { $resolvedCommand = $cand; break }
            }
        }
        if (-not $resolvedCommand -and $serviceConfig.GetCommand) {
            if (Get-Command -Name $serviceConfig.GetCommand -ErrorAction SilentlyContinue) { $resolvedCommand = $serviceConfig.GetCommand }
        }

        if (-not $resolvedCommand) {
            Write-ScriptOutput "Command $($serviceConfig.GetCommand) not available. Install the appropriate AWS Tools module for $ServiceName." -Level Warning
            return
        }

        Write-ScriptOutput "Using command $resolvedCommand for service $ServiceName" -Level Info
        $serviceItems = & $resolvedCommand @getParams

        if ($serviceConfig.PropertyPath -and $serviceItems) {
            $prop = $serviceConfig.PropertyPath
            try {
                if ($serviceItems -and $serviceItems.PSObject -and ($serviceItems.PSObject.Properties.Name -contains $prop)) {
                    $extracted = $serviceItems.($prop)
                    if ($extracted -is [System.Collections.IEnumerable] -and -not ($extracted -is [string])) { $serviceItems = $extracted } else { $serviceItems = @($extracted) }
                }
                elseif ($serviceItems -is [System.Collections.IEnumerable] -and -not ($serviceItems -is [string]) -and $serviceItems.Count -gt 0 -and ($serviceItems[0].PSObject -and ($serviceItems[0].PSObject.Properties.Name -contains $prop))) {
                    $serviceItems = $serviceItems | ForEach-Object { $_.($prop) } | Where-Object { $_ -ne $null }
                }
            } catch {
            }
        }

        if ($serviceItems -ne $null -and -not ($serviceItems -is [System.Collections.IEnumerable] -and -not ($serviceItems -is [string]))) {
            $serviceItems = @($serviceItems)
        }

        if (-not $serviceItems) {
            Write-ScriptOutput "No $ServiceName found in region $Region" -Level Info
            return
        }

        $serviceList = $script:ServiceDataByAccount[$AccountInfo.Account][$ServiceName]
        $count = 1

        foreach ($item in $serviceItems) {
            $itemId = $item.($serviceConfig.IdProperty)
            Show-ScriptProgress -Activity "$ServiceName $itemId" -Status "$count / $($serviceItems.Count)" -PercentComplete (($count / $serviceItems.Count) * 100)

            $processedItem = & $serviceConfig.ProcessorFunction -Item $item -Credential $Credential -Region $Region -AccountInfo $AccountInfo -AccountAlias $AccountAlias

            if ($processedItem) {
                $serviceList.Add($processedItem) | Out-Null
            }
            $count++
        }

        Write-ScriptOutput "$ServiceName done" -Level Success
    }
    catch {
        Write-ScriptOutput "Failed to get $ServiceName for region $Region`: $_" -Level Error
    }
}

function Process-EC2Instance {
    param($Item, $Credential, $Region, $AccountInfo, $AccountAlias)

    $ebsVolumes = @()
    $totalEbsBytes = 0

    foreach ($blockDevice in $Item.BlockDeviceMappings) {
        if ($blockDevice.Ebs) {
            try {
                $volume = Get-EC2Volume -VolumeId $blockDevice.Ebs.VolumeId -Credential $Credential -Region $Region -ErrorAction Stop
                $ebsVolumes += $volume
                $totalEbsBytes += $volume.Size * 1GB
            }
            catch {
                Write-ScriptOutput "Failed to get volume $($blockDevice.Ebs.VolumeId): $($_.Exception.Message)" -Level Warning
            }
        }
    }

    $instanceTags = @()
    try {
        $tagFilter = @{
            Name = "resource-id"
            Values = @($Item.InstanceId)
        }
        $tagResponse = Get-EC2Tag -Filter $tagFilter -Credential $Credential -Region $Region -ErrorAction Stop
        $instanceTags = $tagResponse

        if ($instanceTags -and $instanceTags.Count -gt 0) {
            Write-ScriptOutput "Retrieved $($instanceTags.Count) tags for instance $($Item.InstanceId)" -Level Info
        }
    }
    catch {
        try {
            $instanceDetail = Get-EC2Instance -InstanceId $Item.InstanceId -Credential $Credential -Region $Region -ErrorAction Stop
            if ($instanceDetail -and $instanceDetail.Instances -and $instanceDetail.Instances[0].Tags) {
                $instanceTags = $instanceDetail.Instances[0].Tags
                Write-ScriptOutput "Retrieved tags via instance detail for $($Item.InstanceId)" -Level Info
            }
        }
        catch {
            Write-ScriptOutput "Failed to get tags for instance $($Item.InstanceId): $($_.Exception.Message)" -Level Warning
        }
    }

    $sizes = Convert-BytesToSizes -Bytes $totalEbsBytes

        $ec2Obj = [PSCustomObject]@{
            AwsAccountId = "`u{200B}$($AccountInfo.Account)"
        AwsAccountAlias = $AccountAlias
        Region = $Region
        InstanceId = $Item.InstanceId
        InstanceType = $Item.InstanceType
        State = $Item.State.Name
        LaunchTime = $Item.LaunchTime
        VolumeCount = $ebsVolumes.Count
        SizeGiB = $sizes.SizeGiB
        SizeTiB = $sizes.SizeTiB
        SizeGB = $sizes.SizeGB
        SizeTB = $sizes.SizeTB
        VolumeDetails = ($ebsVolumes | ForEach-Object { "$($_.VolumeId):$($_.Size)GB:$($_.VolumeType)" }) -join ";"
    }

    Add-TagProperties -Object $ec2Obj -Tags $instanceTags

    return $ec2Obj
}

function Process-S3Bucket {
    param($Item, $Credential, $Region, $AccountInfo, $AccountAlias)

    try {
        $bucketName = if ($Item.BucketName) { $Item.BucketName } elseif ($Item.Name) { $Item.Name } else { $null }
        if (-not $bucketName) { return $null }

        try {
            $bucketLocation = Get-S3BucketLocation -BucketName $bucketName -Credential $Credential -ErrorAction Stop
            $actualRegion = if ($bucketLocation.Value -eq "" -or -not $bucketLocation.Value) { "us-east-1" } else { $bucketLocation.Value }
        }
        catch {
            Write-ScriptOutput "Warning: could not determine region for bucket $bucketName, using passed region '$Region' : $_" -Level Warning
            $actualRegion = if ($Region) { $Region } else { "us-east-1" }
        }

        $storageClasses = @("STANDARD", "STANDARD_IA", "ONEZONE_IA", "REDUCED_REDUNDANCY", "GLACIER", "DEEP_ARCHIVE", "INTELLIGENT_TIERING")
        $classBytes = @{}
        $totalBytes = 0.0

        foreach ($storageClass in $storageClasses) {
            try {
                $bytes = Get-S3BucketSizeByStorageClass -BucketName $bucketName -StorageClass $storageClass -Credential $Credential -Region $actualRegion
                $bytes = if ($bytes) { [double]$bytes } else { 0.0 }
            } catch {
                Write-ScriptOutput "Warning: failed to get metric for storage class $storageClass on bucket ${bucketName}: $_" -Level Warning
                $bytes = 0.0
            }
            $classBytes[$storageClass] = $bytes
            $totalBytes += $bytes

            if ($script:Config.DebugBucketTags) {
                $classSizes = Convert-BytesToSizes -Bytes $bytes
                Write-ScriptOutput ("S3 metric: Bucket='{0}' StorageClass={1} Bytes={2} GiB(binary)={3:N4} GB(decimal)={4:N4}" -f $bucketName, $storageClass, $bytes, $classSizes.SizeGiB, $classSizes.SizeGB) -Level Info
            }
        }


        if ($totalBytes -eq 0) {    try {
                $val = Get-S3BucketSize -BucketName $bucketName -Credential $Credential -Region $actualRegion
                if ($val -and $val -gt 0) {
                    $totalBytes = [double]$val
                    if ($script:Config.DebugBucketTags) {
                        $sizes = Convert-BytesToSizes -Bytes $totalBytes
                        Write-ScriptOutput ("S3 metric fallback (AllStorageTypes): Bucket='{0}' Bytes={1} GiB(binary)={2:N4} GB(decimal)={3:N4}" -f $bucketName, $totalBytes, $sizes.SizeGiB, $sizes.SizeGB) -Level Info
                    }
                }
            } catch {
            }
        }
        if ($totalBytes -eq 0) {
            try {
                $accurate = Get-S3BucketSizeAccurate -BucketName $bucketName -Credential $Credential -Region $actualRegion
                if ($accurate -and $accurate -gt 0) {
                    $totalBytes = [double]$accurate
                    if ($script:Config.DebugBucketTags) {
                        $sizes = Convert-BytesToSizes -Bytes $totalBytes
                        Write-ScriptOutput ("S3 enumeration: Bucket='{0}' Bytes={1} GiB(binary)={2:N4} GB(decimal)={3:N4}" -f $bucketName, $totalBytes, $sizes.SizeGiB, $sizes.SizeGB) -Level Info
                    }
                } else {
                    Write-ScriptOutput "S3 enumeration for bucket '$bucketName'" -Level Info
                }
            } catch {

            }
        }

        $bucketTags = @()
        if (-not $script:Config.SkipBucketTags) {
            try {
                $taggingResponse = Get-S3BucketTagging -BucketName $bucketName -Credential $Credential -Region $actualRegion -ErrorAction Stop
                $bucketTags = if ($taggingResponse.TagSet) { $taggingResponse.TagSet } elseif ($taggingResponse -is [array]) { $taggingResponse } else { $taggingResponse }
                if ($script:Config.DebugBucketTags) {
                    Write-ScriptOutput "Retrieved $($bucketTags.Count) tags for bucket $bucketName (region $actualRegion)" -Level Info
                }
            } catch {
                if ($script:Config.DebugBucketTags) {
                    Write-ScriptOutput "No tags for bucket $bucketName (region $actualRegion): $_" -Level Info
                }
                $bucketTags = @()
            }
        }

        $sizes = Convert-BytesToSizes -Bytes $totalBytes

        $s3Obj = [PSCustomObject]@{
            AwsAccountId = "`u{200B}$($AccountInfo.Account)"
            AwsAccountAlias = $AccountAlias
            Region = $actualRegion
            BucketName = $bucketName
            CreationDate = if ($Item.CreationDate) { $Item.CreationDate } else { $null }
            SizeGiB = $sizes.SizeGiB
            SizeTiB = $sizes.SizeTiB
            SizeGB = $sizes.SizeGB
            SizeTB = $sizes.SizeTB
        }

        foreach ($storageClass in $storageClasses) {
            $classSizeBytes = if ($classBytes.ContainsKey($storageClass)) { $classBytes[$storageClass] } else { 0.0 }
            $classSizes = Convert-BytesToSizes -Bytes $classSizeBytes

            Add-Member -InputObject $s3Obj -NotePropertyName "${storageClass}_SizeGiB" -NotePropertyValue $classSizes.SizeGiB -Force
            Add-Member -InputObject $s3Obj -NotePropertyName "${storageClass}_SizeGB" -NotePropertyValue $classSizes.SizeGB -Force
            Add-Member -InputObject $s3Obj -NotePropertyName "${storageClass}_SizeTiB" -NotePropertyValue $classSizes.SizeTiB -Force
            Add-Member -InputObject $s3Obj -NotePropertyName "${storageClass}_SizeTB" -NotePropertyValue $classSizes.SizeTB -Force
        }

        Add-TagProperties -Object $s3Obj -Tags $bucketTags

        return $s3Obj
    }
    catch {
        Write-ScriptOutput "Failed to process S3 bucket $($Item.BucketName) : $_" -Level Warning
        return $null
    }
}

function Process-EFSFileSystem {
    param($Item, $Credential, $Region, $AccountInfo, $AccountAlias)

    $efsSizeBytes = Get-EFSFileSystemSize -EFS $Item -Credential $Credential -Region $Region

    $efsTags = @()
    try {
        if (Get-Command "Get-EFSResourceTag" -ErrorAction SilentlyContinue) {
            $efsTags = Get-EFSResourceTag -ResourceId $Item.FileSystemArn -Credential $Credential -Region $Region -ErrorAction Stop
        } else {
            $efsTags = Get-EFSTag -FileSystemId $Item.FileSystemId -Credential $Credential -Region $Region -ErrorAction Stop -WarningAction SilentlyContinue
        }
    }
    catch {
        Write-ScriptOutput "Failed to get tags for EFS $($Item.FileSystemId)" -Level Warning
    }

    $sizes = Convert-BytesToSizes -Bytes $efsSizeBytes

        $efsObj = [PSCustomObject]@{
            AwsAccountId = "`u{200B}$($AccountInfo.Account)"
        AwsAccountAlias = $AccountAlias
        Region = $Region
        FileSystemId = $Item.FileSystemId
        CreationTime = $Item.CreationTime
        PerformanceMode = $Item.PerformanceMode
        State = $Item.LifeCycleState
        SizeGiB = $sizes.SizeGiB
        SizeTiB = $sizes.SizeTiB
        SizeGB = $sizes.SizeGB
        SizeTB = $sizes.SizeTB
    }

    Add-TagProperties -Object $efsObj -Tags $efsTags

    return $efsObj
}

function Get-EmptyRow {
    return [PSCustomObject]@{
        "ResourceType" = ""
        "Region" = ""
        "Count" = ""
        "Total Size (GiB)" = ""
        "Total Size (GB)" = ""
        "Total Size (TiB)" = ""
        "Total Size (TB)" = ""
    }
}

function Process-FSXFileSystem {
    param($Item, $Credential, $Region, $AccountInfo, $AccountAlias)

    $fsxTags = @()
    try {
        if ($Item.ResourceARN -and $Item.ResourceARN.Trim() -ne "") {
            if (Get-Command "Get-FSXResourceTag" -ErrorAction SilentlyContinue) {
                $fsxTags = Get-FSXResourceTag -ResourceArn $Item.ResourceARN -Credential $Credential -Region $Region -ErrorAction Stop
                Write-ScriptOutput "Retrieved $($fsxTags.Count) tags for FSx $($Item.FileSystemId) via ResourceTag API" -Level Info
            } else {
                Write-ScriptOutput "Get-FSXResourceTag command not available for FSx $($Item.FileSystemId)" -Level Warning
            }
        } else {
            Write-ScriptOutput "No ResourceARN available for FSx $($Item.FileSystemId), skipping tag retrieval" -Level Warning
        }
    }
    catch {
        Write-ScriptOutput "Failed to get tags for FSx $($Item.FileSystemId): $($_.Exception.Message)" -Level Warning
        try {
            if ($Item.Tags -and $Item.Tags.Count -gt 0) {
                $fsxTags = $Item.Tags
                Write-ScriptOutput "Retrieved $($fsxTags.Count) tags for FSx $($Item.FileSystemId) from item properties" -Level Info
            }
        }
        catch {
            Write-ScriptOutput "Alternative tag retrieval also failed for FSx $($Item.FileSystemId)" -Level Warning
        }
    }

    $capacityBytes = $Item.StorageCapacity * 1GB
    $sizes = Convert-BytesToSizes -Bytes $capacityBytes

    $fsxObj = [PSCustomObject]@{
        AwsAccountId = "`u{200B}$($AccountInfo.Account)"
        AwsAccountAlias = $AccountAlias
        Region = $Region
        FileSystemId = $Item.FileSystemId
        Type = $Item.FileSystemType
        State = $Item.Lifecycle
        CreationTime = $Item.CreationTime
        StorageCapacity = $Item.StorageCapacity
        StorageCapacityGiB = $sizes.SizeGiB
        StorageCapacityTiB = $sizes.SizeTiB
        StorageCapacityGB = $sizes.SizeGB
        StorageCapacityTB = $sizes.SizeTB
    }

    Add-TagProperties -Object $fsxObj -Tags $fsxTags

    return $fsxObj
}

function Process-FSXStorageVirtualMachine {
    param($Item, $Credential, $Region, $AccountInfo, $AccountAlias)
    try {
        $svmId = $null
        $svmName = $null
        $svmState = $null
        $svmUUID = $null
        $fileSystemId = $null
        $endpointIps = $null

        if ($Item -is [string]) {
            $svmName = $Item
        } else {
            $svmId   = $Item.StorageVirtualMachineId ?? $Item.SvmId
            $svmName = $Item.Name ?? $Item.StorageVirtualMachineName
            $svmState = $Item.Lifecycle ?? $Item.OperationalState
            $svmUUID = $Item.UUID ?? $Item.StorageVirtualMachineArn
            $fileSystemId = $Item.FileSystemId

            try {
                if ($Item.ManagementEndpoint -and $Item.ManagementEndpoint.IpAddresses) {
                    $endpointIps = ($Item.ManagementEndpoint.IpAddresses -join ";")
                } elseif ($Item.Endpoints -and $Item.Endpoints.Nfs) {
                    $endpointIps = ($Item.Endpoints.Nfs -join ";")
                }
            } catch {}
        }

        $svmTags = @()
        try {
            if ($Item.ResourceARN -and (Get-Command Get-FSXResourceTag -ErrorAction SilentlyContinue)) {
                $svmTags = Get-FSXResourceTag -ResourceArn $Item.ResourceARN -Credential $Credential -Region $Region -ErrorAction SilentlyContinue
            } elseif ($Item.Tags) {
                $svmTags = $Item.Tags
            }
        } catch {}

        $volumes = @()
        try {
            $allVolumes = Get-FSXVolume -Region $Region -Credential $Credential -ErrorAction Stop
            if ($allVolumes) {
                $volumes = $allVolumes | Where-Object {
                    $_.OntapConfiguration -and $_.OntapConfiguration.StorageVirtualMachineId -eq $svmId
                }
            }
            Write-ScriptOutput ("DEBUG: Found {0} volume(s) for SVM {1}" -f ($volumes.Count), $svmId) -Level Info
        } catch {
            Write-ScriptOutput ("DEBUG: Get-FSXVolume failed: {0}" -f $_.Exception.Message) -Level Warning
        }

        function Get-VolumeSizeBytes {
            param($v)
            if (-not $v) { return 0 }
            if ($v.OntapConfiguration) {
                if ($v.OntapConfiguration.SizeInBytes) { return [double]$v.OntapConfiguration.SizeInBytes }
                if ($v.OntapConfiguration.SizeInMegabytes) { return [double]$v.OntapConfiguration.SizeInMegabytes * 1024 * 1024 }
            }
            if ($v.SizeInBytes) { return [double]$v.SizeInBytes }
            if ($v.SizeInMegabytes) { return [double]$v.SizeInMegabytes * 1024 * 1024 }
            if ($v.SizeGiB) { return [double]$v.SizeGiB * 1024 * 1024 * 1024 }
            return 0
        }

        $totalBytes = ($volumes | ForEach-Object { Get-VolumeSizeBytes $_ } | Measure-Object -Sum).Sum

        $volumeCount = $volumes.Count
        $volumeDetails = $volumes | ForEach-Object {
            $bytes = Get-VolumeSizeBytes $_
            $sizeGiB = [math]::Round($bytes / 1GB, 2)
            "$($_.VolumeId):${sizeGiB}GiB"
        }

        $sizes = Convert-BytesToSizes -Bytes $totalBytes

        $obj = [PSCustomObject]@{
            AwsAccountId          = "`u{200B}$($AccountInfo.Account)"
            AwsAccountAlias       = $AccountAlias
            Region                = $Region
            FileSystemId          = $fileSystemId
            StorageVirtualMachineId = $svmId
            Name                  = $svmName
            OperationalState      = $svmState
            UUID                  = $svmUUID
            VolumeCount           = $volumeCount
            VolumesSizeGiB        = $sizes.SizeGiB
            VolumesSizeTiB        = $sizes.SizeTiB
            VolumesSizeGB         = $sizes.SizeGB
            VolumesSizeTB         = $sizes.SizeTB
            VolumeDetails         = ($volumeDetails -join ";")
        }

        Add-TagProperties -Object $obj -Tags $svmTags
        return $obj
    } catch {
        Write-ScriptOutput "Failed to process FSx ONTAP SVM item: $_" -Level Warning
        return $null
    }
}

function Ensure-Kubectl {
    if (-not (Get-Command kubectl -ErrorAction SilentlyContinue)) {
        Write-ScriptOutput "kubectl not found. Attempting to install for PVC/node enumeration." -Level Warning

        if ($env:OS -like "*Windows*") {
            try {
                $kubectlVersion = (Invoke-RestMethod https://dl.k8s.io/release/stable.txt).Trim()
                $kubectlUrl = "https://dl.k8s.io/release/$kubectlVersion/bin/windows/amd64/kubectl.exe"
                $kubectlDir = "C:\kubectl"
                if (-not (Test-Path $kubectlDir)) {
                    try {
                        New-Item -ItemType Directory -Path $kubectlDir -Force | Out-Null
                    } catch {
                        $kubectlDir = Join-Path $env:TEMP "kubectl"
                        if (-not (Test-Path $kubectlDir)) {
                            New-Item -ItemType Directory -Path $kubectlDir -Force | Out-Null
                        }
                        Write-ScriptOutput "Using fallback path $kubectlDir instead of C:\kubectl" -Level Warning
                    }
                }

                $kubectlPath = Join-Path $kubectlDir "kubectl.exe"

                Write-ScriptOutput "Downloading kubectl from $kubectlUrl to $kubectlPath..." -Level Info
                Invoke-WebRequest -Uri $kubectlUrl -OutFile $kubectlPath -UseBasicParsing

                $env:Path += ";$kubectlDir"
                $script:KubectlInstalled = $true
                $script:KubectlDir = $kubectlDir 

                if (Get-Command kubectl -ErrorAction SilentlyContinue) {
                    Write-ScriptOutput "kubectl is now available: $((kubectl version --client --short 2>$null) -join ' ')" -Level Success
                } else {
                    Write-ScriptOutput "kubectl download succeeded, but it's not in PATH. Please add $kubectlDir to your system PATH manually." -Level Warning
                }
            } catch {
                Write-ScriptOutput "Failed to auto-install kubectl: $_" -Level Error
                Write-ScriptOutput "Please install kubectl manually from https://kubernetes.io/docs/tasks/tools/ and add it to your PATH." -Level Warning
            }
        } else {
            Write-ScriptOutput "Non-Windows OS detected. Please install kubectl manually." -Level Warning
        }
    } else {
        Write-ScriptOutput "kubectl is already available." -Level Info
    }
}

function Convert-K8sSizeToBytes {
     param([string]$Size)
     if (-not $Size) { return 0 }
     $s = $Size.Trim()
     if ($s -match '^(?<num>[\d\.]+)(?<unit>Ei|Pi|Ti|Gi|Mi|Ki|E|P|T|G|M|K|i)?$') {
         $num = [double]$matches['num']
         $unit = $matches['unit']
         switch ($unit) {
             'Ei' { return $num * [math]::Pow(1024,6) }
             'Pi' { return $num * [math]::Pow(1024,5) }
             'Ti' { return $num * [math]::Pow(1024,4) }
             'Gi' { return $num * [math]::Pow(1024,3) }
             'Mi' { return $num * [math]::Pow(1024,2) }
             'Ki' { return $num * 1024 }
             'E'  { return $num * 1e18 }
             'P'  { return $num * 1e15 }
             'T'  { return $num * 1e12 }
             'G'  { return $num * 1e9 }
             'M'  { return $num * 1e6 }
             'K'  { return $num * 1e3 }
             default { return $num }
         }
     } else {
         # try to parse numeric-only
         try { return [double]$s } catch { return 0 }
     }
 }

function Get-AllEKSClusters {
    param(
        $Credential,
        [string]$Region,
        $AccountInfo,
        $AccountAlias
    )

    try {
        if (-not (Get-Command Get-EKSClusterList -ErrorAction SilentlyContinue)) {
            Write-ScriptOutput "EKS: Get-EKSClusterList command not available. Skipping EKS in region $Region." -Level Warning
            return
        }

        # Include Credential when listing clusters so cross-account/profile contexts are correct
        $clusterNames = Get-EKSClusterList -Credential $Credential -Region $Region -ErrorAction Stop

        if (-not $clusterNames -or $clusterNames.Count -eq 0) {
            Write-ScriptOutput "EKS: No clusters found in region $Region" -Level Info
            return
        }

        foreach ($clusterName in $clusterNames) {
            try {
                $result = Process-EKSCluster -Item $clusterName -Credential $Credential -Region $Region -AccountInfo $AccountInfo -AccountAlias $AccountAlias
                if ($result) { $result }
            } catch {
                Write-ScriptOutput "EKS: Failed processing cluster $clusterName in $Region : $_" -Level Warning
            }
        }
    } catch {
        Write-ScriptOutput "EKS: Error fetching clusters for $Region : $_" -Level Warning
    }
}

function Process-EKSCluster {
    param(
        $Item,
        $Credential,
        [string]$Region,
        $AccountInfo,
        [string]$AccountAlias,
        [string]$RoleArn
    )

    $clusterName = if ($Item -is [string]) { $Item }
                   elseif ($Item.Name) { $Item.Name }
                   elseif ($Item.name) { $Item.name }
                   elseif ($Item.ClusterName) { $Item.ClusterName }
                   else { $null }

    if (-not $clusterName) {
        $itemType = if ($Item) { $Item.GetType().FullName } else { 'Null' }
        Write-ScriptOutput "EKS: Cluster name not found in item (type=${itemType}), skipping" -Level Warning
        return $null
    }

    $kubernetesVersion = "Unknown"
    try {
        $clusterInfo = Get-EKSCluster -Name $clusterName -Credential $Credential -Region $Region -ErrorAction Stop
        $kubernetesVersion = $clusterInfo.Version ?? "Unknown"
    } catch {
        Write-ScriptOutput "EKS: Failed to get cluster details for ${clusterName}: $_" -Level Warning
    }

    $tmpKube = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), "kubeconfig_${clusterName}_${([guid]::NewGuid())}.yaml")

    Ensure-Kubectl

    $oldEnv = @{
        AWS_ACCESS_KEY_ID     = $env:AWS_ACCESS_KEY_ID
        AWS_SECRET_ACCESS_KEY = $env:AWS_SECRET_ACCESS_KEY
        AWS_SESSION_TOKEN     = $env:AWS_SESSION_TOKEN
    }

    try {
        $accessKey = $null; $secretKey = $null; $sessionToken = $null

        if ($Credential) {
            if ($Credential.PSObject.Methods.Name -contains 'GetCredentials') {
                try {
                    $credsObj = $Credential.GetCredentials()
                    if ($credsObj) {
                        $accessKey    = $credsObj.AccessKey   ?? $accessKey
                        $secretKey    = $credsObj.SecretKey   ?? $secretKey
                        $sessionToken = $credsObj.Token       ?? $credsObj.SessionToken ?? $sessionToken
                    }
                } catch {}
            }

            $accessKey    = $accessKey    ?? $Credential.AccessKey ?? $Credential.AccessKeyId
            $secretKey    = $secretKey    ?? $Credential.SecretKey ?? $Credential.SecretAccessKey
            $sessionToken = $sessionToken ?? $Credential.SessionToken ?? $Credential.Token
        }

        if (-not $accessKey -and $AccountInfo -and $AccountInfo.AccessKey) {
            $accessKey    = $AccountInfo.AccessKey
            $secretKey    = $AccountInfo.SecretKey
            $sessionToken = $AccountInfo.SessionToken
        }

        if ($accessKey -and $secretKey) {
            $env:AWS_ACCESS_KEY_ID     = $accessKey
            $env:AWS_SECRET_ACCESS_KEY = $secretKey
            if ($sessionToken) { $env:AWS_SESSION_TOKEN = $sessionToken } else { Remove-Item Env:AWS_SESSION_TOKEN -ErrorAction SilentlyContinue }
            Write-ScriptOutput "EKS: Exported AWS env vars for cluster '${clusterName}' (region ${Region})" -Level Info
        }

        $awsArgs = @('eks','update-kubeconfig','--name',$clusterName,'--region',$Region,'--kubeconfig',$tmpKube)
        if ($RoleArn) { $awsArgs += @('--role-arn',$RoleArn) }
        $updateResult = & aws @awsArgs 2>&1
         if ($LASTEXITCODE -ne 0 -or -not (Test-Path $tmpKube)) {
             Write-ScriptOutput "EKS: Failed to update kubeconfig for ${clusterName} in ${Region}: ${updateResult}" -Level Warning
             return [PSCustomObject]@{
                 AwsAccountId    = "`u{200B}$($AccountInfo.Account)"
                 AwsAccountAlias = $AccountInfo.AccountAlias
                 Region          = $Region
                 ClusterName     = $clusterName
                 KubernetesVersion = $kubernetesVersion
                 Count           = 1
                 PVCCount        = 0
                 SizeGiB         = 0
                 SizeTiB         = 0
                 SizeGB          = 0
                 SizeTB          = 0
                 PVCDetails      = ""
                 NodeCount       = 0
                 NodeDetails     = ""
             }
         }

        $pvcCount = 0; $sizeBytes = 0; $pvcDetails = @()

        if (Get-Command kubectl -ErrorAction SilentlyContinue) {
            $env:KUBECONFIG = $tmpKube
            $pvcRaw = & kubectl get pvc --all-namespaces -o custom-columns="NAMESPACE:.metadata.namespace,NAME:.metadata.name,VOLUME:.spec.volumeName,CAPACITY:.spec.resources.requests.storage,STATUS:.status.phase" 2>&1
            if ($LASTEXITCODE -ne 0) {
                Write-ScriptOutput "EKS: kubectl get pvc failed for ${clusterName}, skipping PVCs: ${pvcRaw}" -Level Warning
            } elseif ($pvcRaw) {
                $lines = $pvcRaw -split "`n" | Where-Object { $_ -and $_ -notmatch '^NAMESPACE' }  # Skip header
                foreach ($line in $lines) {
                    if ($line -match '^(\S+)\s+(\S+)\s+(\S+)\s+(\S+)\s+(\S+)$') {
                        $namespace = $matches[1]
                        $name = $matches[2]
                        $capacity = $matches[4]
                        $status = $matches[5]
                        if ($capacity -and $capacity -ne '<none>') {
                            $bytes = Convert-K8sSizeToBytes -Size $capacity
                            if ($bytes -gt 0) {
                                $sizeBytes += $bytes
                                $pvcCount++
                                $pvcDetails += ("${namespace}/${name}:${capacity}:${status}")
                            }
                        }
                    }
                }
            }
        } else {
            Write-ScriptOutput "kubectl not available; skipping PVC enumeration for ${clusterName}" -Level Info
        }

        $nodeCount = 0
        $nodeDetails = @()
        if (Get-Command kubectl -ErrorAction SilentlyContinue) {
            $env:KUBECONFIG = $tmpKube
            # Get node names and provider IDs
            $nodeRaw = & kubectl get nodes -o custom-columns="NAME:.metadata.name,PROVIDER_ID:.spec.providerID" 2>&1
            if ($LASTEXITCODE -eq 0 -and $nodeRaw) {
                $nodeLines = $nodeRaw | Where-Object { $_ -and $_.Trim() -ne "" -and -not $_.StartsWith("NAME") }
                $nodeCount = $nodeLines.Count
                foreach ($line in $nodeLines) {
                    $parts = $line -split '\s+', 2
                    if ($parts.Count -ge 2) {
                        $nodeName = $parts[0].Trim()
                        $providerId = $parts[1].Trim()
                        $instanceType = '<none>'
                        if ($providerId -and $providerId -match 'aws://.*/(.*)') {
                            $instanceId = $matches[1]
                            try {
                                $ec2Instance = Get-EC2Instance -InstanceId $instanceId -Credential $Credential -Region $Region -ErrorAction Stop
                                if ($ec2Instance -and $ec2Instance.Instances) {
                                    $instanceType = $ec2Instance.Instances[0].InstanceType.Value
                                }
                            } catch {
                                Write-ScriptOutput "EKS: Failed to get instance type for node ${nodeName} (ID: ${instanceId}): $_" -Level Warning
                            }
                        }
                        $nodeDetails += "${nodeName}:${instanceType}"
                    }
                }
            } else {
                Write-ScriptOutput "EKS: kubectl get nodes failed for ${clusterName}, skipping nodes: ${nodeRaw}" -Level Warning
            }
        } else {
            Write-ScriptOutput "kubectl not available; skipping node enumeration for ${clusterName}" -Level Info
        }

        $sizeGiB = [math]::Round($sizeBytes / 1GB, 4)
        $sizeTiB = [math]::Round($sizeGiB / 1024, 6)
        $sizeGB  = [math]::Round($sizeGiB * 1.073741824, 4)
        $sizeTB  = [math]::Round($sizeGB / 1024, 6)

        return [PSCustomObject]@{
            AwsAccountId    = "`u{200B}$($AccountInfo.Account)"
            AwsAccountAlias = $AccountInfo.AccountAlias
            Region          = $Region
            ClusterName     = $clusterName
            KubernetesVersion = $kubernetesVersion
            PVCCount        = $pvcCount
            NodeCount       = $nodeCount
            SizeGiB         = $sizeGiB
            SizeTiB         = $sizeTiB
            SizeGB          = $sizeGB
            SizeTB          = $sizeTB
            PVCDetails      = ($pvcDetails -join ";")
            NodeDetails     = ($nodeDetails -join ";")
        }

    } catch {
        Write-ScriptOutput "EKS: Exception processing cluster ${clusterName}: $_" -Level Warning
        return $null
    } finally {
        try { if (Test-Path $tmpKube) { Remove-Item $tmpKube -Force -ErrorAction SilentlyContinue } } catch {}
        Remove-Item Env:KUBECONFIG -ErrorAction SilentlyContinue
        foreach ($k in $oldEnv.Keys) {
            if ($null -ne $oldEnv[$k]) { ${env:$k} = $oldEnv[$k] } else { Remove-Item Env:$k -ErrorAction SilentlyContinue }
        }
    }
}

function Process-UnattachedVolume {
    param($Item, $Credential, $Region, $AccountInfo, $AccountAlias)

    $volumeTags = @()
    try {
        $tagFilter = @{
            Name = "resource-id"
            Values = @($Item.VolumeId)
        }
        $tagResponse = Get-EC2Tag -Filter $tagFilter -Credential $Credential -Region $Region -ErrorAction Stop
        $volumeTags = $tagResponse

        if ($volumeTags -and $volumeTags.Count -gt 0) {
            Write-ScriptOutput "Retrieved $($volumeTags.Count) tags for volume $($Item.VolumeId)" -Level Info
        }
    }
    catch {
        try {
            $volumeDetail = Get-EC2Volume -VolumeId $Item.VolumeId -Credential $Credential -Region $Region -ErrorAction Stop
            if ($volumeDetail -and $volumeDetail.Tags) {
                $volumeTags = $volumeDetail.Tags
                Write-ScriptOutput "Retrieved tags via volume detail for $($Item.VolumeId)" -Level Info
            }
        }
        catch {
            Write-ScriptOutput "Failed to get tags for volume $($Item.VolumeId): $($_.Exception.Message)" -Level Warning
        }
    }

    $volumeBytes = $Item.Size * 1GB
    $sizes = Convert-BytesToSizes -Bytes $volumeBytes

    $volumeObj = [PSCustomObject]@{
        AwsAccountId = "`u{200B}$($AccountInfo.Account)"
        AwsAccountAlias = $AccountAlias
        Region = $Region
        VolumeId = $Item.VolumeId
        VolumeType = $Item.VolumeType
        Size = $Item.Size
        State = $Item.State
        CreateTime = $Item.CreateTime
        SizeGiB = $sizes.SizeGiB
        SizeTiB = $sizes.SizeTiB
        SizeGB = $sizes.SizeGB
        SizeTB = $sizes.SizeTB
    }

    Add-TagProperties -Object $volumeObj -Tags $volumeTags

    return $volumeObj
}

function Process-RDSInstance {
    param($Item, $Credential, $Region, $AccountInfo, $AccountAlias)
    try {
        if ($Item.Engine -and ($Item.Engine -match '^(?i:docdb)')) {
            Write-ScriptOutput "Skipping DocumentDB resource (engine='$($Item.Engine)') $($Item.DBInstanceIdentifier)" -Level Info
            return $null
        }

        $sizeGB = 0
        if ($Item.AllocatedStorage) {
            $sizeGB = [double]$Item.AllocatedStorage
        }

        if ($Item.Engine -match '^(?i:aurora)') {
            $sizeGB = 0
        }

        $sizes = Convert-BytesToSizes -Bytes ($sizeGB * 1GB)

        try {
            $clusterId = if ($Item.DBClusterIdentifier) { $Item.DBClusterIdentifier } else { $null }
            if (-not $clusterId -and $Item.DBInstanceIdentifier) {
                $desc = Get-RDSDBInstance -DBInstanceIdentifier $Item.DBInstanceIdentifier -Credential $Credential -Region $Region -ErrorAction SilentlyContinue
                if ($desc -and $desc.DBInstance -and $desc.DBInstance.DBClusterIdentifier) { $clusterId = $desc.DBInstance.DBClusterIdentifier }
            }
            if ($clusterId) {
                $cwParams = @{
                    Namespace  = 'AWS/RDS'
                    MetricName = 'VolumeBytesUsed'
                    Dimension  = @{ Name = 'DBClusterIdentifier'; Value = $clusterId }
                    StartTime  = $utcStartTime
                    EndTime    = $utcEndTime
                    Period     = 86400
                    Statistic  = 'Maximum'
                    Credential = $Credential
                    Region     = $Region
                }
                $metric = Get-SafeCWMetricStatistic -Namespace 'AWS/RDS' -MetricName 'VolumeBytesUsed' -Dimensions @(@{ Name = 'DBClusterIdentifier'; Value = $clusterId }) -Period 86400 -Statistics 'Maximum' -StartTime $utcStartTime -EndTime $utcEndTime -Credential $Credential -Region $Region

                if ($metric -and $metric.Datapoints) {
                    $mv = ($metric.Datapoints | Sort-Object Timestamp -Descending | Select-Object -First 1).Maximum
                    if ($mv -and $mv -gt 0) {
                        $sizes = Convert-BytesToSizes -Bytes ([double]$mv)
                        Write-ScriptOutput "Aurora cluster storage-used via CloudWatch for ${clusterId}: $([math]::Round($mv/1GB,2)) GB" -Level Info
                    }
                }
                if ($sizes.SizeGiB -eq 0) {
                    $cinfo = Get-RDSDBCluster -DBClusterIdentifier $clusterId -Credential $Credential -Region $Region -ErrorAction SilentlyContinue
                    if ($cinfo) {
                        if ($cinfo.StorageUsage) {
                            $sizes = Convert-BytesToSizes -Bytes ($cinfo.StorageUsage * 1GB)
                            Write-ScriptOutput "Aurora cluster used storage via DescribeDBCluster for ${clusterId}: $($cinfo.StorageUsage) GB" -Level Info
                        } elseif ($cinfo.AllocatedStorage) {
                            $sizes = Convert-BytesToSizes -Bytes ($cinfo.AllocatedStorage * 1GB)
                            Write-ScriptOutput "Aurora cluster allocated storage via DescribeDBCluster for ${clusterId}: $($cinfo.AllocatedStorage) GB" -Level Info
                        }
                    }
                }
            }
        } catch {}

        $rdsObj = [PSCustomObject]@{
            AwsAccountId          = "`u{200B}$($AccountInfo.Account)"
            AwsAccountAlias       = $AccountAlias
            Region                = $Region
            DBInstanceIdentifier  = $Item.DBInstanceIdentifier
            Engine                = $Item.Engine
            DBInstanceClass       = $Item.DBInstanceClass
            SizeGiB               = $sizes.SizeGiB
            SizeTiB               = $sizes.SizeTiB
            SizeGB                = $sizes.SizeGB
            SizeTB                = $sizes.SizeTB
        }

        try {
            if (Get-Command Get-RDSTag -ErrorAction SilentlyContinue) {
                $tags = Get-RDSTag -ResourceName $Item.DBInstanceArn -Credential $Credential -Region $Region -ErrorAction Stop
                Add-TagProperties -Object $rdsObj -Tags $tags
            }
        } catch { }

        return $rdsObj
    } catch {
        Write-ScriptOutput "Failed to process RDS instance $($Item.DBInstanceIdentifier): $_" -Level Warning
        return $null
    }
}

function Process-DynamoDBTable {
    param($Item, $Credential, $Region, $AccountInfo, $AccountAlias)
    try {
        $tableName = if ($Item -is [string]) { $Item } elseif ($Item.TableName) { $Item.TableName } else { $Item }

        $tbl = $null
        try { $tbl = Get-DDBTable -TableName $tableName -Credential $Credential -Region $Region -ErrorAction Stop } catch { $tbl = $null }

        $tableArn = $null
        $tableId = $null
        $tableStatus = $null
        $tableSizeBytes = 0
        $itemCount = $null

        if ($tbl) {
            if ($tbl -and $tbl.Table) {
                $src = $tbl.Table
            } else {
                $src = $tbl
            }

            if ($src.TableArn)   { $tableArn = $src.TableArn }
            if ($src.TableId)    { $tableId = $src.TableId }
            if ($src.TableStatus) { $tableStatus = $src.TableStatus }
            if ($src.TableSizeBytes -ne $null) { $tableSizeBytes = [double]$src.TableSizeBytes }
            if ($src.ItemCount -ne $null) { $itemCount = [long]$src.ItemCount }

            if ($tableStatus -and $tableStatus.PSObject.Properties.Name -contains 'Value') {
                $tableStatus = $tableStatus.Value
            }
        }

        if (-not $tableArn -and $Item.TableArn) { $tableArn = $Item.TableArn }
        if (-not $tableId  -and $Item.TableId)  { $tableId = $Item.TableId }
        if (-not $tableStatus -and $Item.TableStatus) { $tableStatus = $Item.TableStatus }
        if (($tableSizeBytes -eq 0) -and $Item.TableSizeBytes -ne $null) {
            $tableSizeBytes = [double]$Item.TableSizeBytes
        }
        if (-not $itemCount -and $Item.ItemCount -ne $null) { $itemCount = [long]$Item.ItemCount }

        $sizes = Convert-BytesToSizes -Bytes $tableSizeBytes

        $ddbObj = [PSCustomObject]@{
            AwsAccountId     = "`u{200B}$($AccountInfo.Account)"
            AwsAccountAlias  = $AccountAlias
            Region           = $Region
            TableName        = $tableName
            TableId          = $tableId
            TableArn         = $tableArn
            TableSizeBytes   = $tableSizeBytes
            TableStatus      = $tableStatus
            ItemCount        = $itemCount
            TableSizeGiB     = $sizes.SizeGiB
            TableSizeTiB     = $sizes.SizeTiB
            TableSizeGB      = $sizes.SizeGB
            TableSizeTB      = $sizes.SizeTB
        }

        try {
            if (Get-Command Get-DDBTableTag -ErrorAction SilentlyContinue) {
                $tags = Get-DDBTableTag -TableName $tableName -Credential $Credential -Region $Region -ErrorAction Stop
                Add-TagProperties -Object $ddbObj -Tags $tags
            }
        } catch { }

        return $ddbObj
    } catch {
        Write-ScriptOutput "Failed to process DynamoDB table $($Item): $_" -Level Warning
        return $null
    }
}

function Process-RedshiftCluster {
    param($Item, $Credential, $Region, $AccountInfo, $AccountAlias)
    try {
        $clusterId = $Item.ClusterIdentifier
        $nodeCount = if ($Item.NumberOfNodes) { [int]$Item.NumberOfNodes } else { 0 }
        $nodeType = if ($Item.NodeType) { $Item.NodeType } else { $null }

        $totalBytes = 0
        if ($Item -and $Item.PSObject -and $Item.PSObject.Properties.Name) {
            if ($Item.PSObject.Properties.Name -contains 'TotalStorageInBytes') { $totalBytes = [double]$Item.TotalStorageInBytes }
            elseif ($Item.PSObject.Properties.Name -contains 'TotalStorageInMegaBytes') { $totalBytes = [double]$Item.TotalStorageInMegaBytes * 1MB }
            elseif ($Item.PSObject.Properties.Name -contains 'TotalStorageInMegabytes') { $totalBytes = [double]$Item.TotalStorageInMegabytes * 1MB }
            elseif ($Item.PSObject.Properties.Name -contains 'TotalStorageCapacityInMegaBytes') { $totalBytes = [double]$Item.TotalStorageCapacityInMegaBytes * 1MB }
            elseif ($Item.PSObject.Properties.Name -contains 'TotalStorageCapacityInMegabytes') { $totalBytes = [double]$Item.TotalStorageCapacityInMegabytes * 1MB }
            elseif ($Item.PSObject.Properties.Name -contains 'TotalStorageCapacityInBytes') { $totalBytes = [double]$Item.TotalStorageCapacityInBytes }
            elseif ($Item.PSObject.Properties.Name -contains 'TotalStorage') {
                $v = $Item.TotalStorage
                if ($v -ne $null) { $totalBytes = [double]$v }
            }
            elseif ($Item.PSObject.Properties.Name -contains 'Storage') { $totalBytes = [double]$Item.Storage }
        }

        if ($totalBytes -eq 0 -and $Item.Cluster -and $Item.Cluster.PSObject.Properties.Name) {
            if ($Item.Cluster.PSObject.Properties.Name -contains 'TotalStorageInBytes') { $totalBytes = [double]$Item.Cluster.TotalStorageInBytes }
            elseif ($Item.Cluster.PSObject.Properties.Name -contains 'TotalStorageInMegaBytes') { $totalBytes = [double]$Item.Cluster.TotalStorageInMegaBytes * 1MB }
            elseif ($Item.Cluster.PSObject.Properties.Name -contains 'TotalStorageInMegabytes') { $totalBytes = [double]$Item.Cluster.TotalStorageInMegabytes * 1MB }
            elseif ($Item.Cluster.PSObject.Properties.Name -contains 'TotalStorageCapacityInMegaBytes') { $totalBytes = [double]$Item.Cluster.TotalStorageCapacityInMegaBytes * 1MB }
            elseif ($Item.Cluster.PSObject.Properties.Name -contains 'TotalStorageCapacityInMegabytes') { $totalBytes = [double]$Item.Cluster.TotalStorageCapacityInMegabytes * 1MB }
            elseif ($Item.Cluster.PSObject.Properties.Name -contains 'TotalStorageCapacityInBytes') { $totalBytes = [double]$Item.Cluster.TotalStorageCapacityInBytes }
            elseif ($Item.Cluster.PSObject.Properties.Name -contains 'TotalStorage') { $totalBytes = [double]$Item.Cluster.TotalStorage }
        }

        if ($totalBytes -eq 0 -and $clusterId) {
            try {
                $desc = Get-RSCluster -ClusterIdentifier $clusterId -Credential $Credential -Region $Region -ErrorAction Stop
                Write-ScriptOutput "DEBUG: DescribeClusters returned properties: $(($desc.PSObject.Properties.Name) -join ', ')" -Level Info
                if ($desc -and $desc.PSObject.Properties.Name -contains 'Cluster') { $src = $desc.Cluster }
                elseif ($desc -and $desc.PSObject.Properties.Name -contains 'Clusters') { $src = if ($desc.Clusters.Count -gt 0) { $desc.Clusters[0] } else { $null } }
                else { $src = $desc }

                if ($src) {
                    if ($src.PSObject.Properties.Name -contains 'TotalStorageInBytes') { $totalBytes = [double]$src.TotalStorageInBytes }
                    elseif ($src.PSObject.Properties.Name -contains 'TotalStorageInMegaBytes') { $totalBytes = [double]$src.TotalStorageInMegaBytes * 1MB }
                    elseif ($src.PSObject.Properties.Name -contains 'TotalStorageInMegabytes') { $totalBytes = [double]$src.TotalStorageInMegabytes * 1MB }
                    elseif ($src.PSObject.Properties.Name -contains 'TotalStorageCapacityInMegaBytes') { $totalBytes = [double]$src.TotalStorageCapacityInMegaBytes * 1MB }
                    elseif ($src.PSObject.Properties.Name -contains 'TotalStorageCapacityInMegabytes') { $totalBytes = [double]$src.TotalStorageCapacityInMegabytes * 1MB }
                    elseif ($src.PSObject.Properties.Name -contains 'TotalStorageCapacityInBytes') { $totalBytes = [double]$src.TotalStorageCapacityInBytes }
                    elseif ($src.PSObject.Properties.Name -contains 'TotalStorage') {
                        $v = $src.TotalStorage
                        if ($v -ne $null) { $totalBytes = [double]$v }
                    }
                    elseif ($src.PSObject.Properties.Name -contains 'Storage') { $totalBytes = [double]$src.Storage }
                }
            } catch {
                Write-ScriptOutput "DescribeClusters fallback for $clusterId did not return a storage field" -Level Info
            }
        }

        $totalSizes = Convert-BytesToSizes -Bytes $totalBytes

        $rsObj = [PSCustomObject]@{
            AwsAccountId      = "`u{200B}$($AccountInfo.Account)"
            AwsAccountAlias   = $AccountAlias
            Region            = $Region
            ClusterIdentifier = $clusterId
            NodeType          = $nodeType
            NodeCount         = $nodeCount

            TotalSizeGiB      = $totalSizes.SizeGiB
            TotalSizeTiB      = $totalSizes.SizeTiB
            TotalSizeGB       = $totalSizes.SizeGB
            TotalSizeTB       = $totalSizes.SizeTB
        }

        try {
            if (Get-Command Get-RSClusterTag -ErrorAction SilentlyContinue) {
                $tags = Get-RSClusterTag -ClusterIdentifier $clusterId -Credential $Credential -Region $Region -ErrorAction Stop
                Add-TagProperties -Object $rsObj -Tags $tags
            }
        } catch { }

        return $rsObj
    } catch {
        Write-ScriptOutput "Failed to process Redshift cluster $($Item.ClusterIdentifier): $_" -Level Warning
        return $null
    }
}


function New-ServiceSummaryData {
    param(
        [string]$ServiceName,
        [array]$ServiceList
    )

    $serviceConfig = $script:ServiceRegistry[$ServiceName]
    $summaryData = @()
    $totalCount = if ($ServiceList) { $ServiceList.Count } else { 0 }

    $totals = if ($ServiceName -eq "S3") {
        Get-S3ServiceTotals -ServiceList $ServiceList
    } elseif ($ServiceName -eq "FSX") {
        Get-FSxServiceTotals -ServiceList $ServiceList
    } else {
        Get-StandardServiceTotals -ServiceList $ServiceList -SizeProperty $serviceConfig.SizeProperty
    }

    if ($ServiceName -eq "DynamoDB") {
        $totalTableBytes = 0
        $totalItemCount = 0
        if ($ServiceList -and $ServiceList.Count -gt 0) {
            $totalTableBytes = ($ServiceList | ForEach-Object { if ($_.TableSizeBytes -ne $null -and $_.TableSizeBytes -ne '') { [double]$_.TableSizeBytes } else { 0 } } | Measure-Object -Sum).Sum
            $totalItemCount  = ($ServiceList | ForEach-Object { if ($_.ItemCount -ne $null -and $_.ItemCount -ne '') { [double]$_.ItemCount } else { 0 } } | Measure-Object -Sum).Sum
        }
    }

    if ($ServiceName -eq "FSX" -and $ServiceList -and $ServiceList.Count -gt 0) {
        $fsxTypes = $ServiceList | Select-Object -ExpandProperty Type -Unique | Sort-Object

        foreach ($fsxType in $fsxTypes) {
            $typeItems = $ServiceList | Where-Object { $_.Type -eq $fsxType }
            $typeTotals = Get-FSxServiceTotals -ServiceList $typeItems

            $summaryData += [PSCustomObject]@{
                "ResourceType" = $fsxType
                "Region" = "All"
                "Count" = $typeItems.Count
                "Total Size (GiB)" = [math]::Round($typeTotals.GiB, 3)
                "Total Size (GB)" = [math]::Round($typeTotals.GB, 3)
                "Total Size (TiB)" = [math]::Round($typeTotals.TiB, 4)
                "Total Size (TB)" = [math]::Round($typeTotals.TB, 4)
            }
        }

        $summaryData += [PSCustomObject]@{
            "ResourceType" = "Total"
            "Region" = "All"
            "Count" = $totalCount
            "Total Size (GiB)" = [math]::Round($totals.GiB, 3)
            "Total Size (GB)" = [math]::Round($totals.GB, 3)
            "Total Size (TiB)" = [math]::Round($totals.TiB, 4)
            "Total Size (TB)" = [math]::Round($totals.TB, 4)
        }
    } else {
        $displayName = "$ServiceName $($serviceConfig.DisplayName)"
        $row = [ordered]@{
            "ResourceType" = $displayName
            "Region" = "All"
            "Count" = $totalCount
            "Total Size (GiB)" = [math]::Round($totals.GiB, 3)
            "Total Size (GB)" = [math]::Round($totals.GB, 3)
            "Total Size (TiB)" = [math]::Round($totals.TiB, 4)
            "Total Size (TB)" = [math]::Round($totals.TB, 4)
        }
        if ($ServiceName -eq "DynamoDB") {
            $row["Total Table Size (Bytes)"] = $totalTableBytes
            $row["Total Item Count"] = $totalItemCount
        }
        if ($ServiceName -eq "EKS") {
            $totalNodeCount = ($ServiceList | ForEach-Object { if ($_.NodeCount -ne $null) { [int]$_.NodeCount } else { 0 } } | Measure-Object -Sum).Sum
            $row["Total Node Count"] = $totalNodeCount
        }
        $summaryData += [PSCustomObject]$row
    }

    $summaryData += Get-EmptyRow
    $summaryData += Get-EmptyRow

    $summaryData += Get-RegionalBreakdown -ServiceName $ServiceName -ServiceList $ServiceList

    return $summaryData
}

function Get-RegionalBreakdown {
    param([string]$ServiceName, [array]$ServiceList)

    $regions = if ($ServiceList) {
        $ServiceList | Select-Object -ExpandProperty Region -Unique | Sort-Object
    } else { @() }

    if ($regions.Count -eq 0) { return @() }

    $breakdownData = @()

    $header = [ordered]@{
        "ResourceType" = "--- $($ServiceName.ToUpper()) REGIONAL BREAKDOWN ---"
        "Region" = ""
        "Count" = ""
        "Total Size (GiB)" = ""
        "Total Size (GB)" = ""
        "Total Size (TiB)" = ""
        "Total Size (TB)" = ""
    }
    if ($ServiceName -eq "DynamoDB") {
        $header["Total Table Size (Bytes)"] = ""
        $header["Total Item Count"] = ""
    }
    if ($ServiceName -eq "EKS") {
        $header["Total Node Count"] = ""
    }
    $breakdownData += [PSCustomObject]$header

    if ($ServiceName -eq "FSX") {
        $breakdownData += [PSCustomObject]@{
            "ResourceType" = "Region"
            "Region" = "Type"
            "Count" = "Count"
            "Total Size (GiB)" = "Total Size (GiB)"
            "Total Size (GB)" = "Total Size (GB)"
            "Total Size (TiB)" = "Total Size (TiB)"
            "Total Size (TB)" = "Total Size (TB)"
        }
    } else {
        $colHeader = [ordered]@{
            "ResourceType" = "Region"
            "Region" = ""
            "Count" = "Count"
            "Total Size (GiB)" = "Total Size (GiB)"
            "Total Size (GB)" = "Total Size (GB)"
            "Total Size (TiB)" = "Total Size (TiB)"
            "Total Size (TB)" = "Total Size (TB)"
        }
        if ($ServiceName -eq "DynamoDB") {
            $colHeader["Total Table Size (Bytes)"] = "Total Table Size (Bytes)"
            $colHeader["Total Item Count"] = "Total Item Count"
        }
        $breakdownData += [PSCustomObject]$colHeader
    }

    if ($ServiceName -eq "FSX") {
        $breakdownData += Get-FSxRegionalBreakdown -ServiceList $ServiceList -Regions $regions
    } else {
        foreach ($region in $regions) {
            $regionItems = $ServiceList | Where-Object Region -eq $region
            if ($regionItems.Count -gt 0) {
                $regionTotals = if ($ServiceName -eq "S3") {
                    Get-S3ServiceTotals -ServiceList $regionItems
                } else {
                    Get-StandardServiceTotals -ServiceList $regionItems -SizeProperty $script:ServiceRegistry[$ServiceName].SizeProperty
                }

                $row = [ordered]@{
                    "ResourceType" = $region
                    "Region" = ""
                    "Count" = $regionItems.Count
                    "Total Size (GiB)" = [math]::Round($regionTotals.GiB, 3)
                    "Total Size (GB)" = [math]::Round($regionTotals.GB, 3)
                    "Total Size (TiB)" = [math]::Round($regionTotals.TiB, 4)
                    "Total Size (TB)" = [math]::Round($regionTotals.TB, 4)
                }

                if ($ServiceName -eq "DynamoDB") {
                    $regionBytes = ($regionItems | ForEach-Object { if ($_.TableSizeBytes -ne $null -and $_.TableSizeBytes -ne '') { [double]$_.TableSizeBytes } else { 0 } } | Measure-Object -Sum).Sum
                    $regionItemsCount = ($regionItems | ForEach-Object { if ($_.ItemCount -ne $null -and $_.ItemCount -ne '') { [double]$_.ItemCount } else { 0 } } | Measure-Object -Sum).Sum
                    $row["Total Table Size (Bytes)"] = $regionBytes
                    $row["Total Item Count"] = $regionItemsCount
                }
                if ($ServiceName -eq "EKS") {
                    $regionNodeCount = ($regionItems | ForEach-Object { if ($_.NodeCount -ne $null) { [int]$_.NodeCount } else { 0 } } | Measure-Object -Sum).Sum
                    $row["Total Node Count"] = $regionNodeCount
                }

                $breakdownData += [PSCustomObject]$row
            }
        }
    }

    return $breakdownData
}

function Get-StandardServiceTotals {
    param([array]$ServiceList, [string]$SizeProperty)

    if (-not $ServiceList -or $ServiceList.Count -eq 0) {
        return @{ GiB = 0; TiB = 0; GB = 0; TB = 0 }
    }

    if ($SizeProperty -and $SizeProperty -eq 'Custom') { $SizeProperty = $null }

    if ($SizeProperty -and ($SizeProperty.ToLower().Contains('byte'))) {
        $rawSum = ($ServiceList | ForEach-Object { if ($_.$SizeProperty -ne $null -and $_.$SizeProperty -ne '') { [double]($_.$SizeProperty) } else { 0 } } | Measure-Object -Sum).Sum
        return @{
            GiB = $rawSum / 1GB
            TiB = $rawSum / 1TB
            GB  = $rawSum / 1e9
            TB  = $rawSum / 1e12
        }
    }

    try {
        if ($SizeProperty -and (($ServiceList | Select-Object -First 1).PSObject.Properties.Name -contains $SizeProperty)) {
            $giB = ($ServiceList | ForEach-Object { if ($_.$SizeProperty -ne $null -and $_.$SizeProperty -ne '') { [double]($_.$SizeProperty) } else { 0 } } | Measure-Object -Sum).Sum
            $tib = ($ServiceList | ForEach-Object { if ($_.SizeTiB -ne $null -and $_.SizeTiB -ne '') { [double]$_.SizeTiB } else { 0 } } | Measure-Object -Sum).Sum
            $gb  = ($ServiceList | ForEach-Object { if ($_.SizeGB -ne $null -and $_.SizeGB -ne '') { [double]$_.SizeGB } else { 0 } } | Measure-Object -Sum).Sum
            $tb  = ($ServiceList | ForEach-Object { if ($_.SizeTB -ne $null -and $_.SizeTB -ne '') { [double]$_.SizeTB } else { 0 } } | Measure-Object -Sum).Sum

            return @{ GiB = $giB; TiB = $tib; GB = $gb; TB = $tb }
        }

        $sumProp = {
            param($names)
            ($ServiceList | ForEach-Object {
                foreach ($n in $names) { if ($_.PSObject.Properties.Name -contains $n -and $_.$n -ne $null -and $_.$n -ne '') { return [double]($_.$n) } }
                return 0
            } | Measure-Object -Sum).Sum
        }

        $giB = & $sumProp @('TotalSizeGiB','SizeGiB','StorageCapacityGiB','TableSizeGiB','VolumesSizeGiB')
        $tib = & $sumProp @('TotalSizeTiB','SizeTiB','StorageCapacityTiB','TableSizeTiB','VolumesSizeTiB')
        $gb  = & $sumProp @('TotalSizeGB','SizeGB','StorageCapacityGB','TableSizeGB','VolumesSizeGB')
        $tb  = & $sumProp @('TotalSizeTB','SizeTB','StorageCapacityTB','TableSizeTB','VolumesSizeTB')

        if ($giB -eq 0 -and $gb -gt 0) {
            $giB = $gb / (1024/1000)
        }
        if ($tib -eq 0 -and $giB -gt 0) { $tib = $giB / 1024 }
        if ($tb -eq 0 -and $gb -gt 0) { $tb = $gb / 1000 }

        return @{ GiB = $giB; TiB = $tib; GB = $gb; TB = $tb }
    } catch {
        return @{ GiB = 0; TiB = 0; GB = 0; TB = 0 }
    }
}

function Get-S3ServiceTotals {
    param([array]$ServiceList)

    if (-not $ServiceList -or $ServiceList.Count -eq 0) {
        return @{ GiB = 0; TiB = 0; GB = 0; TB = 0 }
    }

    $totalGiB = 0.0
    $totalGB  = 0.0
    $totalTiB = 0.0
    $totalTB  = 0.0

    foreach ($item in $ServiceList) {
        if (-not $item -or -not $item.PSObject) { continue }
        $props = $item.PSObject.Properties.Name

        foreach ($p in $props) {
            $val = $null
            try { $val = $item.$p } catch { continue }
            if ($null -eq $val -or $val -eq '') { continue }

            try { $d = [double]$val } catch { continue }

            if ($p -match '(_SizeGiB$|^SizeGiB$)')       { $totalGiB += $d; continue }
            if ($p -match '(_SizeGB$|^SizeGB$)')         { $totalGB  += $d; continue }
            if ($p -match '(_SizeTiB$|^SizeTiB$)')       { $totalTiB += $d; continue }
            if ($p -match '(_SizeTB$|^SizeTB$)')         { $totalTB  += $d; continue }
        }
    }

    if ($totalGiB -eq 0 -and $totalGB -gt 0) { $totalGiB = $totalGB / (1024/1000) }   
    if ($totalTiB -eq 0 -and $totalGiB -gt 0) { $totalTiB = $totalGiB / 1024 }
    if ($totalTB -eq 0 -and $totalGB -gt 0) { $totalTB = $totalGB / 1000 }

    return @{ GiB = $totalGiB; TiB = $totalTiB; GB = $totalGB; TB = $totalTB }
}
function Get-FSxServiceTotals {
    param([array]$ServiceList)

    if (-not $ServiceList -or $ServiceList.Count -eq 0) {
        return @{ GiB = 0; TiB = 0; GB = 0; TB = 0 }
    }

    return @{
        GiB = ($ServiceList.StorageCapacityGiB | Measure-Object -Sum).Sum
        TiB = ($ServiceList.StorageCapacityTiB | Measure-Object -Sum).Sum
        GB = ($ServiceList.StorageCapacityGB | Measure-Object -Sum).Sum
        TB = ($ServiceList.StorageCapacityTB | Measure-Object -Sum).Sum
    }
}

function Get-FSxRegionalBreakdown {
    param([array]$ServiceList, [array]$Regions)

    $breakdownData = @()
    $fsxTypes = $ServiceList | Select-Object -ExpandProperty Type -Unique | Sort-Object

    foreach ($region in $Regions) {
        $breakdownData += [PSCustomObject]@{
            "ResourceType" = $region
            "Region" = ""
            "Count" = ""
            "Total Size (GiB)" = ""
            "Total Size (GB)" = ""
            "Total Size (TiB)" = ""
            "Total Size (TB)" = ""
        }

        foreach ($fsxType in $fsxTypes) {
            $regionTypeFsx = $ServiceList | Where-Object { $_.Region -eq $region -and $_.Type -eq $fsxType }
            if ($regionTypeFsx.Count -gt 0) {
                $typeTotals = Get-FSxServiceTotals -ServiceList $regionTypeFsx

                $breakdownData += [PSCustomObject]@{
                    "ResourceType" = ""
                    "Region" = $fsxType
                    "Count" = $regionTypeFsx.Count
                    "Total Size (GiB)" = [math]::Round($typeTotals.GiB, 3)
                    "Total Size (GB)" = [math]::Round($typeTotals.GB, 3)
                    "Total Size (TiB)" = [math]::Round($typeTotals.TiB, 4)
                    "Total Size (TB)" = [math]::Round($typeTotals.TB, 4)
                }
            }
        }
    }

    return $breakdownData
}



function New-AccountLevelSummary {
    param([string]$AccountId, [string]$AccountAlias)

    try {
        $accountSummaryFile = "${AccountId}_summary_$date_string.xlsx"
        $dataSheets = [ordered]@{}

        foreach ($serviceName in $script:ServiceRegistry.Keys) {
            $serviceList = $script:ServiceDataByAccount[$AccountId][$serviceName]
            $summarySheetName = if ($serviceName -eq "UnattachedVolumes") {
                "Unattached Volume Summary"
            } elseif ($serviceName -eq "FSX_SVM") {
                "FSx SVM Summary"
            } else {
                "$serviceName Summary"
            }
            if ($serviceList -and $serviceList.Count -gt 0) {
                $dataSheets[$summarySheetName] = New-ServiceSummaryData -ServiceName $serviceName -ServiceList $serviceList
            }        }

        $detailSheets = @{
            "EC2 Details" = $script:ServiceDataByAccount[$AccountId]["EC2"]
            "S3 Details" = $script:ServiceDataByAccount[$AccountId]["S3"]
            "Unattached Volumes" = $script:ServiceDataByAccount[$AccountId]["UnattachedVolumes"]
            "EFS Details" = $script:ServiceDataByAccount[$AccountId]["EFS"]
            "FSx Details" = $script:ServiceDataByAccount[$AccountId]["FSX"]
            "FSx SVM Details" = $script:ServiceDataByAccount[$AccountId]["FSX_SVM"]
            "RDS Details" = $script:ServiceDataByAccount[$AccountId]["RDS"]
            "DynamoDB Details" = $script:ServiceDataByAccount[$AccountId]["DynamoDB"]
            "Redshift Details" = $script:ServiceDataByAccount[$AccountId]["Redshift"]
            "EKS Details" = $script:ServiceDataByAccount[$AccountId]["EKS"]
        }

        foreach ($sheetName in $detailSheets.Keys) {
            if ($detailSheets[$sheetName] -and $detailSheets[$sheetName].Count -gt 0) {
                if ($sheetName -eq "S3 Details") {
                    $dataSheets[$sheetName] = $detailSheets[$sheetName] |
                        Select-Object AwsAccountId, AwsAccountAlias, Region, BucketName, CreationDate, SizeGiB, SizeTiB, SizeGB, SizeTB
                }
                elseif ($sheetName -eq "FSx SVM Details") {
                    $svmItems = $detailSheets[$sheetName]
                    $tagProps = ($svmItems | ForEach-Object { $_.PSObject.Properties.Name } | Where-Object { $_ -like 'Tag_*' } | Sort-Object -Unique)
                    $cols = @(
                        'AwsAccountId','AwsAccountAlias','Region','FileSystemId','StorageVirtualMachineId','Name',
                        'OperationalState','UUID','VolumeCount',
                        @{Name='SizeGiB';Expression={ $_.VolumesSizeGiB }},
                        @{Name='SizeTiB';Expression={ $_.VolumesSizeTiB }},
                        @{Name='SizeGB';Expression={ $_.VolumesSizeGB }},
                        @{Name='SizeTB';Expression={ $_.VolumesSizeTB }},
                        'VolumeDetails'
                    ) + $tagProps
                    $dataSheets[$sheetName] = $svmItems | Select-Object $cols
                }
                else {
                    $dataSheets[$sheetName] = $detailSheets[$sheetName]
                }
            }
        }

        $result = Export-DataToExcel -FilePath $accountSummaryFile -DataSheets $dataSheets

        if ($result) {
            Write-ScriptOutput "Account-level summary created: $accountSummaryFile" -Level Success
            $script:AllOutputFiles.Add($accountSummaryFile) | Out-Null
        }

        return $result
    }
    catch {
        Write-ScriptOutput "Error creating account-level summary for account ${AccountId}: $_" -Level Error
        return $false
    }
}

function Export-ServiceCSVFiles {
    param([string]$AccountId, [string]$AccountAlias)

    try {
        $accountSuffix = if ($AccountAlias -and $AccountAlias -ne "Unknown") { $AccountAlias } else { $AccountId }

        $outputFileEc2Instance = "${baseOutputEc2Instance}_${accountSuffix}_$date_string.csv"
        $outputFileEc2UnattachedVolume = "${baseOutputEc2UnattachedVolume}_${accountSuffix}_$date_string.csv"
        $outputFileS3 = "${baseOutputS3}_${accountSuffix}_$date_string.csv"
        $outputFileEFS = "${baseOutputEFS}_${accountSuffix}_$date_string.csv"
        $outputFileFSx = "${baseOutputFSx}_${accountSuffix}_$date_string.csv"

        $outputFileFSxSVM = "${baseOutputFSxSVM}_${accountSuffix}_$date_string.csv"
        $outputFileRDS = "${baseOutputRDS}_${accountSuffix}_$date_string.csv"
        $outputFileDynamoDB = "${baseOutputDynamoDB}_${accountSuffix}_$date_string.csv"
        $outputFileRedshift = "${baseOutputRedshift}_${accountSuffix}_$date_string.csv"

        $serviceExports = @{
            EC2 = @{ File = $outputFileEc2Instance; Data = $script:ServiceDataByAccount[$AccountId]["EC2"] }
            UnattachedVolumes = @{ File = $outputFileEc2UnattachedVolume; Data = $script:ServiceDataByAccount[$AccountId]["UnattachedVolumes"] }
            S3 = @{ File = $outputFileS3; Data = $script:ServiceDataByAccount[$AccountId]["S3"] }
            EFS = @{ File = $outputFileEFS; Data = $script:ServiceDataByAccount[$AccountId]["EFS"] }
            FSX = @{ File = $outputFileFSx; Data = $script:ServiceDataByAccount[$AccountId]["FSX"] }
            FSX_SVM = @{ File = $outputFileFSxSVM; Data = $script:ServiceDataByAccount[$AccountId]["FSX_SVM"] }
            RDS = @{ File = $outputFileRDS; Data = $script:ServiceDataByAccount[$AccountId]["RDS"] }
            DynamoDB = @{ File = $outputFileDynamoDB; Data = $script:ServiceDataByAccount[$AccountId]["DynamoDB"] }
            Redshift = @{ File = $outputFileRedshift; Data = $script:ServiceDataByAccount[$AccountId]["Redshift"] }
            EKS = @{ File = $outputFileEKS; Data = $script:ServiceDataByAccount[$AccountId]["EKS"] }
        }

        foreach ($serviceName in $serviceExports.Keys) {
            $export = $serviceExports[$serviceName]
            if ($export.Data -and $export.Data.Count -gt 0) {
                if (-not $export.File -or [string]::IsNullOrWhiteSpace($export.File)) {
                    Write-ScriptOutput "Skipping export for $serviceName because target file path is empty" -Level Warning
                    continue
                }

                if ($serviceName -eq 'S3') {
                    $export.Data |
                        Select-Object AwsAccountId, AwsAccountAlias, Region, BucketName, CreationDate, SizeGiB, SizeTiB, SizeGB, SizeTB |
                        Export-Csv -Path $export.File -NoTypeInformation
                }
                else {
                    $export.Data | Export-Csv -Path $export.File -NoTypeInformation
                }

                Write-ScriptOutput "Exported $serviceName data to $($export.File)" -Level Info
                $script:AllOutputFiles.Add($export.File) | Out-Null
            }
        }
    }
    catch {
        Write-ScriptOutput "Error exporting CSV files for account ${AccountId}: $_" -Level Error
    }
}

function Export-DataToExcel {
    param(
        [string]$FilePath,
        [System.Collections.IDictionary]$DataSheets
    )

    try {
        if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
            Write-ScriptOutput "ImportExcel module not found. Install it with: Install-Module ImportExcel -Scope CurrentUser" -Level Warning
            return $false
        }

        Import-Module ImportExcel -Force -ErrorAction Stop
        Write-ScriptOutput "Using ImportExcel module for Excel generation..." -Level Success

        if (Test-Path $FilePath) {
            Remove-Item $FilePath -Force
        }
        $worksheetOrder = @(
            "EC2 Summary", "S3 Summary","EFS Summary", "FSX Summary","FSx SVM Summary",
            "RDS Summary", "DynamoDB Summary", "Redshift Summary","EKS Summary", "Unattached Volume Summary",
            "EC2 Details", "S3 Details", "EFS Details", "FSx Details", "FSx SVM Details",
            "RDS Details", "DynamoDB Details", "Redshift Details", "EKS Details","Unattached Volumes Details"
        )
        foreach ($sheetName in $worksheetOrder) {
            if ($DataSheets.Keys -contains $sheetName) {
                $data = $DataSheets[$sheetName]
                $data | Export-Excel -Path $FilePath -WorksheetName $sheetName -AutoSize -FreezeTopRow -BoldTopRow

                if ($sheetName -like "*Summary*") {
                    try {
                        $excel = Open-ExcelPackage -Path $FilePath
                        $worksheet = $excel.Workbook.Worksheets[$sheetName]

                        for ($row = 1; $row -le $worksheet.Dimension.End.Row; $row++) {
                            $cellValue = $worksheet.Cells[$row, 1].Value
                            if ($cellValue -and ($cellValue.ToString() -eq "Region" -or
                                ($cellValue.ToString() -eq "Region" -and $worksheet.Cells[$row, 2].Value -eq "Type"))) {

                                $prevRowValue = if ($row -gt 1) { $worksheet.Cells[$row-1, 1].Value } else { $null }
                                if ($prevRowValue -and $prevRowValue.ToString() -like "*BREAKDOWN*") {
                                    for ($col = 1; $col -le $worksheet.Dimension.End.Column; $col++) {
                                        $worksheet.Cells[$row, $col].Style.Font.Bold = $true
                                    }
                                }
                            }
                        }
                        Close-ExcelPackage $excel -Save
                    }
                    catch {
                        Write-ScriptOutput "Warning: Could not apply additional formatting to $sheetName : $_" -Level Warning
                    }
                }

            }
        }
        Write-ScriptOutput "Excel summary file created: $FilePath" -Level Success
        return $true
    }
    catch {
        Write-ScriptOutput "Failed to create Excel summary file: $_" -Level Error
        return $false
    }
}
function Get-S3BucketSize {
    param([string]$BucketName, [object]$Credential, [string]$Region)

    try {
        $metricRegion = if ($script:Config.Partition -eq 'GovCloud') { $script:Config.DefaultGovCloudQueryRegion } else { 'us-east-1' }

        $commonParams = @{
            Namespace  = 'AWS/S3'
            MetricName = 'BucketSizeBytes'
            Period     = 86400
            Statistic  = 'Average'
            Credential = $Credential
        }

        function QueryMetric {
            param($QueryRegion, $StorageType)
            $dims = @(@{Name='BucketName';Value=$BucketName}, @{Name='StorageType';Value=$StorageType})
            $m = Get-SafeCWMetricStatistic -Namespace $commonParams.Namespace -MetricName $commonParams.MetricName -Dimensions $dims -Period $commonParams.Period -Statistics $commonParams.Statistic -StartTime $script:utcStartTime -EndTime $script:utcEndTime -Credential $Credential -Region $QueryRegion
            if ($m -and $m.Datapoints) {
                $dp = $m.Datapoints | Sort-Object Timestamp -Descending | Select-Object -First 1
                return $dp.Average
            }
            return $null
        }

        $val = QueryMetric -QueryRegion $metricRegion -StorageType 'AllStorageTypes'
        if ($val -and $val -gt 0) { return [double]$val }

        if ($Region -and ($Region -ne $metricRegion)) {
            $val = QueryMetric -QueryRegion $Region -StorageType 'AllStorageTypes'
            if ($val -and $val -gt 0) { return [double]$val }
        }

        $storageTypes = @(
            'StandardStorage','StandardIAStorage','OneZoneIAStorage',
            'ReducedRedundancyStorage','GlacierStorage','DeepArchiveStorage','IntelligentTieringStorage'
        )

        $totalBytes = 0
        foreach ($st in $storageTypes) {
            $v = QueryMetric -QueryRegion $metricRegion -StorageType $st
            if ($v) { $totalBytes += [double]$v }
        }

        if ($totalBytes -eq 0 -and $Region -and ($Region -ne $metricRegion)) {
            foreach ($st in $storageTypes) {
                $v = QueryMetric -QueryRegion $Region -StorageType $st
                if ($v) { $totalBytes += [double]$v }
            }
        }

        return $totalBytes
    }
    catch {
        return 0
    }
}
function Get-S3BucketSizeByStorageClass {
    param([string]$BucketName, [string]$StorageClass, [object]$Credential, [string]$Region)

    try {
        $map = @{
            'STANDARD'             = 'StandardStorage'
            'STANDARD_IA'          = 'StandardIAStorage'
            'ONEZONE_IA'           = 'OneZoneIAStorage'
            'REDUCED_REDUNDANCY'   = 'ReducedRedundancyStorage'
            'GLACIER'              = 'GlacierStorage'
            'DEEP_ARCHIVE'         = 'DeepArchiveStorage'
            'INTELLIGENT_TIERING'  = 'IntelligentTieringStorage'
        }

        $key = $StorageClass.ToUpper()
        $storageType = if ($map.ContainsKey($key)) { $map[$key] } else { $StorageClass }

        $metricRegion = if ($script:Config.Partition -eq 'GovCloud') { $script:Config.DefaultGovCloudQueryRegion } else { 'us-east-1' }

        $metric = Get-SafeCWMetricStatistic -Namespace 'AWS/S3' -MetricName 'BucketSizeBytes' -Dimensions @(@{Name='BucketName';Value=$BucketName}, @{Name='StorageType';Value=$storageType}) -Period 86400 -Statistics 'Average' -StartTime $script:utcStartTime -EndTime $script:utcEndTime -Credential $Credential -Region $metricRegion
        if ($metric -and $metric.Datapoints) {
            return ($metric.Datapoints | Sort-Object Timestamp -Descending | Select-Object -First 1).Average
        }

        if ($Region -and $Region -ne $metricRegion) {
            $metric = Get-SafeCWMetricStatistic -Namespace 'AWS/S3' -MetricName 'BucketSizeBytes' -Dimensions @(@{Name='BucketName';Value=$BucketName}, @{Name='StorageType';Value=$storageType}) -Period 86400 -Statistics 'Average' -StartTime $script:utcStartTime -EndTime $script:utcEndTime -Credential $Credential -Region $Region
            if ($metric -and $metric.Datapoints) {
                return ($metric.Datapoints | Sort-Object Timestamp -Descending | Select-Object -First 1).Average
            }
        }

        return 0
    }
    catch {
        return 0
    }
}

function Get-S3BucketSizeAccurate {
    param(
        [string]$BucketName,
        [object]$Credential,
        [string]$Region,
        [int]$MaxSeconds = 120,
        [int]$MaxObjects = 100000
    )

    try {
        if (-not $BucketName) { return 0 }

        $regionName = if ($Region) { $Region } else { 'us-east-1' }

        try {
            if (-not ([System.Type]::GetType("Amazon.S3.AmazonS3Client, AWSSDK.S3"))) {
                $mods = Get-Module -ListAvailable | Where-Object { $_.ModuleBase }
                foreach ($m in $mods) {
                    try {
                        $candidate = Get-ChildItem -Path $m.ModuleBase -Filter "AWSSDK.S3.dll" -Recurse -ErrorAction SilentlyContinue | Select-Object -First 1
                        if ($candidate) { [Reflection.Assembly]::LoadFrom($candidate.FullName) | Out-Null; break }
                    } catch { }
                }
            }
        } catch { }

        $regionEndpoint = [Amazon.RegionEndpoint]::GetBySystemName($regionName)

        if ($Credential -and ($Credential -is [Amazon.Runtime.AWSCredentials])) {
            $s3Client = [Amazon.S3.AmazonS3Client]::new($Credential, $regionEndpoint)
        } else {
            $s3Client = [Amazon.S3.AmazonS3Client]::new($regionEndpoint)
        }

        $totalBytes = 0
        $request = [Amazon.S3.Model.ListObjectsV2Request]::new()
        $request.BucketName = $BucketName
        $request.MaxKeys = 1000

        $sw = [System.Diagnostics.Stopwatch]::StartNew()
        $objectCount = 0
        $page = 0

        do {
            $page++
            $response = $s3Client.ListObjectsV2Async($request).GetAwaiter().GetResult()

            if ($response -and $response.S3Objects) {
                foreach ($obj in $response.S3Objects) {
                    if ($obj.Size -ne $null) {
                        $totalBytes += [double]$obj.Size
                        $objectCount++
                    }

                    if (($objectCount % 1000) -eq 0) {
                        Write-Progress -Activity "Enumerating S3 bucket objects" -Status "Bucket: $BucketName — Objects: $objectCount — Elapsed: $([math]::Round($sw.Elapsed.TotalSeconds,1))s" -PercentComplete 0
                        if ($sw.Elapsed.TotalSeconds -gt $MaxSeconds) {
                            return 0
                        }
                        if ($objectCount -ge $MaxObjects) {
                            return 0
                        }
                    }
                }
            }

            $request.ContinuationToken = $response.NextContinuationToken
            $isTruncated = $response.IsTruncated

            if ($sw.Elapsed.TotalSeconds -gt $MaxSeconds) {
                return 0
            }
        } while ($isTruncated)

        Write-Progress -Activity "Enumerating S3 bucket objects" -Completed
        return $totalBytes
    }
    catch {
        return 0
    }
}

function Get-EFSFileSystemSize {
    param([object]$EFS, [object]$Credential, [string]$Region)

    try {
        if ($EFS.SizeInBytes -and $EFS.SizeInBytes.Value -gt 0) {
            Write-ScriptOutput "Found direct size information (FileSystemSize object) for EFS $($EFS.FileSystemId): $([math]::Round($EFS.SizeInBytes.Value / 1GB, 2)) GB" -Level Info
            return $EFS.SizeInBytes.Value
        }

        $metric = Get-SafeCWMetricStatistic -Namespace 'AWS/EFS' -MetricName 'StorageBytes' -Dimensions @{ Name='FileSystemId'; Value=$EFS.FileSystemId } -Period 86400 -Statistics 'Average' -StartTime $script:utcStartTime -EndTime $script:utcEndTime -Credential $Credential -Region $Region

        if ($metric -and $metric.Datapoints) {
            return ($metric.Datapoints | Sort-Object Timestamp -Descending | Select-Object -First 1).Average
        }
        return 0
    }
    catch {
        return 0
    }
}

function Invoke-AWSDataCollection {
    param([object]$Credential)

    try {
        $queryRegion = if ($script:Config.Partition -eq "GovCloud") {
            $script:Config.DefaultGovCloudQueryRegion
        } else {
            $script:Config.DefaultQueryRegion
        }

        $awsRegions = Get-ProcessingRegions -Credential $Credential -QueryRegion $queryRegion
        $accountInfo = Get-AWSAccountInfo -Credential $Credential -QueryRegion $queryRegion

        if (-not $accountInfo) {
            Write-ScriptOutput "Failed to get AWS account information" -Level Error
            return
        }

        Write-ScriptOutput "Processing AWS Account: $($accountInfo.Account) ($($accountInfo.AccountAlias))" -Level Success

        Initialize-ServiceCollections -AccountId $accountInfo.Account
        $script:AccountsProcessed += $accountInfo

        $awsRegionCounter = 1
        foreach ($awsRegion in $awsRegions) {
            $awsRegion = Format-RegionName -Region $awsRegion
            Show-ScriptProgress -Activity "Region $awsRegion" -Status "Region $awsRegionCounter of $($awsRegions.Count)" -PercentComplete (($awsRegionCounter / $awsRegions.Count) * 100)

            Set-RegionPartition -Region $awsRegion

            foreach ($serviceName in $script:ServiceRegistry.Keys) {
                Invoke-ServiceInventory -ServiceName $serviceName -Credential $Credential -Region $awsRegion -AccountInfo $accountInfo -AccountAlias $accountInfo.AccountAlias
            }

            $awsRegionCounter++
        }

        Write-ScriptOutput "Region processing completed for account $($accountInfo.Account)" -Level Success

        New-AccountLevelSummary -AccountId $accountInfo.Account -AccountAlias $accountInfo.AccountAlias
    }
    catch {
        Write-ScriptOutput "Error in AWS data collection: $_" -Level Error
    }
}

function Get-ProcessingRegions {
    param([object]$Credential, [string]$QueryRegion)

    if ($Regions -and $Regions.Trim() -ne '') {
        return $Regions.Split(',') | ForEach-Object { $_.Trim() } | Where-Object { $_ }
    } else {
        try {
            $profileLocationParams = $script:Config.ProfileLocation
            return Get-EC2Region @profileLocationParams -Region $QueryRegion -Credential $Credential | Select-Object -ExpandProperty RegionName
        } catch {
            Write-ScriptOutput "Failed to list EC2 regions (query region $QueryRegion)" -Level Error
            return @()
        }
    }
}

function Get-AWSAccountInfo {
    param([object]$Credential, [string]$QueryRegion)

    Write-ScriptOutput "Getting AWS account information..." -Level Info
    try {
        $awsAccountInfo = Get-STSCallerIdentity -Credential $Credential -Region $QueryRegion -ErrorAction Stop
        $awsAccountAlias = try {
            Get-IAMAccountAlias -Credential $Credential -Region $QueryRegion -ErrorAction Stop
        } catch {
            "Unknown"
        }

        return @{
            Account = $awsAccountInfo.Account
            AccountAlias = $awsAccountAlias
        }
    } catch {
        Write-ScriptOutput "Failed to get AWS account information: $_" -Level Error
        return $null
    }
}

function Format-RegionName {
    param([string]$Region)
    return $Region.Trim().ToLower()
}

function Set-RegionPartition {
    param([string]$Region)
    if ($Region -like "us-gov-*") {
        $script:Config.Partition = "GovCloud"
        if (-not $ProfileLocation) {
            $script:Config.ProfileLocation = @{ ProfileLocation = "SAMLRoleProfile" }
        }
    } else {
        $script:Config.Partition = "Standard"
        if (-not $ProfileLocation) {
            $script:Config.ProfileLocation = @{}
        }
    }
}

function Invoke-AuthenticationScenarios {
    $authFunctions = @{
        AllLocalProfiles = {
            try {
                $profileLocationParams = $script:Config.ProfileLocation
                Write-ScriptOutput "Using profile location parameters: $($profileLocationParams | ConvertTo-Json -Compress)" -Level Info

                if ($profileLocationParams.Count -gt 0) {
                    $profileFile = $profileLocationParams.ProfileLocation
                    if (Test-Path $profileFile) {
                        Write-ScriptOutput "Reading profiles from custom location: $profileFile" -Level Info
                        $content = Get-Content $profileFile -Raw
                        $profiles = [regex]::Matches($content, '\[([^\]]+)\]') | ForEach-Object { $_.Groups[1].Value }
                        Write-ScriptOutput "Found profiles in custom file: $($profiles -join ', ')" -Level Info
                        $allProfiles = $profiles
                    } else {
                        Write-ScriptOutput "Custom profile file not found: $profileFile" -Level Error
                        return
                    }
                } else {
                    $allProfiles = Get-AWSCredential -ListProfiles
                }

                Write-ScriptOutput "Found $($allProfiles.Count) profiles to process" -Level Info

                $allProfiles | ForEach-Object {
                    Write-ScriptOutput "Processing profile: $_" -Level Info
                    try {
                        $cred = $null
                        try {
                            if ($profileLocationParams.Count -gt 0) {
                                $cred = Get-AWSCredential -ProfileName $_ @profileLocationParams -ErrorAction Stop
                                Write-ScriptOutput "Successfully loaded credentials from custom location for profile '$_'" -Level Success
                            } else {
                                $cred = Get-AWSCredential -ProfileName $_ -ErrorAction Stop
                                Write-ScriptOutput "Successfully loaded credentials from default location for profile '$_'" -Level Success
                            }
                        }
                        catch {
                            Write-ScriptOutput "Failed to load credentials for profile '$_': $($_.Exception.Message)" -Level Error

                            if ($profileLocationParams.Count -gt 0) {
                                Write-ScriptOutput "Skipping profile '$_' as custom ProfileLocation was specified but failed" -Level Warning
                                return
                            }

                            try {
                                Write-ScriptOutput "Attempting alternative credential loading for profile '$_'..." -Level Info
                                $cred = Get-AWSCredential -ProfileName $_ -ErrorAction Stop
                                Write-ScriptOutput "Successfully loaded credentials via alternative method for profile '$_'" -Level Success
                            }
                            catch {
                                Write-ScriptOutput "All credential loading methods failed for profile '$_': $($_.Exception.Message)" -Level Error
                                return
                            }
                        }

                        if ($cred) {
                            Write-ScriptOutput "Credential object obtained for profile '$_', proceeding with data collection" -Level Success
                            Invoke-AWSDataCollection -Credential $cred
                        } else {
                            Write-ScriptOutput "No credentials object returned for profile: $_" -Level Warning
                        }
                    } catch {
                        Write-ScriptOutput "Error processing profile $_`: $($_.Exception.Message)" -Level Error
                    }
                }
            }
            catch {
                Write-ScriptOutput "Error in AllLocalProfiles processing: $($_.Exception.Message)" -Level Error
            }
        }
        UserSpecifiedProfiles = {
            $UserSpecifiedProfileNames.Split(',') | ForEach-Object { $_.Trim() } | Where-Object { $_ } | ForEach-Object {
                Write-ScriptOutput "Processing profile: $_" -Level Info
                try {
                    $profileLocationParams = $script:Config.ProfileLocation
                    Write-ScriptOutput "Using profile location parameters: $($profileLocationParams | ConvertTo-Json -Compress)" -Level Info

                    $cred = $null
                    try {
                        if ($profileLocationParams.Count -gt 0) {
                            $cred = Get-AWSCredential -ProfileName $_ @profileLocationParams -ErrorAction Stop
                            Write-ScriptOutput "Successfully loaded credentials from custom location for profile '$_'" -Level Success
                        } else {
                            $cred = Get-AWSCredential -ProfileName $_ -ErrorAction Stop
                            Write-ScriptOutput "Successfully loaded credentials from default location for profile '$_'" -Level Success
                        }
                    }
                    catch {
                        Write-ScriptOutput "Failed to load credentials for profile '$_': $($_.Exception.Message)" -Level Error

                        if ($profileLocationParams.Count -gt 0) {
                            Write-ScriptOutput "Skipping profile '$_' as custom ProfileLocation was specified but failed" -Level Warning
                            return
                        }

                        try {
                            Write-ScriptOutput "Attempting alternative credential loading for profile '$_'..." -Level Info
                            $cred = Get-AWSCredential -ProfileName $_ -ErrorAction Stop
                            Write-ScriptOutput "Successfully loaded credentials via alternative method for profile '$_'" -Level Success
                        }
                        catch {
                            Write-ScriptOutput "All credential loading methods failed for profile '$_': $($_.Exception.Message)" -Level Error
                            return
                        }
                    }

                    if ($cred) {
                        Write-ScriptOutput "Credential object obtained for profile '$_', proceeding with data collection" -Level Success
                        Invoke-AWSDataCollection -Credential $cred
                    } else {
                        Write-ScriptOutput "No credentials object returned for profile: $_" -Level Warning
                    }
                } catch {
                    Write-ScriptOutput "Error processing profile $_`: $($_.Exception.Message)" -Level Error
                }
            }
        }
        CrossAccountRole = {
            function Assume-RoleOrFail {
                param($AccountId)

                Write-ScriptOutput "Processing cross-account role for account: $AccountId" -Level Info
                try {
                    Write-ScriptOutput "DEBUG: AccountId being used for role ARN: '$AccountId'" -Level Info

                    $AccountId = if ($AccountId) { $AccountId.Trim() } else { $AccountId }
                    $roleName = if ($CrossAccountRoleName) { ($CrossAccountRoleName -replace '^(?:role/|/)+','').Trim() } else { $null }

                    if (-not $AccountId -or -not $roleName -or -not ($AccountId -match '^\d{12}$')) {
                        Write-ScriptOutput "ERROR: AccountId or RoleName is missing or invalid for role assumption! AccountId='$AccountId' RoleName='$roleName'" -Level Error
                        return
                    }

                    $roleArn = "arn:aws:iam::${AccountId}:role/${roleName}"
                    Write-ScriptOutput "DEBUG: Built role ARN: $roleArn" -Level Info

                    $sessionName = if ($CrossAccountRoleSessionName) { $CrossAccountRoleSessionName } else { "CVAWS-Cost-Sizing" }
                    $stsParams = @{
                        RoleArn = $roleArn
                        RoleSessionName = $sessionName
                        ErrorAction = 'Stop'
                    }
                    if ($ExternalId) { $stsParams.ExternalId = $ExternalId }

                    $profileLocationParams = $script:Config.ProfileLocation
                    if ($profileLocationParams -and $profileLocationParams.Count -gt 0) {
                        foreach ($k in $profileLocationParams.Keys) { $stsParams[$k] = $profileLocationParams[$k] }
                    }

                    Write-ScriptOutput "Attempting STS AssumeRole for $AccountId" -Level Info
                    $assumeResult = Use-STSRole @stsParams
                    if (-not $assumeResult) {
                        throw "Failed to assume role $roleArn for account $AccountId."
                    }

                    if ($assumeResult -is [Amazon.Runtime.AWSCredentials]) {
                        Write-ScriptOutput "DEBUG: Use-STSRole returned AWSCredentials for account $AccountId" -Level Info
                        Invoke-AWSDataCollection -Credential $assumeResult
                        return
                    }

                    $creds = $null
                    if ($assumeResult -and $assumeResult.Credentials) { $creds = $assumeResult.Credentials }

                    if (-not $creds) {
                        throw "AssumeRole did not return usable credentials."
                    }

                    $basicType = [System.Type]::GetType("Amazon.Runtime.BasicSessionAWSCredentials, AWSSDK.Core")
                    if (-not $basicType) {
                        try {
                            $mods = Get-Module -ListAvailable | Where-Object { $_.ModuleBase }
                            $dllPath = $null
                            foreach ($m in $mods) {
                                try {
                                    $candidate = Get-ChildItem -Path $m.ModuleBase -Filter "AWSSDK.Core.dll" -Recurse -ErrorAction SilentlyContinue | Select-Object -First 1
                                    if ($candidate) { $dllPath = $candidate.FullName; break }
                                } catch { continue }
                            }
                            if ($dllPath) {
                                [Reflection.Assembly]::LoadFrom($dllPath) | Out-Null
                                Write-ScriptOutput "DEBUG: Loaded AWSSDK.Core from $dllPath" -Level Info
                            }
                        } catch {
                            Write-ScriptOutput "DEBUG: Could not auto-load AWSSDK.Core: $($_.Exception.Message)" -Level Warning
                        }
                        $basicType = [System.Type]::GetType("Amazon.Runtime.BasicSessionAWSCredentials, AWSSDK.Core")
                    }

                    if ($basicType) {
                        try {
                            $sessionCred = New-Object Amazon.Runtime.BasicSessionAWSCredentials(
                                $creds.AccessKeyId,
                                $creds.SecretAccessKey,
                                $creds.SessionToken
                            )
                            Write-ScriptOutput "DEBUG: Converted AssumeRoleResponse to BasicSessionAWSCredentials for account $AccountId" -Level Info
                            Invoke-AWSDataCollection -Credential $sessionCred
                            return
                        }
                        catch {
                            Write-ScriptOutput "DEBUG: Failed to construct BasicSessionAWSCredentials: $($_.Exception.Message)" -Level Warning
                        }
                    }

                    Write-ScriptOutput "DEBUG: Falling back to environment-variable based credentials for account $AccountId" -Level Info

                    $oldAWSAccessKey = $env:AWS_ACCESS_KEY_ID
                    $oldAWSSecretKey = $env:AWS_SECRET_ACCESS_KEY
                    $oldAWSSessionToken = $env:AWS_SESSION_TOKEN

                    try {
                        $env:AWS_ACCESS_KEY_ID = $creds.AccessKeyId
                        $env:AWS_SECRET_ACCESS_KEY = $creds.SecretAccessKey
                        $env:AWS_SESSION_TOKEN = $creds.SessionToken

                        Invoke-AWSDataCollection -Credential $null
                    }
                    finally {
                        if ($null -ne $oldAWSAccessKey) { $env:AWS_ACCESS_KEY_ID = $oldAWSAccessKey } else { Remove-Item Env:AWS_ACCESS_KEY_ID -ErrorAction SilentlyContinue }
                        if ($null -ne $oldAWSSecretKey) { $env:AWS_SECRET_ACCESS_KEY = $oldAWSSecretKey } else { Remove-Item Env:AWS_SECRET_ACCESS_KEY -ErrorAction SilentlyContinue }
                        if ($null -ne $oldAWSSessionToken) { $env:AWS_SESSION_TOKEN = $oldAWSSessionToken } else { Remove-Item Env:AWS_SESSION_TOKEN -ErrorAction SilentlyContinue }
                    }
                }
                catch {
                    Write-ScriptOutput "Failed to assume cross-account role for account ${AccountId}: $($_.Exception.Message)" -Level Error
                }
            }

            Write-ScriptOutput "DEBUG: UserSpecifiedAccounts='$UserSpecifiedAccounts'" -Level Info

            if ($UserSpecifiedAccounts) {
                $UserSpecifiedAccounts.Split(',') | ForEach-Object {
                    $accountId = $_.Trim()
                    if ($accountId -and $accountId -match '^\d{12}$') {
                        Assume-RoleOrFail $accountId
                    } else {
                        Write-ScriptOutput "Skipping invalid or empty account id: '$accountId'" -Level Warning
                    }
                }
            } elseif ($UserSpecifiedAccountsFile -and (Test-Path $UserSpecifiedAccountsFile)) {
                $accounts = Get-Content $UserSpecifiedAccountsFile | Where-Object { $_.Trim() -and -not $_.StartsWith('#') }
                $accounts | ForEach-Object {
                    $accountId = $_.Trim()
                    if ($accountId -and $accountId -match '^\d{12}$') {
                        Assume-RoleOrFail $accountId
                    } else {
                        Write-ScriptOutput "Skipping invalid or empty account id from file: '$accountId'" -Level Warning
                    }
                }
            } else {
                Write-ScriptOutput "Cross-account role processing requires UserSpecifiedAccounts or UserSpecifiedAccountsFile" -Level Error
            }
        }
        Default = {
            Invoke-AWSDataCollection -Credential $null
        }
    }

    $scenario = if ($AllLocalProfiles) { 'AllLocalProfiles' }
                elseif ($UserSpecifiedProfileNames) { 'UserSpecifiedProfiles' }
                elseif ($CrossAccountRoleName) { 'CrossAccountRole' }
                else { 'Default' }

    Write-ScriptOutput "Executing authentication scenario: $scenario" -Level Info
    & $authFunctions[$scenario]
}

function ComprehensiveSummary {
    try {
        Write-ScriptOutput "Creating comprehensive summary across all accounts..." -Level Info

        $comprehensiveData = @{}

        foreach ($serviceName in $script:ServiceRegistry.Keys) {
            $allServiceData = @()
            foreach ($accountId in $script:ServiceDataByAccount.Keys) {
                $accountData = $script:ServiceDataByAccount[$accountId][$serviceName]
                if ($accountData) {
                    $allServiceData += $accountData
                }
            }
            $comprehensiveData[$serviceName] = $allServiceData
        }

        $comprehensiveFile = "comprehensive_all_aws_accounts_summary_$date_string.xlsx"
        $dataSheets = [ordered]@{}

        foreach ($serviceName in $script:ServiceRegistry.Keys) {
            $serviceList = $comprehensiveData[$serviceName]
            $summarySheetName = if ($serviceName -eq "UnattachedVolumes") {
                "Unattached Volume Summary"
            } elseif ($serviceName -eq "FSX_SVM") {
                "FSx SVM Summary"
            }else {
                "$serviceName Summary"
            }
            if ($serviceList -and $serviceList.Count -gt 0) {
                $dataSheets[$summarySheetName] = MultiAccountComprehensiveSummaryData -ServiceName $serviceName -ServiceList $serviceList
            }        
        }

        $detailSheets = @{
            "EC2 Details" = $comprehensiveData["EC2"]
            "S3 Details" = $comprehensiveData["S3"]
            "Unattached Volumes" = $comprehensiveData["UnattachedVolumes"]
            "EFS Details" = $comprehensiveData["EFS"]
            "FSx Details" = $comprehensiveData["FSX"]
            "FSx SVM Details" = $comprehensiveData["FSX_SVM"]
            "RDS Details" = $comprehensiveData["RDS"]
            "EKS Details" = $comprehensiveData["EKS"]
            "DynamoDB Details" = $comprehensiveData["DynamoDB"]
            "Redshift Details" = $comprehensiveData["Redshift"]
        }

        foreach ($sheetName in $detailSheets.Keys) {
            if ($detailSheets[$sheetName] -and $detailSheets[$sheetName].Count -gt 0) {
                if ($sheetName -eq "S3 Details") {
                    $dataSheets[$sheetName] = $detailSheets[$sheetName] |
                        Select-Object AwsAccountId, AwsAccountAlias, Region, BucketName, CreationDate, SizeGiB, SizeTiB, SizeGB, SizeTB
                }
                elseif ($sheetName -eq "FSx SVM Details") {
                    $svmItems = $detailSheets[$sheetName]
                    $tagProps = ($svmItems | ForEach-Object { $_.PSObject.Properties.Name } | Where-Object { $_ -like 'Tag_*' } | Sort-Object -Unique)
                    $cols = @(
                        'AwsAccountId','AwsAccountAlias','Region','FileSystemId','StorageVirtualMachineId','Name',
                        'OperationalState','UUID','VolumeCount',
                        @{Name='SizeGiB';Expression={ $_.VolumesSizeGiB }},
                        @{Name='SizeTiB';Expression={ $_.VolumesSizeTiB }},
                        @{Name='SizeGB';Expression={ $_.VolumesSizeGB }},
                        @{Name='SizeTB';Expression={ $_.VolumesSizeTB }},
                        'VolumeDetails'
                    ) + $tagProps
                    $dataSheets[$sheetName] = $svmItems | Select-Object $cols
                }
                else {
                    $dataSheets[$sheetName] = $detailSheets[$sheetName]
                }
            }
        }

        $result = Export-DataToExcel -FilePath $comprehensiveFile -DataSheets $dataSheets

        if ($result) {
            Write-ScriptOutput "Comprehensive summary created: $comprehensiveFile" -Level Success
            $script:AllOutputFiles.Add($comprehensiveFile) | Out-Null
        }
    }
    catch {
        Write-ScriptOutput "Error creating comprehensive summary: $_" -Level Error
    }
}

function MultiAccountComprehensiveSummaryData {
    param(
        [string]$ServiceName,
        [array]$ServiceList
    )

    $serviceConfig = $script:ServiceRegistry[$ServiceName]
    $summaryData = @()

    if (-not $ServiceList -or $ServiceList.Count -eq 0) {
        $displayName = "$ServiceName $($serviceConfig.DisplayName)"
        $row = [ordered]@{
            "ResourceType" = $displayName
            "Region" = "All"
            "Count" = 0
            "Total Size (GiB)" = 0
            "Total Size (GB)" = 0
            "Total Size (TiB)" = 0
            "Total Size (TB)" = 0
        }
        if ($ServiceName -eq "DynamoDB") {
            $row["Total Table Size (Bytes)"] = 0
            $row["Total Item Count"] = 0
        }
        if ($ServiceName -eq "EKS") {
            $row["Total Node Count"] = 0
        }
        $summaryData += [PSCustomObject]$row
        return $summaryData
    }

    $accounts = $ServiceList | Select-Object -ExpandProperty AwsAccountId -Unique | Sort-Object

    foreach ($accountId in $accounts) {
        $accountItems = $ServiceList | Where-Object { $_.AwsAccountId -eq $accountId }
        $accountTotals = if ($ServiceName -eq "S3") {
            Get-S3ServiceTotals -ServiceList $accountItems
        } elseif ($ServiceName -eq "FSX") {
            Get-FSxServiceTotals -ServiceList $accountItems
        } else {
            Get-StandardServiceTotals -ServiceList $accountItems -SizeProperty $serviceConfig.SizeProperty
        }

        $row = [ordered]@{
            "ResourceType" = "`u{200B}$accountId"
            "Region" = "All"
            "Count" = $accountItems.Count
            "Total Size (GiB)" = [math]::Round($accountTotals.GiB, 3)
            "Total Size (GB)" = [math]::Round($accountTotals.GB, 3)
            "Total Size (TiB)" = [math]::Round($accountTotals.TiB, 4)
            "Total Size (TB)" = [math]::Round($accountTotals.TB, 4)
        }

        if ($ServiceName -eq "DynamoDB") {
            $acctBytes = ($accountItems | ForEach-Object { if ($_.TableSizeBytes -ne $null -and $_.TableSizeBytes -ne '') { [double]$_.TableSizeBytes } else { 0 } } | Measure-Object -Sum).Sum
            $acctItems  = ($accountItems | ForEach-Object { if ($_.ItemCount -ne $null -and $_.ItemCount -ne '') { [double]$_.ItemCount } else { 0 } } | Measure-Object -Sum).Sum
            $row["Total Table Size (Bytes)"] = $acctBytes
            $row["Total Item Count"] = $acctItems
        }

        $summaryData += [PSCustomObject]$row
    }

    $grandTotals = if ($ServiceName -eq "S3") {
        Get-S3ServiceTotals -ServiceList $ServiceList
    } elseif ($ServiceName -eq "FSX") {
        Get-FSxServiceTotals -ServiceList $ServiceList
    } else {
        Get-StandardServiceTotals -ServiceList $ServiceList -SizeProperty $serviceConfig.SizeProperty
    }

    $totalRow = [ordered]@{
        "ResourceType" = "Total"
        "Region" = "All"
        "Count" = $ServiceList.Count
        "Total Size (GiB)" = [math]::Round($grandTotals.GiB, 3)
        "Total Size (GB)" = [math]::Round($grandTotals.GB, 3)
        "Total Size (TiB)" = [math]::Round($grandTotals.TiB, 4)
        "Total Size (TB)" = [math]::Round($grandTotals.TB, 4)
    }
    if ($ServiceName -eq "DynamoDB") {
        $grandBytes = ($ServiceList | ForEach-Object { if ($_.TableSizeBytes -ne $null -and $_.TableSizeBytes -ne '') { [double]$_.TableSizeBytes } else { 0 } } | Measure-Object -Sum).Sum
        $grandItems = ($ServiceList | ForEach-Object { if ($_.ItemCount -ne $null -and $_.ItemCount -ne '') { [double]$_.ItemCount } else { 0 } } | Measure-Object -Sum).Sum
        $totalRow["Total Table Size (Bytes)"] = $grandBytes
        $totalRow["Total Item Count"] = $grandItems
    }
    if ($ServiceName -eq "EKS") {
        $grandNodeCount = ($ServiceList | ForEach-Object { if ($_.NodeCount -ne $null) { [int]$_.NodeCount } else { 0 } } | Measure-Object -Sum).Sum
        $totalRow["Total Node Count"] = $grandNodeCount
    }
    $summaryData += [PSCustomObject]$totalRow

    $summaryData += Get-EmptyRow
    $summaryData += Get-EmptyRow

    $summaryData += Get-MultiAccountRegionalBreakdown -ServiceName $ServiceName -ServiceList $ServiceList -Accounts $accounts

    return $summaryData
}

function Get-MultiAccountRegionalBreakdown {
    param([string]$ServiceName, [array]$ServiceList, [array]$Accounts)

    $regions = if ($ServiceList) {
        $ServiceList | Select-Object -ExpandProperty Region -Unique | Sort-Object
    } else { @() }

    if ($regions.Count -eq 0) { return @() }

    $breakdownData = @()

    $header = [ordered]@{
        "ResourceType" = "--- $($ServiceName.ToUpper()) REGIONAL BREAKDOWN ---"
        "Region" = ""
        "Count" = ""
        "Total Size (GiB)" = ""
        "Total Size (GB)" = ""
        "Total Size (TiB)" = ""
        "Total Size (TB)" = ""
    }
    if ($ServiceName -eq "DynamoDB") {
        $header["Total Table Size (Bytes)"] = ""
        $header["Total Item Count"] = ""
    }
    if ($ServiceName -eq "EKS") {
        $header["Total Node Count"] = ""
    }
    $breakdownData += [PSCustomObject]$header

    $colHeader = [ordered]@{
        "ResourceType" = "Region"
        "Region" = "Account"
        "Count" = "Count"
        "Total Size (GiB)" = "Total Size (GiB)"
        "Total Size (GB)" = "Total Size (GB)"
        "Total Size (TiB)" = "Total Size (TiB)"
        "Total Size (TB)" = "Total Size (TB)"
    }
    if ($ServiceName -eq "DynamoDB") {
        $colHeader["Total Table Size (Bytes)"] = "Total Table Size (Bytes)"
        $colHeader["Total Item Count"] = "Total Item Count"
    }
    if ($ServiceName -eq "EKS") {
        $colHeader["Total Node Count"] = "Total Node Count"
    }
    $breakdownData += [PSCustomObject]$colHeader

    foreach ($region in $regions) {
        $regionItems = $ServiceList | Where-Object Region -eq $region

        if ($regionItems.Count -gt 0) {
            $regionTotals = if ($ServiceName -eq "S3") {
                Get-S3ServiceTotals -ServiceList $regionItems
            } elseif ($ServiceName -eq "FSX") {
                Get-FSxServiceTotals -ServiceList $regionItems
            } else {
                Get-StandardServiceTotals -ServiceList $regionItems -SizeProperty $script:ServiceRegistry[$ServiceName].SizeProperty
            }

            $regionRow = [ordered]@{
                "ResourceType" = $region
                "Region" = ""
                "Count" = $regionItems.Count
                "Total Size (GiB)" = [math]::Round($regionTotals.GiB, 3)
                "Total Size (GB)" = [math]::Round($regionTotals.GB, 3)
                "Total Size (TiB)" = [math]::Round($regionTotals.TiB, 4)
                "Total Size (TB)" = [math]::Round($regionTotals.TB, 4)
            }
            if ($ServiceName -eq "DynamoDB") {
                $regionBytes = ($regionItems | ForEach-Object { if ($_.TableSizeBytes -ne $null -and $_.TableSizeBytes -ne '') { [double]$_.TableSizeBytes } else { 0 } } | Measure-Object -Sum).Sum
                $regionItemsCount = ($regionItems | ForEach-Object { if ($_.ItemCount -ne $null -and $_.ItemCount -ne '') { [double]$_.ItemCount } else { 0 } } | Measure-Object -Sum).Sum
                $regionRow["Total Table Size (Bytes)"] = $regionBytes
                $regionRow["Total Item Count"] = $regionItemsCount
            }
            if ($ServiceName -eq "EKS") {
                $regionNodeCount = ($regionItems | ForEach-Object { if ($_.NodeCount -ne $null) { [int]$_.NodeCount } else { 0 } } | Measure-Object -Sum).Sum
                $regionRow["Total Node Count"] = $regionNodeCount
            }
            $breakdownData += [PSCustomObject]$regionRow

            foreach ($accountId in $Accounts) {
                $regionAccountItems = $regionItems | Where-Object { $_.AwsAccountId -eq $accountId }
                if ($regionAccountItems.Count -gt 0) {
                    $accountTotals = if ($ServiceName -eq "S3") {
                        Get-S3ServiceTotals -ServiceList $regionAccountItems
                    } elseif ($ServiceName -eq "FSX") {
                        Get-FSxServiceTotals -ServiceList $regionAccountItems
                    } else {
                        Get-StandardServiceTotals -ServiceList $regionAccountItems -SizeProperty $script:ServiceRegistry[$ServiceName].SizeProperty
                    }

                    $acctRow = [ordered]@{
                        "ResourceType" = ""
                        "Region" = "`u{200B}$accountId"
                        "Count" = $regionAccountItems.Count
                        "Total Size (GiB)" = [math]::Round($accountTotals.GiB, 3)
                        "Total Size (GB)" = [math]::Round($accountTotals.GB, 3)
                        "Total Size (TiB)" = [math]::Round($accountTotals.TiB, 4)
                        "Total Size (TB)" = [math]::Round($accountTotals.TB, 4)
                    }
                    if ($ServiceName -eq "DynamoDB") {
                        $acctBytes = ($regionAccountItems | ForEach-Object { if ($_.TableSizeBytes -ne $null -and $_.TableSizeBytes -ne '') { [double]$_.TableSizeBytes } else { 0 } } | Measure-Object -Sum).Sum
                        $acctItems  = ($regionAccountItems | ForEach-Object { if ($_.ItemCount -ne $null -and $_.ItemCount -ne '') { [double]$_.ItemCount } else { 0 } } | Measure-Object -Sum).Sum
                        $acctRow["Total Table Size (Bytes)"] = $acctBytes
                        $acctRow["Total Item Count"] = $acctItems
                    }
                    if ($ServiceName -eq "EKS") {
                        $accountRegionNodeCount = ($regionAccountItems | ForEach-Object { if ($_.NodeCount -ne $null) { [int]$_.NodeCount } else { 0 } } | Measure-Object -Sum).Sum
                        $acctRow["Total Node Count"] = $accountRegionNodeCount
                    }
                    $breakdownData += [PSCustomObject]$acctRow
                }
            }
        }
    }

    return $breakdownData
}

function New-OutputArchive {
    try {
        Write-ScriptOutput "Creating output archive..." -Level Info

        if (Test-Path $archiveFile) {
            Remove-Item $archiveFile -Force
        }

        $filesToArchive = $script:AllOutputFiles | Where-Object { Test-Path $_ }

        if ($filesToArchive.Count -gt 0) {
            Add-Type -AssemblyName System.IO.Compression.FileSystem

            $archive = [System.IO.Compression.ZipFile]::Open($archiveFile, [System.IO.Compression.ZipArchiveMode]::Create)

            foreach ($file in $filesToArchive) {
                $fileName = Split-Path $file -Leaf
                Write-ScriptOutput "Adding $fileName to archive..." -Level Info
                [System.IO.Compression.ZipFileExtensions]::CreateEntryFromFile($archive, $file, $fileName) | Out-Null
            }

            $archive.Dispose()

            Write-ScriptOutput "Archive created: $archiveFile" -Level Success

            if (Test-Path $archiveFile) {
                foreach ($file in $filesToArchive) {
                    try {
                        Remove-Item $file -Force
                        Write-ScriptOutput "Removed individual file: $file" -Level Info
                    }
                    catch {
                        Write-ScriptOutput "Could not remove individual file $file`: $_" -Level Warning
                    }
                }
            }
        }
        else {
            Write-ScriptOutput "No files found to archive" -Level Warning
        }
    }
    catch {
        Write-ScriptOutput "Error creating output archive: $_" -Level Error
    }
}


try {
    Write-ScriptOutput "=== AWS Cost Sizing Analysis Started ===" -Level Success
    Write-ScriptOutput "Log file: $script:LogFile" -Level Info
    Write-ScriptOutput "Script supports the following services: $($script:ServiceRegistry.Keys -join ', ')" -Level Info

    $requiredModules = @(
    'ImportExcel',
    'AWS.Tools.Common',
    'AWS.Tools.EC2',
    'AWS.Tools.S3',
    'AWS.Tools.SecurityToken',
    'AWS.Tools.IdentityManagement',
    'AWS.Tools.CloudWatch',
    'AWS.Tools.RDS',
    'AWS.Tools.DynamoDBv2',
    'AWS.Tools.Redshift',
    'AWS.Tools.FSx',
    'AWS.Tools.ElasticFileSystem',
    'AWS.Tools.EKS'
    )

    $missingModules = @()
    foreach ($module in $requiredModules) {
        if (-not (Get-Module -ListAvailable -Name $module)) {
            $missingModules += $module
        }
    }

    if ($missingModules.Count -gt 0) {
        Write-ScriptOutput "Necessary modules not present. Missing modules: $($missingModules -join ', ')" -Level Error
        exit 1
    }

    if ($ProfileLocation) {
        Write-ScriptOutput "ProfileLocation parameter provided: $ProfileLocation" -Level Info
        if (Test-Path $ProfileLocation) {
            Write-ScriptOutput "ProfileLocation file exists and is accessible" -Level Success
        } else {
            Write-ScriptOutput "ProfileLocation file does not exist or is not accessible: $ProfileLocation" -Level Warning
        }
    } else {
        Write-ScriptOutput "No ProfileLocation parameter provided, using default AWS credentials" -Level Info
    }

    if ($UserSpecifiedProfileNames) {
        Write-ScriptOutput "UserSpecifiedProfileNames: $UserSpecifiedProfileNames" -Level Info
        $profileList = $UserSpecifiedProfileNames.Split(',') | ForEach-Object { $_.Trim() } | Where-Object { $_ }
        Write-ScriptOutput "Parsed profiles: $($profileList -join ', ')" -Level Info
    }

    Invoke-AuthenticationScenarios

    Write-ScriptOutput "=== Processing Summary ===" -Level Success
    Write-ScriptOutput "Accounts processed: $($script:AccountsProcessed.Count)" -Level Info

    if ($script:AccountsProcessed.Count -gt 0) {
        ComprehensiveSummary
    }

    foreach ($account in $script:AccountsProcessed) {
        Write-ScriptOutput "Account: $($account.Account) ($($account.AccountAlias))" -Level Info
        foreach ($serviceName in $script:ServiceRegistry.Keys) {
            $count = $script:ServiceDataByAccount[$account.Account][$serviceName].Count
            Write-ScriptOutput "  $serviceName`: $count items" -Level Info
        }
    }

    if ($SelectiveZipping -and $script:AllOutputFiles.Count -gt 1) {
        New-OutputArchive
    }

    if ($script:KubectlInstalled -and $script:KubectlDir) {
        try {
            $env:Path = ($env:Path -split ';' | Where-Object { $_ -ne $script:KubectlDir }) -join ';'
            if (Test-Path $script:KubectlDir) {
                Remove-Item $script:KubectlDir -Recurse -Force
                Write-ScriptOutput "kubectl uninstalled successfully from $script:KubectlDir." -Level Info
            }
        } catch {
            Write-ScriptOutput "Failed to uninstall kubectl: $_" -Level Warning
        }
    }

    $executionTime = (Get-Date) - $script:startTime
    Write-ScriptOutput "Total execution time: $($executionTime.ToString('hh\:mm\:ss'))" -Level Success
    Write-ScriptOutput "Log file created: $script:LogFile" -Level Info
    Write-ScriptOutput "=== AWS Cost Sizing Analysis Completed ===" -Level Success
}
catch {
    Write-ScriptOutput "Critical error in main execution: $_" -Level Error
    exit 1
}
