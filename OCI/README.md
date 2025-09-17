# OCI Cloud Sizing Script

This script collects and summarizes Oracle Cloud Infrastructure (OCI) resource data for Compute Instances and Object Storage Buckets across multiple regions and compartments. It generates an Excel report with detailed and summary information, and logs the process for auditing and troubleshooting.

## Features

- **Compute Instances:** Lists all active instances, their shapes, volumes, and storage usage.
- **Object Storage:** Lists all buckets, their storage tier, object count, and total size.
- **Multi-region and multi-compartment support.**
- **Excel output:** Generates a formatted `.xlsx` file with resource details and summary.
- **Self-bootstrapping:** Automatically installs required Python packages if missing.

## Requirements

The script will attempt to install these packages automatically if they are not present, but you can install them manually with pip if you prefer:
- `oci`
- `openpyxl`
- `pandas`
- [OCI configuration file (`~/.oci/config`)](https://docs.oracle.com/en-us/iaas/Content/API/Concepts/sdkconfig.htm)

**Important:**  
Ensure that the user or profile specified in your OCI configuration has the necessary IAM permissions to list, read, and inspect Compute and Object Storage resources in the target compartments and regions. Without these permissions, the script will not be able to retrieve resource information.

## Usage


**Note:**
- `--workload` is **optional**. If not provided, the script will retrieve information for all supported workloads (instances, object_storage, db_systems, oke_clusters).
- If you do not specify `--profile`, the script uses the `DEFAULT` profile from your OCI config.
- If you do not specify `--region`, the script processes all subscribed regions.
- If you do not specify `--compartment`, the script processes all compartments in your tenancy.
- For faster report generation, it is recommended you set the required params based on your workload.
## OKE Cluster Requirements

To retrieve information for OKE clusters (workload `oke_clusters`), you must have the following installed and available in your system PATH:

- **OCI CLI**: [Install OCI CLI](https://docs.oracle.com/en-us/iaas/Content/API/SDKDocs/cliinstall.htm)
- **kubectl**: [Install kubectl](https://kubernetes.io/docs/tasks/tools/)

Both tools must be accessible from the command line for the script to collect OKE cluster and Kubernetes resource information.

### Basic Usage

```sh
python oci_cs.py
```

### Specify Workload, Profile, Regions, and Compartments

```sh
python oci_cs.py --workload=instances --profile=MY_PROFILE --region=us-ashburn-1,us-phoenix-1 --compartment=ocid1.compartment.oc1..xxxx,ocid1.compartment.oc1..yyyy
python oci_cs.py --workload=object_storage --profile=MY_PROFILE --region=us-ashburn-1 --compartment=ocid1.compartment.oc1..xxxx
```

### Show Help

```sh
python oci_cs.py --help
```
Displays usage instructions and exits.

## Arguments

- `--workload=<instances|object_storage|db_systems|oke_clusters>`: Type of resource to report. Defaults to getting all supported workloads
- `--profile=<profile_name>`: Optional. OCI config profile name. Defaults to `DEFAULT`.
- `--region=<region1>,<region2>`: Optional. Comma-separated list of regions. If omitted, all subscribed regions are processed.
- `--compartment=<compartment_ocid1>,<compartment_ocid2>`: Optional. Comma-separated list of compartment OCIDs. If omitted, all compartments are processed.
- `--help`: Show usage information.

## Output

- **Excel file:** Saved in the `Metrics` directory, named with profile, workload, and timestamp.
- **Log file:** Saved in the `Logs` directory, named with profile, workload, and timestamp.

## Example

```sh
python oci_cs.py --workload=instances --profile=DEFAULT
python oci_cs.py --workload=object_storage --region=us-ashburn-1
```

## Notes

- Ensure your OCI config file is set up and you have the necessary permissions.
- The script will install required Python packages if missing, but you can install them manually with pip if you prefer.
- Make sure your OCI user/profile has the required IAM policies to list and read Compute and Object Storage resources in the target compartments