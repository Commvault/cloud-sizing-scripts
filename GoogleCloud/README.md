### Google Cloud - Execution Instructions

Below are two ways to run the Google Cloud sizing script. Method 1 (Local Powershell) is the recommended as Google Cloud Shell timeouts long processes leading to non graceful terminations of the execution.

### Environment Definitions
The term "environment" here refers to the scale of the Google Cloud (GC) environment in use. The scale is categorized based on the number of projects, as defined below:

1. Small-Scale GC Environments: Up to 10 projects

2. Medium-Scale GC Environments: Up to 50 projects

3. Large-Scale GC Environments: More than 50 projects

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

