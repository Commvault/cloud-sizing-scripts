# Google Cloud Sizing - Quick Guide

Use this script:

- `CVGoogleCloudSizingScript.ps1`

This guide keeps only the most common and basic usage patterns.

## Quick Start

1. Open PowerShell in this folder:

```powershell
cd .\GoogleCloud
```

2. Run with a service account key:

```powershell
.\CVGoogleCloudSizingScript.ps1 -ServiceAccount "C:\path\to\service-account.json"
```

3. Check the output path shown at the end of the run.

## Basic Commands

## 1) Run all supported workloads

```powershell
.\CVGoogleCloudSizingScript.ps1 -ServiceAccount "C:\path\to\service-account.json"
```

## 2) Run only selected workload types

```powershell
.\CVGoogleCloudSizingScript.ps1 -ServiceAccount "C:\path\to\service-account.json" -Types VM,Storage,Fileshare
```

## 3) Run only selected projects

```powershell
.\CVGoogleCloudSizingScript.ps1 -ServiceAccount "C:\path\to\service-account.json" -Projects "proj-a,proj-b"
```

## 4) Skip selected projects

```powershell
.\CVGoogleCloudSizingScript.ps1 -ServiceAccount "C:\path\to\service-account.json" -Projects "proj-a,proj-b,proj-c" -SkipProjects "proj-b"
```

## 5) Keep output folder and skip ZIP

```powershell
.\CVGoogleCloudSizingScript.ps1 -ServiceAccount "C:\path\to\service-account.json" -SkipZip -KeepOutputDir
```

## Notes

- `Projects` accepts comma-separated values.
- `Types` accepts comma-separated values.
- To skip projects from a file, use `-SkipProjectsFile <file>` (one project per line).

