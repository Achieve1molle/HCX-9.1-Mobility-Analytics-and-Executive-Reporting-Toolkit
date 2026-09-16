# HCX 9.1 Mobility Analytics and Executive Reporting Toolkit

PowerShell 7 and WPF reporting utility for collecting VMware HCX 9.1 mobility-group activity, correlating VM-level migration records, and producing an executive HTML report with supporting CSV, JSON, Excel, logging, and audit artifacts.

**Release:** 1.0  
**Primary script:** `HCX91_Mobility_Analytics_Executive_Report_Rev_1.0.ps1`  
**Target platform:** Windows with PowerShell 7 or later  
**HCX scope:** VMware HCX 9.1 in a VMware Cloud Foundation 9 environment  
**Primary interface:** HCX REST API  
**User interface:** Windows Presentation Foundation, WPF

## Overview

The toolkit presents a desktop interface for connecting to an HCX 9.1 Manager, selecting a reporting period, entering a migration wave name and estimated VM total, and generating a consolidated migration report.

The report combines:

- Mobility-group status and scheduled VM totals
- Successfully migrated, failed, and cancelled VM records
- Wave completion against an operator-entered estimate
- Migrated vCPU, memory, and provisioned disk totals
- Individual VM migration duration statistics
- Daily migration activity
- Top-five migration charts
- Compute Transfer Score analysis
- HCX service-mesh transport analytics
- Path MTU measurements
- Destination network mappings and placement
- Error and later-recovery status
- Detailed time-filtered VM migration records

The toolkit is reporting-only. It does not create, start, cancel, or modify HCX migrations.

## Key Features

### Wave progress reporting

Operators provide:

- Wave name, such as `Wave 0`
- Estimated total VMs planned for the wave
- Reporting start date
- Reporting end date

The report compares the wave estimate with unique successfully migrated VM names and displays:

- Unique VMs completed
- Estimated wave VMs
- Estimated VMs remaining
- Completion percentage
- Circular completion visualization

Repeated successful records for the same VM do not inflate wave completion.

### Executive migration metrics

The report header includes:

- Mobility Groups
- Total Scheduled VMs
- Successfully Migrated VMs
- Error VMs
- Cancelled VMs
- Migrated vCPU
- Migrated Memory in TB
- Migrated Disk in TB
- Average VM Migration Time
- Maximum VM Migration Time

Resource totals use all successful migration records within the selected reporting period. Average and maximum migration times use individual successful VM durations, not mobility-group elapsed time.

### Migration charts

The HTML report includes:

- Daily Migrated VM Count
- Top Five VMs by Migration Time
- Top Five VMs by Storage Transferred
- Top Five VMs by Compute Transfer Score

Chart cards are sized for up to five displayed rows and avoid unnecessary horizontal scrolling.

### Compute Transfer Score

The Compute Transfer Score is a comparative workload-effort index:

```text
Compute Transfer Score = vCPU x max(Migration Minutes, 1)
```

Example:

```text
4 vCPU x 30 minutes = 120
```

A higher score represents a larger combination of assigned vCPU and migration duration. The score is not network throughput, CPU utilization, HCX health, or a pass/fail rating.

### Transport Analytics

The report retrieves HCX service-mesh transport data and presents:

- Available upload bandwidth
- Available download bandwidth
- Latency
- Packet loss
- Uplink name
- Current Path MTU
- HCX service health states

Service health is displayed as a readable list alongside centered transport metrics.

### Path MTU reporting

The toolkit maps the HCX Path MTU response fields as follows:

- `discoveredMtu` to Discovered MTU
- `currentMtu` to Current MTU
- `configuredMtu` to Configured MTU
- `pmtuMetrics[].rxPmtu` to Receive Path MTU
- `pmtuMetrics[].txPmtu` to Transmit Path MTU

The service-mesh Path MTU request uses:

```text
/hybridity/api/interconnect/underlay/pmtu/serviceMesh/{serviceMeshId}?vcGuid={vcGuid}
```

### Destination network and placement reporting

The toolkit enriches VM records with HCX mobility intent information, including:

- Destination network name
- Destination network identifier
- Destination network type
- Source-to-destination network mapping
- Destination placement

For a successful VM record without a populated destination network, the report can use another collected record for the same VM entity ID or VM name. The report omits empty optional network content rather than displaying invalid placeholders.

### Scrollable detail tables

Detailed report sections use fixed-height, independently scrollable table regions with:

- Vertical scrolling
- Horizontal scrolling for wide datasets
- Sticky headers
- Nonwrapping cells
- Scroll containment

This supports production waves containing long VM, mobility-group, network, and error lists.

## Requirements

- Windows workstation or administrative desktop
- PowerShell 7 or later
- Interactive desktop session
- STA apartment state for WPF
- HTTPS connectivity to the HCX Manager
- HCX account with permission to read the required reporting resources
- Write access to the script directory or selected output base path
- Optional `ImportExcel` PowerShell module for `.xlsx` output

The script relaunches itself in PowerShell 7 STA mode when needed.

## Installation

1. Download the release script.
2. Save the script to an approved administrative directory, such as `C:\Script`.
3. Optionally install the `ImportExcel` module if Excel output is required.
4. Open PowerShell 7 in an interactive Windows session.
5. Run the script according to organizational execution-policy and code-signing requirements.

```powershell
Set-Location C:\Script

& '.\HCX91_Mobility_Analytics_Executive_Report_Rev_1.0.ps1'
```

## Usage

### 1. Connect to HCX

Enter:

- HCX Manager FQDN or IP address
- HCX username
- HCX password

Select **Connect**. The application authenticates through the HCX session workflow and retains the session token in memory.

### 2. Define the reporting scope

Select:

- Start Date
- End Date
- Output Base folder
- Wave Name
- Estimated Wave VMs

The estimated wave total must be a whole number greater than zero.

### 3. Generate the report

Select **Collect and Generate Reports**.

The application:

1. Discovers the HCX mobility-group reporting endpoint.
2. Retrieves mobility-group summaries.
3. Retrieves VM-level migration records for each group.
4. Filters records to the selected reporting period.
5. Enriches network and placement information.
6. Retrieves service-mesh transport and Path MTU information.
7. Builds the report model and exports all artifacts.
8. Validates required artifacts before reporting success.

### 4. Review output

Use:

- **Open HTML** to open the executive report
- **Open Run Folder** to open the complete evidence directory

## Output Directory Structure

Each application launch creates a timestamped directory:

```text
HCX91-Mobility-Analytics-Run-YYYYMMDD-HHMMSS
```

Expected structure:

```text
HCX91-Mobility-Analytics-Run-YYYYMMDD-HHMMSS/
├── Configuration/
├── Debug-Artifacts/
├── Exports/
├── Logs/
├── Raw-API/
└── Reports/
```

### Reports

- Executive HTML report
- Artifact manifest with SHA-256 hashes

### Exports

- Raw migration CSV
- Summary CSV
- Daily migration CSV
- Daily Guest OS CSV, when retained by the release
- Mobility Group Summary CSV
- Excel workbook when `ImportExcel` is available

### Logs and evidence

- Application log
- PowerShell transcript
- Sanitized REST request artifacts
- Sanitized REST response artifacts
- Endpoint-discovery evidence
- Raw normalized collection JSON

## Excel Workbook

When `ImportExcel` is installed, the workbook includes worksheets for:

- All Migrations
- Migrated
- Errors
- Daily
- Daily Guest OS
- Mobility Groups

If `ImportExcel` is unavailable or Excel generation fails, CSV and HTML outputs remain available.

## Security

- Passwords and tokens are not intentionally written to disk.
- Diagnostic artifacts pass through sanitization designed to mask credentials, tokens, authorization headers, cookies, and the active password.
- HCX session information remains in memory for the active application session.
- Output artifacts can contain infrastructure names, VM names, managed object identifiers, network names, service-mesh identifiers, error details, and administrative endpoints.
- Treat every run folder as sensitive operational data.
- Do not commit production run folders, logs, raw API captures, CSV files, JSON artifacts, or reports to a public repository.

## Troubleshooting

### PowerShell 7 is required

Install PowerShell 7 or run the script from a system where `pwsh.exe` is available.

### WPF or STA error

Run the application from an interactive Windows session. The script cannot run as a Linux utility or in a noninteractive service context.

### HCX authentication fails

- Confirm the HCX Manager address.
- Confirm TCP 443 connectivity.
- Re-enter the username and password.
- Confirm the account can create an HCX API session.
- Review the application log and sanitized REST artifacts.

### No mobility records are returned

- Confirm the selected date range.
- Confirm the configured vCenter GUID matches the HCX environment.
- Confirm the account can read mobility groups and migrations.
- Review endpoint-discovery evidence and REST responses.

### Path MTU is blank

- Confirm Transport Analytics returned a service-mesh ID.
- Review the Path MTU REST request and response artifacts.
- Confirm the response includes matching local or remote uplink names.
- Confirm the response includes `currentMtu`, `configuredMtu`, `discoveredMtu`, and `pmtuMetrics`.

### Destination network is blank

Network information appears only when HCX returns a usable mobility-intent mapping or another collected record for the same VM provides a reliable destination network.

### Excel output is missing

Install the optional module:

```powershell
Install-Module ImportExcel -Scope CurrentUser
```

Rerun the report after the module is available.

### Table requires additional rows

All HTML tables have independent vertical and horizontal scrolling. Move the pointer over the table and use the mouse wheel, scrollbar, Page Up, or Page Down.

## Release Notes

### Release 1.0

- Added PowerShell 7 WPF operator interface.
- Added HCX REST authentication using the HCX session workflow.
- Added per-launch run folders.
- Added sanitized debug logging and PowerShell transcripts.
- Added mobility-group discovery and VM-level migration collection.
- Added date filtering.
- Added Wave Name and Estimated Wave VMs inputs.
- Added unique-VM wave completion reporting.
- Added executive metrics for migration status and migrated resources.
- Added migrated memory and disk totals in TB.
- Added individual VM average and maximum migration time.
- Added daily and Top Five charts.
- Added Compute Transfer Score and administrator guidance.
- Added service-mesh transport analytics.
- Added Path MTU collection and mapping.
- Added destination network and placement enrichment.
- Added error-recovery correlation.
- Added independently scrollable detail tables.
- Added CSV, JSON, HTML, Excel, and artifact-manifest output.
- Added physical artifact validation and SHA-256 manifest generation.

## Recommended Repository Layout

```text
/
├── HCX91_Mobility_Analytics_Executive_Report_Rev_1.0.ps1
├── README.md
├── Wiki.md
├── LICENSE
├── screenshots/
│   ├── application.png
│   └── executive-report.png
└── examples/
    └── README.md
```

Use synthetic or sanitized values in all examples and screenshots.

## Support and Contributions

When reporting an issue, include:

- PowerShell version
- Windows version
- HCX build information
- Script release
- Sanitized application log
- Sanitized exception artifact
- Relevant sanitized REST response shape
- Clear reproduction steps

Remove credentials, tokens, customer names, infrastructure addresses, and other sensitive information before sharing evidence.

