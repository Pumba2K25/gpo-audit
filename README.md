# gpo-audit

A PowerShell-based GPO documentation and audit tool for Windows/Active Directory environments. Pulls Group Policy Object data from the domain and generates clean, readable reports.

---

## What It Does

- Enumerates all GPOs in the domain
- Reports GPO status, owner, creation/modification dates
- Identifies unlinked and fully disabled GPOs
- Maps each GPO to its linked OUs
- Flags WMI filters
- Exports results as an HTML report and/or CSV

---

## Requirements

- Windows machine joined to the domain
- PowerShell 5.1+
- RSAT installed with the following modules:
  - `GroupPolicy`
  - `ActiveDirectory`
- Read access to Group Policy in AD

---

## Usage

```powershell
# Run from the repo root — generates both HTML and CSV by default
.\GPO-Audit.ps1

# HTML report only
.\GPO-Audit.ps1 -HTMLReport

# CSV only
.\GPO-Audit.ps1 -CSVReport

# Target a specific domain
.\GPO-Audit.ps1 -Domain "contoso.local"

# Custom output path
.\GPO-Audit.ps1 -OutputPath "C:\Reports"
```

Reports are saved to the `output\` folder with a timestamp in the filename.

---

## Output

### HTML Report
- Summary cards — total GPOs, unlinked, disabled counts
- Full sortable table with GPO name, status, owner, last modified, link count, WMI filter, and linked OUs
- Unlinked GPOs flagged with a badge
- Color-coded status indicators

### CSV
Flat export of all GPO metadata — useful for importing into Excel or ticketing systems.

---

## File Structure

```
gpo-audit/
├── GPO-Audit.ps1            # Entry point
├── modules/
│   ├── Get-GPOData.ps1      # AD data collection
│   └── Export-Report.ps1    # HTML + CSV report generation
├── output/                  # Reports land here (gitignored)
└── README.md
```

---

## Notes

- Large domains with many GPOs may take several minutes to run — a progress bar is shown during collection
- Run from a machine with domain connectivity and appropriate AD read permissions
- Tested on Windows 10/11 with RSAT against Server 2016/2019/2022 domains
