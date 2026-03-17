# GPO-Audit.ps1
# Entry point for the GPO documentation and audit tool
# Requires: RSAT (GroupPolicy module), ActiveDirectory module, Domain connectivity

#Requires -Modules GroupPolicy, ActiveDirectory

[CmdletBinding()]
param(
    [string]$Domain = $env:USERDNSDOMAIN,
    [string]$OutputPath = "$PSScriptRoot\output",
    [switch]$HTMLReport,
    [switch]$CSVReport
)

# Default to both reports if neither specified
if (-not $HTMLReport -and -not $CSVReport) {
    $HTMLReport = $true
    $CSVReport  = $true
}

# Ensure output folder exists
if (-not (Test-Path $OutputPath)) {
    New-Item -ItemType Directory -Path $OutputPath | Out-Null
}

# Load modules
. "$PSScriptRoot\modules\Get-GPOData.ps1"
. "$PSScriptRoot\modules\Export-Report.ps1"

Write-Host "`nGPO Audit Tool" -ForegroundColor Cyan
Write-Host "Domain : $Domain" -ForegroundColor Cyan
Write-Host "Output : $OutputPath`n" -ForegroundColor Cyan

# Pull GPO data
Write-Host "Collecting GPO data..." -ForegroundColor Yellow
$GPOData = Get-GPOAuditData -Domain $Domain

Write-Host "Found $($GPOData.Count) GPOs.`n" -ForegroundColor Green

# Export
if ($CSVReport) {
    $csvPath = Join-Path $OutputPath "GPO-Audit_$(Get-Date -Format 'yyyyMMdd_HHmm').csv"
    Export-GPOToCSV -GPOData $GPOData -Path $csvPath
    Write-Host "CSV  : $csvPath" -ForegroundColor Green
}

if ($HTMLReport) {
    $htmlPath = Join-Path $OutputPath "GPO-Audit_$(Get-Date -Format 'yyyyMMdd_HHmm').html"
    Export-GPOToHTML -GPOData $GPOData -Domain $Domain -Path $htmlPath
    Write-Host "HTML : $htmlPath" -ForegroundColor Green
}

Write-Host "`nDone." -ForegroundColor Cyan
