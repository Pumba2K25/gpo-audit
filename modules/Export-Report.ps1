# Export-Report.ps1
# Exports GPO audit data to CSV and HTML

function Export-GPOToCSV {
    param(
        [array]$GPOData,
        [string]$Path
    )

    $GPOData | Select-Object Name, ID, Status, Owner, Created, LastModified,
        ComputerEnabled, UserEnabled, LinkCount, IsUnlinked, WMIFilter, LinkedOUs |
        Export-Csv -Path $Path -NoTypeInformation -Encoding UTF8
}

function Export-GPOToHTML {
    param(
        [array]$GPOData,
        [string]$Domain,
        [string]$Path
    )

    $totalGPOs    = $GPOData.Count
    $unlinked     = ($GPOData | Where-Object { $_.IsUnlinked }).Count
    $disabled     = ($GPOData | Where-Object { $_.Status -eq 'AllSettingsDisabled' }).Count
    $compDisabled = ($GPOData | Where-Object { $_.ComputerEnabled -eq $false }).Count
    $userDisabled = ($GPOData | Where-Object { $_.UserEnabled -eq $false }).Count
    $generated    = Get-Date -Format "yyyy-MM-dd HH:mm"

    $rows = foreach ($gpo in $GPOData | Sort-Object Name) {
        $unlinkedBadge = if ($gpo.IsUnlinked) { '<span class="badge warn">Unlinked</span>' } else { '' }
        $statusClass   = switch ($gpo.Status) {
            'AllSettingsDisabled'        { 'disabled' }
            'AllSettingsEnabled'         { 'enabled' }
            'UserSettingsDisabled'       { 'partial' }
            'ComputerSettingsDisabled'   { 'partial' }
            default                      { 'partial' }
        }
        $statusLabel = switch ($gpo.Status) {
            'AllSettingsDisabled'        { 'All Disabled' }
            'AllSettingsEnabled'         { 'Enabled' }
            'UserSettingsDisabled'       { 'User Disabled' }
            'ComputerSettingsDisabled'   { 'Computer Disabled' }
            default                      { $gpo.Status }
        }

        "<tr>
            <td>$($gpo.Name) $unlinkedBadge</td>
            <td><span class='status $statusClass'>$statusLabel</span></td>
            <td>$($gpo.Owner)</td>
            <td>$($gpo.LastModified.ToString('yyyy-MM-dd'))</td>
            <td style='text-align:center'>$($gpo.LinkCount)</td>
            <td>$($gpo.WMIFilter)</td>
            <td class='ou-cell'>$($gpo.LinkedOUs -replace '; ', '<br>')</td>
        </tr>"
    }

    $html = @"
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>GPO Audit — $Domain</title>
<style>
  * { box-sizing: border-box; margin: 0; padding: 0; }
  body { font-family: Segoe UI, Tahoma, Geneva, Verdana, sans-serif; background: #f0f2f5; color: #222; font-size: 15px; }
  header { background: #1a1a2e; color: #fff; padding: 22px 36px; }
  header h1 { font-size: 1.5rem; font-weight: 700; letter-spacing: 0.3px; }
  header p  { font-size: 0.9rem; opacity: 0.75; margin-top: 5px; }
  .summary  { display: flex; gap: 16px; padding: 28px 36px 12px; flex-wrap: wrap; }
  .card     { background: #fff; border-radius: 10px; padding: 18px 26px; min-width: 150px;
              box-shadow: 0 2px 6px rgba(0,0,0,.08); }
  .card .num { font-size: 2.2rem; font-weight: 700; color: #1a1a2e; line-height: 1; }
  .card .lbl { font-size: 0.82rem; color: #555; margin-top: 5px; text-transform: uppercase; letter-spacing: 0.4px; }
  .card.warn .num { color: #c97a00; }
  .card.bad  .num { color: #c0392b; }
  .table-wrap { padding: 20px 36px 40px; overflow-x: auto; }
  table { width: 100%; border-collapse: collapse; background: #fff;
          border-radius: 10px; overflow: hidden; box-shadow: 0 2px 6px rgba(0,0,0,.08); }
  th { background: #1a1a2e; color: #fff; text-align: left; padding: 12px 16px; font-size: 0.85rem; font-weight: 600; white-space: nowrap; }
  td { padding: 11px 16px; font-size: 0.9rem; border-bottom: 1px solid #e8eaed; vertical-align: top; }
  tr:nth-child(even) td { background: #f8f9fb; }
  tr:hover td { background: #eef2ff !important; }
  tr:last-child td { border-bottom: none; }
  .status { display: inline-block; padding: 3px 10px; border-radius: 5px; font-size: 0.8rem; font-weight: 600; }
  .status.enabled  { background: #e6f9ee; color: #1d7a3f; }
  .status.disabled { background: #fdecea; color: #b71c1c; }
  .status.partial  { background: #fff4e5; color: #a05c00; }
  .badge.warn { background: #fff3cd; color: #7a5700; font-size: 0.75rem;
                padding: 2px 8px; border-radius: 5px; margin-left: 7px; font-weight: 600; }
  .ou-cell { font-size: 0.82rem; color: #444; line-height: 1.6; }
  footer { text-align: center; padding: 18px; font-size: 0.8rem; color: #aaa; }
</style>
</head>
<body>
<header>
  <h1>GPO Audit Report</h1>
  <p>Domain: $Domain &nbsp;|&nbsp; Generated: $generated</p>
</header>

<div class="summary">
  <div class="card"><div class="num">$totalGPOs</div><div class="lbl">Total GPOs</div></div>
  <div class="card warn"><div class="num">$unlinked</div><div class="lbl">Unlinked GPOs</div></div>
  <div class="card bad"><div class="num">$disabled</div><div class="lbl">Fully Disabled</div></div>
  <div class="card"><div class="num">$compDisabled</div><div class="lbl">Computer Side Disabled</div></div>
  <div class="card"><div class="num">$userDisabled</div><div class="lbl">User Side Disabled</div></div>
</div>

<div class="table-wrap">
  <table>
    <thead>
      <tr>
        <th>GPO Name</th>
        <th>Status</th>
        <th>Owner</th>
        <th>Last Modified</th>
        <th>Links</th>
        <th>WMI Filter</th>
        <th>Linked OUs</th>
      </tr>
    </thead>
    <tbody>
      $($rows -join "`n")
    </tbody>
  </table>
</div>

<footer>GPO Audit Tool — $generated</footer>
</body>
</html>
"@

    $html | Out-File -FilePath $Path -Encoding UTF8
}
