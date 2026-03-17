# Get-GPOData.ps1
# Pulls all GPO metadata, links, and settings from Active Directory

function Get-GPOAuditData {
    [CmdletBinding()]
    param(
        [string]$Domain
    )

    $results = @()
    $allGPOs = Get-GPO -All -Domain $Domain

    $totalGPOs = $allGPOs.Count
    $i = 0

    foreach ($gpo in $allGPOs) {
        $i++
        Write-Progress -Activity "Reading GPOs" -Status "$i of $totalGPOs : $($gpo.DisplayName)" -PercentComplete (($i / $totalGPOs) * 100)

        # Get OU links via Get-ADOrganizationalUnit + GPO report
        $links = @()
        try {
            $xmlReport = [xml](Get-GPOReport -Guid $gpo.Id -ReportType XML -Domain $Domain)
            $linksNodes = $xmlReport.GPO.LinksTo
            if ($linksNodes) {
                foreach ($link in $linksNodes) {
                    $links += [PSCustomObject]@{
                        SOM        = $link.SOMPath
                        Enabled    = $link.Enabled
                        NoOverride = $link.NoOverride
                    }
                }
            }
        } catch {
            Write-Warning "Could not get report for GPO: $($gpo.DisplayName) — $_"
        }

        $results += [PSCustomObject]@{
            Name             = $gpo.DisplayName
            ID               = $gpo.Id
            Status           = $gpo.GpoStatus
            Owner            = $gpo.Owner
            Created          = $gpo.CreationTime
            LastModified     = $gpo.ModificationTime
            ComputerEnabled  = ($gpo.Computer.Enabled)
            UserEnabled      = ($gpo.User.Enabled)
            LinkedOUs        = if ($links.Count -gt 0) { ($links | ForEach-Object { $_.SOM }) -join "; " } else { "Not Linked" }
            LinkCount        = $links.Count
            IsUnlinked       = ($links.Count -eq 0)
            WMIFilter        = if ($gpo.WmiFilter) { $gpo.WmiFilter.Name } else { "None" }
        }
    }

    Write-Progress -Activity "Reading GPOs" -Completed
    return $results
}
