
# Load PnP module
Import-Module PnP.PowerShell -Force



# Input CSV and output path
$csvPath       = "xxxxxxxxx/SiteUrl.csv"
$exportPath    = "xxxxxxxxxx/sharepoint-group-audit.csv"
$results       = @()

# Read site URLs
$sites = Import-Csv -Path $csvPath

foreach ($site in $sites) {
    $siteUrl = $site.SiteUrl
    Write-Host "🔍 Auditing $siteUrl..." -ForegroundColor Cyan

    try {
        # Connect to SharePoint site
        Connect-PnPOnline -Url $siteUrl `
            -ClientId $clientId `
            -Tenant $tenantId `
            -CertificatePath $certPath `
            -CertificatePassword (ConvertTo-SecureString $certPassword -AsPlainText -Force)

        # Get all SharePoint groups
        $allGroups = Get-PnPGroup

        foreach ($group in $allGroups) {
            $groupName = $group.Title
            try {
                $groupMembers = Get-PnPGroupMember -Identity $groupName | Select-Object -ExpandProperty LoginName
                if ($groupMembers.Count -gt 0) {
                    foreach ($user in $groupMembers) {
                        $results += [PSCustomObject]@{
                            SiteUrl  = $siteUrl
                            Role     = $groupName
                            UserName = $user
                        }
                    }
                }
                else {
                    $results += [PSCustomObject]@{
                        SiteUrl  = $siteUrl
                        Role     = $groupName
                        UserName = "<No members>"
                    }
                }
            }
            catch {
                $results += [PSCustomObject]@{
                    SiteUrl  = $siteUrl
                    Role     = $groupName
                    UserName = "Error: $(${_.Exception.Message})"
                }
            }
        }

        Write-Host "✔️ Success for ${siteUrl}" -ForegroundColor Green
    }
    catch {
        Write-Host "❌ Error processing ${siteUrl}: $(${_.Exception.Message})" -ForegroundColor Red
        $results += [PSCustomObject]@{
            SiteUrl  = $siteUrl
            Role     = "Error"
            UserName = $(${_.Exception.Message})
        }
    }
}

# Export results to CSV
$results | Sort-Object SiteUrl, Role, UserName | Export-Csv -Path $exportPath -NoTypeInformation -Encoding UTF8
Write-Host "`n✅ Audit complete. Results saved to $exportPath" -ForegroundColor Yellow

