
# Load PnP module
Import-Module PnP.PowerShell -Force

# Auth details (placeholders)
$clientId      = "302efa8b-0b1e-49a4-bb84-324437c9bef9"
$tenantId      = "a1fe4560-eb54-46b7-bafc-4e94dd69cb03"
$certPath      = "/Users/pavaoleonardo/Desktop/SharePointAudit/PnPApp.pfx"
$certPassword  = "Alin@:1982!"

# Input CSV and output path
$csvPath       = "/Users/pavaoleonardo/Desktop/SharePointAudit/SiteUrl.csv"
$exportPath    = "/Users/pavaoleonardo/Desktop/SharePointAudit/sharepoint-group-audit.csv"
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

# $clientId      = "302efa8b-0b1e-49a4-bb84-324437c9bef9"
# $tenantId      = "a1fe4560-eb54-46b7-bafc-4e94dd69cb03"
# $certPath      = "/Users/pavaoleonardo/Desktop/SharePointAudit/PnPApp.pfx"
# $certPassword  = "Alin@:1982!"