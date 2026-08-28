#Connect-SPOService -Url https://contoso-admin.sharepoint.com -ClientId "00000000-0000-0000-0001-000000000001"   -Tenant "contoso.onmicrosoft.com"   -Thumbprint "AAAA1111BBBB2222CCCC3333DDDD4444EEEE5555"

Import-Module PnP.PowerShell

$AdminCenterURL = "https://contoso-admin.sharepoint.com"

Connect-PnPOnline `
  -Url $AdminCenterURL `
  -ClientId "00000000-0000-0000-0001-000000000001" `
  -Tenant "contoso.onmicrosoft.com" `
  -Thumbprint "AAAA1111BBBB2222CCCC3333DDDD4444EEEE5555"

$sites = Get-PnPTenantSite -Detailed | Where-Object { $_.Template -ne "RedirectSite#0" }

foreach ($site in $sites) {
    Write-Host "Reindexando: $($site.Url)" -ForegroundColor Cyan

    Connect-PnPOnline `
      -Url $site.Url `
      -ClientId "00000000-0000-0000-0001-000000000001" `
      -Tenant "contoso.onmicrosoft.com" `
      -Thumbprint "AAAA1111BBBB2222CCCC3333DDDD4444EEEE5555"

    Request-PntestIndexWeb
}