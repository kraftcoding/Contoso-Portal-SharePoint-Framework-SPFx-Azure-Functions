# ============================
# CONFIGURACIÓN
# ============================
#DESARROLLO
#$site = "https://contoso.sharepoint.com/sites/DEMOTEST"
#$clientId = "00000000-0000-0000-0001-000000000001"
#test
#$site = "https://test.sharepoint.com/sites/DEMO"
#$clientId = "6bbaaccb-8f7a-4a20-9c8f-f29e45dd3382"

#$connection = Connect-PnPOnline -Url $site -ClientId $clientId -Interactive
#Write-Host "✅ Autenticación correcta"

Import-module ..\ALM\scripts\_functions.ps1 -Force

try {

    Ensure-PowerShellModule -ModuleName "PnP.PowerShell" -Version "2.12.0"

    # Se realiza la conexión a la administración de SharePoint Online del tenant correspondiente
    $connection = Get-ALMAwarePnPConnectionBySiteRelativeUrl "/sites/DEMO"

    $adminGroup = Get-PnPAzureADGroup -Identity "DEMO-administrators" -Connection $connection
    $gestoresGroup = Get-PnPAzureADGroup -Identity "DEMO-gestores" -Connection $connection
    $lectoresGroup = Get-PnPAzureADGroup -Identity "DEMO-lectores" -Connection $connection
    $params = @{}
    $params.Add("DEMO_admins", "c:0t.c|tenant|$($adminGroup.Id)")
    $params.Add("DEMO_gestores", "c:0t.c|tenant|$($gestoresGroup.Id)")
    $params.Add("DEMO_lectores", "c:0t.c|tenant|$($lectoresGroup.Id)")

    Invoke-PnPSiteTemplate -Path ./TemplateDEMO.xml -Parameters $params -Connection $connection
}
catch {
    Write-Host "Error: $_"
}