# Se importa el módulo de funciones
Import-module ..\ALM\scripts\_functions.ps1 -Force

try {
    # Se verifica que el módulo "PnP.PowerShell" está instalado
    Ensure-PowerShellModule -ModuleName "PnP.PowerShell" -Version "2.12.0"

    # Se realiza la conexión a la administración de SharePoint Online del tenant correspondiente
    $connection = Get-ALMAwarePnPConnectionToAdminSite

    # Se aplica la plantilla del tenant, limitándose a los términos de taxonomía
    $hub = New-PnPSite -Type CommunicationSite -Title "Contoso" -Url "https://contoso-test.sharepoint.com/sites/Contoso" -Description "Contoso" -Owner "gcastrma_emeal.nttdata.com@contoso-test.onmicrosoft.com" -Lcid "3082" -Connection $connection | Out-Null

    if ($null -ne $hub.SiteUrl){
        Register-PnPHubSite -Site $hub.SiteUrl -Connection $connection
    }
}
catch {
    Write-Host "Error: $_"
}