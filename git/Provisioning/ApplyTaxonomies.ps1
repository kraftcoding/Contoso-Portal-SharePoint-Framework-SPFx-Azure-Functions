# Se importa el módulo de funciones
Import-module ..\ALM\scripts\_functions.ps1 -Force

#DEV

# Tenant id = contoso-test.onmicrosoft.com 
# spurl =  https://contoso-test.sharepoint.com 
# admin url = https://contoso-test-admin.sharepoint.com

#test

# Tenant id = test.onmicrosoft.com 
# spurl =  https://test.sharepoint.com 
# admin url = https://test-admin.sharepoint.com

# prod

# Tenant id = politicaterritorial.onmicrosoft.com 
# spurl =  https://politicaterritorial.sharepoint.com 
# admin url = https://politicaterritorial-admin.sharepoint.com

try {
    # Se verifica que el módulo "PnP.PowerShell" está instalado
    Ensure-PowerShellModule -ModuleName "PnP.PowerShell" -Version "2.12.0"

    # Se realiza la conexión a la administración de SharePoint Online del tenant correspondiente
    $connection = Get-ALMAwarePnPConnectionToAdminSite

    # Se aplica la plantilla del tenant, limitándose a los términos de taxonomía
    Invoke-PnPTenantTemplate -Path ./tenantTemplate.xml -Handlers TermGroups -Connection $connection

}
catch {
    Write-Host "Error: $_"
}