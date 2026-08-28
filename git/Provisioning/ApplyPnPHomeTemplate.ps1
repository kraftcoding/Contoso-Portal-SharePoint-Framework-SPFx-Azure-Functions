# Se importa el módulo de funciones
Import-module ..\ALM\scripts\_functions.ps1 -Force

try {
    # Se verifica que el módulo "PnP.PowerShell" está instalado
    Ensure-PowerShellModule -ModuleName "PnP.PowerShell" -Version "2.12.0"

    # Se realiza la conexión a la administración de SharePoint Online del tenant correspondiente
    $connection = Get-ALMAwarePnPConnectionBySiteRelativeUrl "/sites/Contoso"

    $adminGroup = Get-PnPAzureADGroup -Identity "contoso-administration" -Connection $connection
    $supportGroup = Get-PnPAzureADGroup -Identity "contoso-support" -Connection $connection

    $ocpGroup = Get-PnPAzureADGroup -Identity "contoso-operations" -Connection $connection
    $params = @{}
    $params.Add("Contoso-Org_admins", "c:0t.c|tenant|$($adminGroup.Id)")
    $params.Add("Contoso-Org_support", "c:0t.c|tenant|$($supportGroup.Id)")
    $params.Add("Contoso-Org_Ocp", "c:0t.c|tenant|$($ocpGroup.Id)")

    # Se aplica la plantilla del tenant, limitándose a los términos de taxonomía
    # Read-PnPSiteTemplate -Path ./templateHomeCNTtest.xml
    Invoke-PnPSiteTemplate -Path ./templateHomeCNTtest.xml -Parameters $params -Connection $connection
}
catch {
    Write-Host "Error: $_"
}