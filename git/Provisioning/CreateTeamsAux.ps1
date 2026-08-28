Import-module ./CreateTeamsFunctions.ps1 -Force
$configSiteUrl = "/sites/Contoso"
$connection = Get-ALMAwarePnPConnectionBySiteRelativeUrl -siteRelativeUrl $configSiteUrl

$adminConnection = Get-ALMAwarePnPConnectionToAdminSite 

try{
    Set-PnPTraceLog -On -WriteToConsole -Level Debug
}catch {
    write-host "Launching on PnP.Powershell new version..."
    Write-Host $_.Exception.Message
    try{
        Start-PnPTraceLog -WriteToConsole -Level Debug
    }catch {
        write-host "Launching on PnP.Powershell OLD version..."
        Write-Host $_.Exception.Message
    }
}

$bodies = Get-PnPListItem -Connection $connection -List ConfiguracionDepartments -Query "<View><ViewFields><FieldRef Name='Title'/><FieldRef Name='TipoDepartment'/><FieldRef Name='Department'/><FieldRef Name='TipoMembresia'/><FieldRef Name='IdentificadorConferencia'/></ViewFields><Query><Where><BeginsWith><FieldRef Name='ContentTypeId' /><Value Type='ContentTypeId'>0x0120D520000639E2BEA58C8844A67C464647ED1640</Value></BeginsWith></Where></Query></View>"
  foreach ($body in $bodies) {
    try {
        $bodyTitle = $body['Title']
        $shpHostUrl = $env:ALM_TENANTADMINURL -replace "-admin", ""
        if($bodyTitle -eq "demomvp-cs-ae"){
            Create-TeamsForSite -Environment $shpHostUrl -Body $body -Bodies $bodies -Connection $connection
        }
    } catch {
        write-host "Error: $_"
    }
}
 
