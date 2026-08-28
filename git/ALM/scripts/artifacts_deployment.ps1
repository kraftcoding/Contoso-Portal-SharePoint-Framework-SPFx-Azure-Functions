Import-module .\_functions.ps1 -Force

write-host "Current directory: $pwd"

write-host "Connecting to SharePoint Online..."
$conn = Get-ALMAwarePnPConnectionToAdminSite

write-host "Updating package..."
Add-PnPApp -Path "./Contoso.sppkg" -Connection $conn -Publish -SkipFeatureDeployment -Overwrite
write-host "SPFx updated."

write-host "Connecting to Azure..."
Get-ALMAwareAzConnection 

write-host "Updating WebApi..."
Publish-AzWebapp -ResourceGroupName ($env:ALM_RESOURCEGROUP) -Name ($env:ALM_AZFUNCNAME) -ArchivePath "./WebApi.zip" -Force
write-host "WebApi updated."