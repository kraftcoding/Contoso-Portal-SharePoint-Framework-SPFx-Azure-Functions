$artifactWebApiPath = "./dropwebapi.zip"
$artifMinutesuxPath = "./dropaux.zip"

# Environment: contoso-test Dev tenant environment
# $tenant = "3963ea86-5f68-44d4-bf00-010ae86e1f9b"
# $subscription = "f8e1e688-dd8c-4e54-9ce2-e58d7b832c3a"
# $rgName = "rg-devcontoso-test-Contoso"
# $appName = "funcdevcontoso-testContoso001"
# $appAuxName = "funcdevcontoso-testContoso001aux"

# Environment: test environment
$tenant = "805543fa-f86a-4a60-9ace-f53521807088"
$subscription = "19dd0758-e955-4244-868e-2f4e008dc58b"
$rgName = "rg-test-Contoso"
$appName = "functestContoso002"
$appAuxName = "functestContoso002aux"

& $PSScriptRoot/support-build-code.ps1

az login --tenant $tenant --use-device-code
az account set --name $subscription

$codeDeploymentResult = (az webapp deploy --resource-group $rgName --name $appName --src-path ($artifactWebApiPath) --type zip --async true | ConvertFrom-Json)
if (!$?) { write-error "Error during code deployment"; exit; }
Write-Host "Code deployment completed"
$codeDeploymentResult | Select-Object complete, site_name, start_time, end_time | format-list

$codeDeploymentAuxResult = (az webapp deploy --resource-group $rgName --name $appAuxName --src-path ($artifMinutesuxPath) --type zip --async true | ConvertFrom-Json)
if (!$?) { write-error "Error during code deployment (aux)"; exit; }
Write-Host "Code deployment completed (aux)"
$codeDeploymentAuxResult | Select-Object complete, site_name, start_time, end_time | format-list
