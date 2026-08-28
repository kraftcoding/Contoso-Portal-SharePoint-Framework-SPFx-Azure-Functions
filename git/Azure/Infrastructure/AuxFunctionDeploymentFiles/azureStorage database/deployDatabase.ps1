# Environment: contoso-test Dev tenant environment
$tenant = "3963ea86-5f68-44d4-bf00-010ae86e1f9b"
$subscription = "f8e1e688-dd8c-4e54-9ce2-e58d7b832c3a"
$rgName = "rg-devcontoso-test-Contoso"
$rgLocation = "westeurope" # only if new
$envParameters = "devcontoso-test"
$lowcost = $true


## Environment: test environment
# $tenant = "805543fa-f86a-4a60-9ace-f53521807088"
# $subscription = "19dd0758-e955-4244-868e-2f4e008dc58b"
# $rgName = "rg-test-Contoso"
# $rgLocation = "westeurope" # only if new
# $envParameters = "test"
# $lowcost = $false
# $azAppServiceSP = "5d30096c-a992-4466-906e-2af6f676704e"

# # Environment: prod environment
# $tenant = "abb2d05d-20f8-4a6d-a887-51b883731814"
# $subscription = "caa35b8b-a47a-4519-96dc-cc377ad6ab8f"
# $rgName = "rg-prod-Contoso"
# $rgLocation = "westeurope" # only if new
# $envParameters = "prod"
# $lowcost = $false
# $azAppServiceSP = "3b689652-386b-4810-a4f0-030dc799cf64"

#& $PSScriptRoot/support-build-code.ps1

az login --tenant $tenant --use-device-code
az account set --name $subscription

# create resource group if not exists
$rg = az group show -n $rgName
if (!$?) {
    #az group create --location $rgLocation --name $rgName; 
    if (!$?) { write-error "Not found rg $rgName"; } 
}

# get service principal for Azure App Service to configure keyvault access
#$azAppServiceSP = ($azAppServiceSP) ? $azAppServiceSP : (az ad sp show --id "abfa0a7c-a6b6-4736-8310-5855508787cd" --query id -o tsv);
#if (!$? -or !$azAppServiceSP) { write-error "Missing Resource provider service principal for Azure App Service o AD permissions to read SP's, register it using 'Register-AzResourceprovider -providerNamespace Microsoft.Web'"; exit; }

# get password for private certificate export
#$pwdForPfxCert = Read-Host "Enter certificate password (only if download is active)" -AsSecureString
#$pwdForPfxCertText = $pwdForPfxCert | ConvertFrom-SecureString -AsPlainText

# deploy arm template
$templatePath = "Contoso-azureStorage-template.json"
$templateParmsPath = "Contoso-azureStorage.parameters.$envParameters.json"

#$armDeploymentResult = New-AzResourceGroupDeployment -ResourceGroupName $rgName -TemplateFile $templatePath -TemplateParameterFile $templateParmsPath

$armDeploymentResult = (az deployment group create --name Contoso-template --resource-group $rgName --template-file $templatePath --parameters $templateParmsPath | ConvertFrom-Json)
#$armDeploymentResult = (az deployment group create --name Contoso-template --resource-group $rgName --template-file $templatePath --parameters $templateParmsPath | ConvertFrom-Json)
if (!$?) { write-error "Error during arm deployment"; exit; }
Write-Host "Arm deployment completed"
$armDeploymentResult.properties | Select-Object prodvisioningState, timestamp, error | format-list

sleep 5 # if done immediately will throw "Operation returned an invalid status code 'BadRequest'"
Write-Host "Finished"
