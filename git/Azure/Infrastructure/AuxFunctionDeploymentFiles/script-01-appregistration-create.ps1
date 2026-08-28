# App name must be unique for SPFx resources configuration
$webApiAppName = "service-account@contoso.local Web API"
$devOpsAppName = "service-account@contoso.local DevOps"

# Environment: contoso-test Dev tenant environment
$tenant = "3963ea86-5f68-44d4-bf00-010ae86e1f9b"

# Environment: test environment
# $tenant = "805543fa-f86a-4a60-9ace-f53521807088"

# # ## Environment: prod environment
# $tenant = "abb2d05d-20f8-4a6d-a887-51b883731814"

az login --tenant $tenant --use-device-code --allow-no-subscriptions

# Create the web API app
$webApiAppId = (az ad app create --display-name $webApiAppName --sign-in-audience AzureADMyOrg | ConvertFrom-Json).appId
$webApiObjectId = (az ad app show --id $webApiAppId | ConvertFrom-Json).id
sleep 10 # should wait for the app to be updated
# IMPORTANT: manually create Service Principal from azure portal app registration home page. 
$webApiSpId = (az ad sp show --id $webApiAppId | ConvertFrom-Json).id

write-host "Updating... - webApiAppId: $webApiAppId - webApiObjectId: $webApiObjectId - webApiSpId: $webApiSpId"

az rest --method PATCH --uri "https://graph.microsoft.com/v1.0/applications/$webApiObjectId" --body=@aptestg-webapi-base.json
sleep 25 # should wait for the app to be updated
az ad app permission admin-consent --id $webApiAppId

# Create the DevOps app
$devopsAppId = (az ad app create --display-name $devOpsAppName --sign-in-audience AzureADMyOrg | ConvertFrom-Json).appId
$devopsObjectId = (az ad app show --id $devopsAppId | ConvertFrom-Json).id
sleep 10 # should wait for the app to be updated

# IMPORTANT: manually create Service Principal from azure portal app registration home page. 
$devopsSpId = (az ad sp show --id $devopsAppId | ConvertFrom-Json).id

write-host "Updating... - devopsAppId: $devopsAppId - devopsObjectId: $devopsObjectId - devopsSpId: $devopsSpId"

az rest --method PATCH --uri "https://graph.microsoft.com/v1.0/applications/$devopsObjectId" --body=@aptestg-devops-base.json
sleep 25 # should wait for the app to be updated
az ad app permission admin-consent --id $devopsAppId

Write-Host "IMPORTANT: Update environment specific parameters with: param_devopsAppObjectId: $devopsSpId and param_aadClientId: $webApiAppId"