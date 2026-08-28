
# Environment: contoso-test Dev tenant environment
$tenant = "3963ea86-5f68-44d4-bf00-010ae86e1f9b"
$envParameters = "devcontoso-test"

# # Environment: test environment
# $tenant = "805543fa-f86a-4a60-9ace-f53521807088"
# $envParameters = "test"

# # Environment: prod environment
# $tenant = "abb2d05d-20f8-4a6d-a887-51b883731814"
# $envParameters = "prod"

az login --tenant $tenant --use-device-code --allow-no-subscriptions

# Update web api app registration

# add certificate
$webApiAppId = (az ad app show  --id api://Contoso | ConvertFrom-Json).appId
az ad app credential reset --id $webApiAppId --cert=@webapi.cer --append
## add swagger endpoint
$objectId = (az ad app show --id $webApiAppId | ConvertFrom-Json).id
az rest --method PATCH --uri "https://graph.microsoft.com/v1.0/applications/$objectId" --body=@aptestg-webapi-update.$envParameters.json

# Update DevOps app registration

$devopsAppId = (az ad app show  --id api://Contosodevops | ConvertFrom-Json).appId
az ad app credential reset --id $devopsAppId --cert=@devops.cer --append
