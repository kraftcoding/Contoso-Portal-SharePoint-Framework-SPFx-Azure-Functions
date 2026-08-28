# Rutas absolutas para evitar prodblemas
$artifactWebApiPath = Join-Path $PSScriptRoot "dropwebapi.zip"
$artifMinutesuxPath    = Join-Path $PSScriptRoot "dropaux.zip"

# Environment: DES environment
$tenant = "00000000-0000-0000-0002-000000000002"
$subscription = "e1fc01b9-d35b-4f9a-9b36-c3010c89b363"
$rgName = "rg-dev-contoso"
$appName = "funcdevCNTmContoso01"
$appAuxName = "funcdevCNTmContoso01aux"

# Build local
& $PSScriptRoot/support-build-code-debug-v2.ps1

# Validación de artefactos
#if (!(Test-Path $artifactWebApiPath)) { Write-Error "No existe dropwebapi.zip"; exit }
if (!(Test-Path $artifMinutesuxPath))    { Write-Error "No existe dropaux.zip"; exit }

# Login
az login --tenant $tenant --use-device-code
az account set --name $subscription

 #Write-Host "Desplegando Web API..."
 #$codeDeploymentResult = az webapp deploy `
 #   --resource-group $rgName `
 #    --name $appName `
 #    --src-path $artifactWebApiPath `
 #    --type zip `
 #    --restart true `
 #   --async false | ConvertFrom-Json

Write-Host "Desplegando Aux..."
$codeDeploymentAuxResult = az webapp deploy `
    --resource-group $rgName `
    --name $appAuxName `
    --src-path $artifMinutesuxPath `
    --type zip `
    --restart true `
    --async false | ConvertFrom-Json

Write-Host "Despliegue completado"