$root = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }

$projectPath = Join-Path $root "../../../WebApi/"
$outputFolder = Join-Path $root "debug_output"
$csprojPath = Join-Path $projectPath "WebApi.csproj"

$artifactName = "dropwebapi.zip"
$artifactNameAux = "dropaux.zip"

$artifactPath = Join-Path $root $artifactName
$artifMinutesuxPath = Join-Path $root $artifactNameAux

# Limpieza testvia
if (Test-Path $outputFolder) {
    Remove-Item $outputFolder -Recurse -Force
}

# Build WebApi
# Write-Host "Publicando WebApi..."
# dotnet clean $csprojPath -c Debug 
# dotnet publish $csprojPath -c Debug -o $outputFolder --no-self-contained 
# dotnet publish $csprojPath -c Debug -o $outputFolder --no-self-contained /p:DebugType=portable /p:DebugSymbols=true

# Copiar carpeta certs a output
# Copy-Item -Path ".\certs" -Destination "$outputFolder\certs" -Recurse -Force

# El ZIP debe contener SOLO el contenido publicado
# Write-Host "Generando ZIP WebApi..."
# if (Test-Path $artifactPath) { Remove-Item $artifactPath -Force }
# Comtestss-Archive -Path "$outputFolder\*" -DestinationPath $artifactPath -Force

# Build Aux Function
Write-Host "Generando ZIP Aux..."
$auxFolder = Join-Path $root "pwshauxfunction"
if (Test-Path $artifMinutesuxPath) { Remove-Item $artifMinutesuxPath -Force }
Comtestss-Archive -Path "$auxFolder\*" -DestinationPath $artifMinutesuxPath -Force

Write-Host "Artefactos generados:"
# Write-Host " - $artifactPath"
Write-Host " - $artifMinutesuxPath"