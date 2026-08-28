$root = if ($PSScriptRoot) { $PSScriptRoot }else { (Get-Location).Path }
$projectPath = join-path $root "/../WebApi/"
$outputFolder = join-path $root "/output"
$csprojPath = join-path $projectPath "WebApi.csproj"
$artifactName = "dropwebapi.zip"
$artifactPath = join-path $root $artifactName

# build webapi
cd $projectPath
dotnet clean --configuration Release /property:GenerateFullPaths=true /consoleloggerparameters:NoSummary
dotnet publish $csprojPath -o $outputFolder -c Release --os win --arch x64
cd $outputFolder
zip -r $artifactName ./
mv ./$artifactName ./../
cd ./../
rm $outputFolder -r --force

# zip auxfunction
$artifactNameAux = "dropaux.zip"
cd ./pwshauxfunction
zip -r $artifactNameAux ./
mv ./$artifactNameAux ./../
cd ./../