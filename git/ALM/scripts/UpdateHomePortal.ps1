Import-module .\_functions.ps1 -Force

# for local development will be ignored if the script is executed in devops
$global:DEV_SHPHOSTURL = "https://test.sharepoint.com";
$homeContoso= "/sites/Contoso"
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

try {
    write-host "prodvisioning home service-account@contoso.local $($homeContoso) ..."
    
    $connection = Get-ALMAwarePnPConnectionBySiteRelativeUrl -siteRelativeUrl $homeContoso
    Invoke-PnPSiteTemplate -path ./../../prodvisioning/templateHomeCNTtest.xml -Connection $connection
    
    write-host "Site $($homeContoso) prodvisioned"
}
catch {
    write-host "Error prodvisioning site $($homeContoso)"
    Write-Host $_.Exception.Message
}