Import-module .\_functions.ps1 -Force

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
    write-host "prodvisioning taxonomy ..."
        
    $connection = Get-ALMAwarePnPConnectionToAdminSite
    Invoke-PnPTenantTemplate -path ./../../prodvisioning/tenantTemplate.xml -Connection $connection
        
    write-host "Taxonomy prodvisioned"
}
catch {
    write-host "Error prodvisioning taxonomy"
    Write-Host $_.Exception.Message
}