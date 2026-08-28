Import-module ..\ALM\scripts\_functions.ps1 -Force
Import-module ./CreateTeamsFunctions.ps1 -Force
# Se importa el módulo de funciones
function createTabsManually {
    param(
        [Parameter(Mandatory = $true)]
        [string]$TeamsId,
        [Parameter(Mandatory = $true)]
        [string]$ChannelId,
        [Parameter(Mandatory = $false)]
        [string]$SiteName
    )

    try {

        # Se verifica que el módulo "PnP.PowerShell" está instalado
        Ensure-PowerShellModule -ModuleName "PnP.PowerShell" -Version "2.12.0"

        # Se realiza la conexión a la administración de SharePoint Online del tenant correspondiente
        $Connection = Get-ALMAwarePnPConnectionToAdminSite  
        $Environment = $env:ALM_TENANTADMINURL -replace "-admin", ""

        $startTime = Get-Date
        $Timeout = 300
        $RetryInterval = 60
        $Success = $false;

        while ((((Get-Date) - $startTime).TotalSeconds -lt $Timeout) -and ($Success -eq $false)) {
            
            try {
                $ChannelUrl = Get-ChannelURL -TeamsId $TeamsId -ChannelId $ChannelId -GraphConnection $Connection
                if ($null -ne $ChannelUrl) {
                    #Save in property bag
                
                    Save-propertyBag -url $ChannelUrl -bodyName $SiteName
                    $Success = $true
                }
            }
            catch {
                Write-Host "Retrying... Error: $($_)"
                Start-Sleep -Seconds $RetryInterval
            }
            
        }
        if ($Success -eq $false) {
            throw "It has not been possible to obtain channel url in $Timeout seconds."
        }

        # Definir los parámetros comunes para las pestañas
        $TabParams = @{
            Team       = $TeamsId
            Channel    = $ChannelId
            Type       = "WebSite" 
            Connection = $Connection
        }

        # Añadir la segunda pestaña al canal
        $TabParams.DisplayName = "Sitio formal"
        $TabParams.ContentUrl = "$Environment/sites/$SiteName"


        $startTime = Get-Date
        $Success = $false;

        while ((((Get-Date) - $startTime).TotalSeconds -lt $Timeout) -and ($Success -eq $false)) {
            try {
            
                Add-PnPTeamsTab @TabParams
                $Success = $true
                Write-Host "Tab '$($TabParams.DisplayName)' for '$($TabParams.ContentUrl)' has been created successfully"
            }
            catch {
                Write-Host "Retrying... Error: $($_)"
                Start-Sleep -Seconds $RetryInterval
            }
            
        }

        if ($Success -eq $false) {
            Write-Host  "Tab '$($TabParams.DisplayName)' for '$($TabParams.ContentUrl)' has not been created"
        }

        # Añadir la segunda pestaña al canal
        $TabParams.DisplayName = "Envío de archivos"
        $TabParams.ContentUrl = $ChannelUrl

        $startTime = Get-Date
        $Success = $false;

        while ((((Get-Date) - $startTime).TotalSeconds -lt $Timeout) -and ($Success -eq $false)) {
            try {
           
                Add-PnPTeamsTab @TabParams
                $Success = $true
                Write-Host "Tab '$($TabParams.DisplayName)' for '$($TabParams.ContentUrl)' has been created successfully"
            }
            catch {
                Write-Host "Retrying... Error: $($_)"
                Start-Sleep -Seconds $RetryInterval
            }
        }
        if ($Success -eq $false) {
            Write-Host  "Tab '$($TabParams.DisplayName)' for '$($TabParams.ContentUrl)' has not been created"
        }
    }
    catch {
        write-host "TEAMS: Error: $_.Exception.Message"
    }
}

$params = @{
    TeamsId   = "2b5f7ff4-62ae-45ee-9bfc-b057bb0d920a"
    ChannelId = "19:68ed38af89ba46f7820d070991565473@thread.tacv2"
    SiteName  = "demomvp-cs-testColCinco"
} 

createTabsManually @params


# Tenant id = 805543fa-f86a-4a60-9ace-f53521807088
# spurl =  https://test.sharepoint.com 
# admin url = https://test-admin.sharepoint.com