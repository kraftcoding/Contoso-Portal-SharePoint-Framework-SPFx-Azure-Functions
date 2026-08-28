Import-module ..\ALM\scripts\_functions.ps1 -Force
# Se importa el módulo de funciones

function Create-TeamsForSite {
    # Parámetros del script
    param(
        [Parameter(Mandatory = $true)]
        [string]$Environment,
        [Parameter(Mandatory = $true)]
        [object]$Body,
        [Parameter(Mandatory = $true)]
        [object[]]$Bodies,
        [Parameter(Mandatory = $true)]
        [object]$Connection
      
    )

    try {
        $SiteName = $Body["Title"]
        $Channel = Truncate-String -str (Get-SafeName -name $Body["Department"].Label) -maxLength 49
        $SiteOwner = $Body["Editor"].Email
        $MembershipType = "Standard"


        $teamsCreated = Get-ConferenciaTeams -Body $Body -Bodies $Bodies -Connection $Connection

        if ($null -eq $teamsCreated) {
            throw "Error to create teams $($SiteName)"
        }

        $ChannelCreated = $null

        try {
            #Comprodbar si existe el canal
            $ChannelCreated = Get-PnPTeamsChannel -Team $teamsCreated.GroupId -Connection $Connection | Where-Object { $_.DisplayName -eq $Channel }
        }
        catch {
            Write-Host "Channel $($ChannelCreated) for $($SiteName) not found. Creating it... Exception: $_"
        }
        
        if ($null -ne $ChannelCreated) { 
            #Si existe no se puede crear con el mismo nombre
            throw "Channel $($ChannelCreated) for $($SiteName) already exists."
        }
        
        # Crear el canal 
        $ChannelCreated = Add-PnPTeamsChannel -Team $teamsCreated.GroupId -DisplayName $Channel -Description  $SiteName -ChannelType $MembershipType -Connection $Connection
        
        write-host 
        write-host "Se ha creado el siguiente canal: "
        write-host "  Nombre: $($Channel)"
        write-host "  prodpietario: $($SiteOwner)"
        write-host "  Department: $($SiteName)"
        write-host 

        # Tiempo para la creacion del canal cuando es privado y que sea recuperable desde graph 

        $startTime = Get-Date
        $Timeout = 1800
        $RetryInterval = 60
        $Success = $false;

        while ((((Get-Date) - $startTime).TotalSeconds -lt $Timeout) -and ($Success -eq $false)) {
            try {
                $ChannelUrl = Get-ChannelURL -TeamsId $teamsCreated.GroupId -ChannelId $ChannelCreated.Id -GraphConnection $Connection
                if ($null -ne $ChannelUrl) {
                    #Save in property bag
                    Save-propertyBag -url $ChannelUrl -bodyName $SiteName
                    $Success = $true
                }
            }
            catch {
              write-host "Exception: $_"
            }
            Start-Sleep -Seconds $RetryInterval
        }
        if ($Success -eq $false) {
            throw "It has not been possible to obtain channel url in $Timeout seconds."
        }

        # Definir los parámetros comunes para las pestañas
        $TabParams = @{
            Team       = $teamsCreated.GroupId
            Channel    = $ChannelCreated.Id
            #Type       = "WebSite" 
            Type        = "SharePointPageAndList"
            Connection = $Connection
        }

        # Añadir la primera pestaña al canal
        $TabParams.DisplayName = "Sitio formal"
        #$TabParams.ContentUrl = "$Environment/sites/$SiteName/sitepages/home.aspx"
        $TabParams.WebsiteUrl = "$Environment/sites/$SiteName/sitepages/home.aspx"
        $Timeout = 1200
        $startTime = Get-Date
        $Success = $false;

        Write-Host "Creating Tab '$($TabParams.DisplayName)' for '$($TabParams.WebsiteUrl)'"

        while ((((Get-Date) - $startTime).TotalSeconds -lt $Timeout) -and ($Success -eq $false)) {
            try {
                Add-PnPTeamsTab @TabParams
                $Success = $true
                Write-Host "Tab '$($TabParams.DisplayName)' for '$($TabParams.WebsiteUrl)' has been created successfully"
            }
            catch {
                Write-Host "Retrying... Error: $($_.Exception.Message)"
            }
            Start-Sleep -Seconds $RetryInterval
        }

        if ($Success -eq $false) {
            Write-Host  "Tab '$($TabParams.DisplayName)' for '$($TabParams.WebsiteUrl)' has not been created"
        }

        # Añadir la segunda pestaña al canal
        $TabParams.DisplayName = "Envío de archivos"
        #$TabParams.ContentUrl = $ChannelUrl
        $TabParams.WebsiteUrl = $ChannelUrl
        $startTime = Get-Date
        $Success = $false;

        Write-Host "Creating Tab '$($TabParams.DisplayName)' for '$($TabParams.WebsiteUrl)'"

        while ((((Get-Date) - $startTime).TotalSeconds -lt $Timeout) -and ($Success -eq $false)) {
            try {
                Add-PnPTeamsTab @TabParams
                $Success = $true
                Write-Host "Tab '$($TabParams.DisplayName)' for '$($TabParams.WebsiteUrl)' has been created successfully"
            }
            catch {
                Write-Host "Retrying... Error: $($_.Exception.Message)"
                
            }
            Start-Sleep -Seconds $RetryInterval
     
        }
        if ($Success -eq $false) {
            Write-Host  "Tab '$($TabParams.DisplayName)' for '$($TabParams.WebsiteUrl)' has not been created"
        }
    }
    catch {
        write-host "Error: $_"
    }
}

function Get-SafeName {
    param (
        [string]$name
    )

    $unsafeChars = @('"', '*', ':', '<', '>', '?', '/', '\', '|', '&', '#', '%', "'", '+', '=', ';', '[', ']', '@')
    $sb = New-Object -TypeName System.Text.StringBuilder

    foreach ($char in $name.ToCharArray()) {
        if ($unsafeChars -notcontains $char) {
            [void]$sb.Append($char)
        }
        else {
            [void]$sb.Append('_')
        }
    }

    return $sb.ToString().Trim()
}

function Truncate-String {
    param (
        [string]$str,
        [int]$maxLength
    )

    if ([string]::IsNullOrECNTy($str)) {
        return ""
    }

    return $str.Substring(0, [Math]::Min($str.Length, $maxLength))
}

function Ensure-TeamsForSite {
    # Parámetros del script
    param(
        [Parameter(Mandatory = $true)]
        [object]$Body,
        [Parameter(Mandatory = $true)]
        [object]$Connection
    )
    try {
        
        $SiteName = $Body['Title']
        $TeamsTitle = $Body["Department"].Label
        $SiteOwner = $Body["Editor"].Email
        $IdentificadorConferencia = $Body["IdentificadorConferencia"]

        if ($null -ne $IdentificadorConferencia) {
            try {
                #Comprodbar si existe el canal
                $teamsCreated = Get-PnPTeamsTeam -Identity $IdentificadorConferencia -Connection $Connection
                return $teamsCreated;
            }
            catch {
                Write-Host "Teams for $($TeamsName) not found. Creating it... Exception: $_.Exception.Message"
            }
        }
        #Parametros de la creacion del teams
        $teamCreateParams = @{

            DisplayName              = $TeamsTitle 
            Description              = $TeamsTitle 
            Visibility               = "Private"
                
            AllowUserEditMessages    = $true
            AllowUserDeleteMessages  = $true
            AllowOwnerDeleteMessages = $true
            AllowTeamMentions        = $true
            AllowChannelMentions     = $true
    
            AllowGiphy               = $true
            GiphyContentRating       = "Moderate"
            AllowStickersAndMemes    = $true
            AllowCustomMemes         = $true
    
            Owner                    = $SiteOwner
            Connection               = $Connection
        }
        
        # Intentar crear el equipo
        try {
            $teamsCreated = New-PnPTeamsTeam @teamCreateParams
        }
        catch {
            Write-Host "Error al crear el equipo con el correo. Reintentando con UPN..."

            # Obtener el UPN desde el campo Editor
            $editorId = $Body["Editor"].LookupId
            $user = Get-PnPUser -Identity $editorId -Connection $Connection
            $upn = $user.LoginName -replace '^.*\|', ''

            # Actualizar el Owner con el UPN
            $teamCreateParams["Owner"] = $upn
            $teamsCreated = New-PnPTeamsTeam @teamCreateParams
        }

        # Crear el teams con parametros
       # $teamsCreated = New-PnPTeamsTeam @teamCreateParams 

        Update-BodyItem -IdDepartment $Body.Id -Values @{"IdentificadorConferencia" = $teamsCreated.GroupId; "IdentificadorTeams"= $teamsCreated.GroupId  } -Connection $Connection
        
        return $teamsCreated;
    }
    catch {
        write-host "Error to ensure team $($TeamsTitle): $_.Exception.Message"
    }
    return $null
}


function Update-BodyItem {
    # Parámetros del script
    param(

        [Parameter(Mandatory = $true)]
        [string]$IdDepartment,
        [Parameter(Mandatory = $true)]
        [hashtable]$Values,
        [Parameter(Mandatory = $true)]
        [object]$Connection
    )
    try {
        $params = @{
            List       = "ConfiguracionDepartments"
            Identity   = $IdDepartment
            Values     = $Values
            Connection = $Connection
        }
        Set-PnPListItem @params
    }
    catch {
        write-host "Error to update body item $($IdDepartment): $_.Exception.Message"
    }
    
}

function Get-ConferenciaTeams {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Body,
        [Parameter(Mandatory = $true)]
        [object]$Connection,
        [Parameter(Mandatory = $false)]
        [object[]]$Bodies
    )

    $tipoDepartment = $Body['TipoDepartment']
    $Conferencia = $Body

    <# if ($tipoDepartment.TermGuid -ne "63ac1500-07b4-4e25-b4ff-6fbd946389ff") {
        $IdentificadorConferencia = $Body["IdentificadorConferencia"]
        $Conferencia = $Bodies | Where-Object { $_['Title'] -eq $IdentificadorConferencia }
        
        if ($null -eq $Conferencia) {
            throw "Identificador conferencia for $($Body['Title']) not found"
        }
    } #>
    return Ensure-TeamsForSite -Body $Conferencia -Connection $Connection
}

function Get-ChannelURL {
    param(
        [Parameter(Mandatory = $true)]
        [object]$TeamsId,
        [Parameter(Mandatory = $true)]
        [object]$GraphConnection,
        [Parameter(Mandatory = $false)]
        [object[]]$ChannelId
    )
    try {
        $res = Invoke-PnPGraphMethod "https://graph.microsoft.com/v1.0/teams/$TeamsId/channels/$ChannelId/filesFolder" -Connection $GraphConnection
        return $res.webUrl
    }
    catch {
        write-host
        write-host "------------------------ SISTEMA COLABORATIVO ------------------------"
        write-host 
        write-host "Es necesario realizar el paso manual del prodcedimiento de generacion de Departments."
        write-host 
        Write-Host "Exception: $($_)"
        write-host 
        write-host "----------------------------------------------------------------------"
        write-host
    }
    return $null
}

function Save-propertyBag {
    param(
        [Parameter(Mandatory = $true)]
        [string]$url,
        [Parameter(Mandatory = $true)]
        [string]$bodyName
    )
    try {
        $ChannelInfo = ExtractSharePointInfo -url $url

        $siteConnection = Get-ALMAwarePnPConnectionBySiteRelativeUrl -siteRelativeUrl "/sites/$($ChannelInfo.TeamsSite)"

        $folderRelative = "/sites/$($ChannelInfo.TeamsSite)/$($ChannelInfo.MainFolder)/$($ChannelInfo.SubFolder)"

        $query = "<View Scope='RecursiveAll'><Query><Where><Eq><FieldRef Name='FileRef'/><Value Type='Text'>$folderRelative</Value></Eq></Where></Query></View>"

        $folderItems = Get-PnPListItem -List $ChannelInfo.MainFolder -Query $query -Connection $siteConnection

        $key = "Bagprodp_$($folderItems.Id)"

        Set-PnPpropertyBagValue -Key $key -Value $bodyName -Folder $ChannelInfo.MainFolder -Connection $siteConnection

        write-host "Se ha almacenado la prodpiedad $key con el valor $bodyName en el sitio $($ChannelInfo.TeamsSite)"
    }
    catch {
        throw "No se ha podido almacenar la property bag de $url con respecto a $bodyName. Exception: $_"
    }
  


}


function ExtractSharePointInfo {
    param(
        [Parameter(Mandatory = $true)]
        [string]$url
    )
    # Extrae la URL del sitio
    $siteName = $url -replace "https://[^/]+/sites/([^/]+).*", '$1'
    Write-Host "Nombre del sitio: $siteName"

    # Extrae la carpeta y las subcarpetas
    $folderPath = $url -replace "https://[^/]+/sites/[^/]+", ''
    $folderPath = [System.Web.HttpUtility]::UrlDecode($folderPath)  # Decodifica la URL

    # Divide la carpeta y las subcarpetas en dos variables
    $folders = $folderPath.Split("/", [System.StringSplitOptions]::RemoveECNTyEntries)
    $mainFolder = $folders[0]
    $subFolders = $folderPath.Replace("/$mainFolder/", "")

    Write-Host "Carpeta principal: /$mainFolder"
    Write-Host "Subcarpetas: /$subFolders"

    return @{
        TeamsSite  = $siteName
        MainFolder = $mainFolder
        SubFolder  = $subFolders
    }
}
