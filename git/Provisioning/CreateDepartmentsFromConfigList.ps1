Import-module ..\ALM\scripts\_functions.ps1 -Force
Import-module ./CreateTeamsFunctions.ps1 -Force

# for local development will be ignored if the script is executed in devops
#$env:ALM_TENANT = "contoso-dev.sharepoint.com";
#$env:ALM_TENANTADMINURL = "https://contoso-dev-admin.sharepoint.com"
# ALM_ENVIRONMENT -> devcdf, test, prod

# Custom Function to Check if Site Collection Exists in Given URL
function CheckSiteExists($SiteURL) {
    return (Get-PnPTenantSite -Url $SiteURL -Connection $adminConnection -ErrorAction SilentlyContinue) -ne $null    
}

# Function to get and create bodies added to config list
# works with DevOps and local development


function Get-UPNFromLogin {
    param([string]$login)
    if ([string]::IsNullOrWhiteSpace($login)) { return $null }

    # Casos típicos:
    # "i:0#.f|membership|user@contoso.com"
    # "i:0#.t|azuread|user@contoso.com"
    # Si hay un '@' al final después del último separador '|', lo tomamos como UPN
    if ($login -match "\|([^|]+@[^|]+)$") {
        return $matches[1]
    }

    # A veces viene como "user@contoso.com" sin claims
    if ($login -match "^[^|]+@[^|]+\.[^|]+$") {
        return $login
    }

    return $null
}

function Get-UserUPN {
    param(
        [Parameter(Mandatory=$true)] $user
    )
    # Prioridad: UserPrincipalName -> Email -> derivado de LoginName
    if ($user.PSObject.properties.Name -contains "UserPrincipalName" -and $user.UserPrincipalName) {
        return $user.UserPrincipalName
    }
    if ($user.Email) {
        return $user.Email
    }
    return Get-UPNFromLogin -login $user.LoginName
}


function Create-ALMAwarePnPBodiesFromConfigList {
    
    Write-Host "Getting bodies from config list:"
    
    $bodies = Get-PnPListItem -Connection $connection -List ConfiguracionDepartments -Query "<View><ViewFields><FieldRef Name='Title'/><FieldRef Name='TipoDepartment'/><FieldRef Name='Department'/><FieldRef Name='TipoMembresia'/><FieldRef Name='IdentificadorConferencia'/></ViewFields><Query><Where><BeginsWith><FieldRef Name='ContentTypeId' /><Value Type='ContentTypeId'>0x0120D520000639E2BEA58C8844A67C464647ED1640</Value></BeginsWith></Where></Query></View>"
    $shpHostUrl = $env:ALM_TENANTADMINURL -replace "-admin", ""

    foreach ($body in $bodies) {
        try {
            $bodyTitle = $body['Title']
            $siteRelativeUrl = "/sites/$bodyTitle"
            
            $siteUrl = New-Object System.Uri(([System.Uri]$shpHostUrl), $siteRelativeUrl)
            $siteOwner = $body["Editor"].Email
            $lcid = 3082
            if (CheckSiteExists($siteUrl)) {
                Write-Host "Site $siteUrl already created!"
            }
            else {
                if( $bodyTitle -match '^.+-(gr|co|cs)-.+$') {
                    Write-Host "Formato de titulo valido"
                } else {
                    throw "Formato inválido: en el título, el tipo de Department debe ser 'gr', 'co' o 'cs'."
                }
                $Department = $body['Department']
                $tipoDepartment = $body['TipoDepartment']
                #Intentar crear el equipo
                try {
                    $createdGroups = Create-ALMAwarePnPGroupsForSite -siteTitle $bodyTitle -siteOwner $siteOwner
                }
                catch {
                    try{
                        Write-Host "Error al crear los grupos con el correo. Reintentando con UPN..."

                        # Obtener el UPN desde el campo Editor
                        $editorId = $Body["Editor"].LookupId
                        $user = Get-PnPUser -Identity $editorId -Connection $Connection
                        $siteOwner = $user.LoginName -replace '^.*\|', ''

                        # Actualizar el Owner con el UPN
                        $createdGroups = Create-ALMAwarePnPGroupsForSite -siteTitle $bodyTitle -siteOwner $siteOwner
                    }catch{
                        try{
                            Write-Host "Error al crear grupos con upn. Reintentando con el author de Contoso..."
                            $ContosoWeb  = Get-PnPWeb -Includes Author -Connection $connection
                            if($ContosoWeb -and $ContosoWeb.Author){
                                if($ContosoWeb.Author.IsSiteAdmin -and -not($ContosoWeb.Author.IsShareByEmailGuestUser)){
                                    $upnAuthor = $ContosoWeb.Author.LoginName -replace '^.*\|', ''
                                    $createdGroups = Create-ALMAwarePnPGroupsForSite -siteTitle $bodyTitle -siteOwner $upnAuthor
                                }else{
                                    throw "grupos no creados con el author del sitio"
                                }
                            }else{
                                throw "grupos no creado con author del sitio"
                            }
                        }catch{
                            Write-Host "Error al crear grupos con el author de Contoso. Reintentando con los admin de la coleccion de sitios..."
                            $admins = Get-PnPUser -Connection $connection | Where-Object {
                                $_.IsSiteAdmin -eq $true -and
                                $_.IsShareByEmailGuestUser -eq $false -and
                                $_.LoginName -notmatch "rolemanager" -and
                                $_.LoginName -notmatch "spo-grid" -and
                                $_.Title -ne "SharePoint System" -and
                                $_.PrincipalType -eq "User"
                            }
                            $adminsSafe = $admins | Where-Object { $_ -ne $null }
                            # 1) Candidatos con "admin" en Email, Title o LoginName (insensible a mayúsculas)
                            $candidatosAdmin = $adminsSafe | Where-Object {
                                ($_.Email      -match "(?i)admin") -or
                                ($_.Title      -match "(?i)admin") -or
                                ($_.LoginName  -match "(?i)admin")
                            }

                            # 2) Elegimos el testferido:
                            #    - Si hay candidatos con "admin", tomamos el primero ordenado por Title, luego Email.
                            #    - Si no hay, tomamos el "siguiente" mejor: el primero ordenado alfabéticamente por Title, luego por Email.
                            $chosen = $null
                            if ($candidatosAdmin -and $candidatosAdmin.Count -gt 0) {
                                $chosen = $candidatosAdmin | Sort-Object Title, Email | Select-Object -First 1
                            } else {
                                $chosen = $adminsSafe        | Sort-Object Title, Email | Select-Object -First 1
                            }

                            if (-not $chosen) {
                                Write-Warning "No se encontró ningún administrador que cumpla los criterios."
                                throw "No se ha podido incluir ningun administrador en los grupos."
                            }
                            # 3) Obtener el UPN del usuario elegido
                            $chosenUPN = Get-UserUPN -user $chosen
                            $createdGroups = Create-ALMAwarePnPGroupsForSite -siteTitle $bodyTitle -siteOwner $chosenUPN
                        }
                    }

                }

                # BYPASS: si Exchange no estuvo disponible, createdGroups sera null o vacio
                if (-not $createdGroups) { $createdGroups = @{} }
                $ExchangeBypass = ($createdGroups.Count -eq 0)
                if ($ExchangeBypass) {
                    Write-Warning "Bypass Exchange activo para $($bodyTitle): grupos EXO omitidos. Continuando con creacion de sitio y Teams."
                }

                $adminGroup = Get-PnPAzureADGroup -Identity "contoso-administration" -Connection $connection
                
                $supportGroup = Get-PnPAzureADGroup -Identity "contoso-support" -Connection $connection
               
                $ocpGroup = Get-PnPAzureADGroup -Identity "contoso-operations" -Connection $connection
                
                $params = $createdGroups.Clone()
                
                $params.Add("Contoso-Org_admins", "c:0t.c|tenant|$($adminGroup.Id)")
                
                $params.Add("Contoso-Org_support", "c:0t.c|tenant|$($supportGroup.Id)")
                
                $params.Add("Contoso-Org_Ocp", "c:0t.c|tenant|$($ocpGroup.Id)")
                
                $params.Add("DefaultBodyNameTaxValue", "-1;#$($Department.Label)|$($Department.TermGuid)")
                
                $params.Add("DefaultBodyTypeTaxValue" , "-1;#$($tipoDepartment.Label)|$($tipoDepartment.TermGuid)")

                Write-Host "Creating site $siteUrl !" -ForegroundColor Yellow
                try {
                    New-PnPSite -Connection $connection -Type CommunicationSite -SiteDesign Blank -Title $Department.Label -Url $siteUrl -Description $Department.Label -Owner $siteOwner -Lcid $lcid -Wait | Out-Null
                    $createdSite = Get-PnPTenantSite -Url $siteUrl -Connection $Connection
                    if(-not $createdSite){
                        throw "Sitio no creado con Editor"
                    }
                    Write-Host "Created site $siteUrl !" -ForegroundColor Green
                }
                catch {
                    Write-Host "Error al crear el sitio con el usuario. Reintentando con administrators..."
                    try{
                        $ContosoWeb  = Get-PnPWeb -Includes Author -Connection $connection
                        if($ContosoWeb -and $ContosoWeb.Author){
                            if($ContosoWeb.Author.IsSiteAdmin -and -not($ContosoWeb.Author.IsShareByEmailGuestUser)){
                                $upnAuthor = $ContosoWeb.Author.LoginName -replace '^.*\|', ''
                                New-PnPSite -Connection $connection -Type CommunicationSite -SiteDesign Blank -Title $Department.Label -Url $siteUrl -Description $Department.Label -Owner $upnAuthor -Lcid $lcid -Wait| Out-Null
                                $createdSite = Get-PnPTenantSite -Url $siteUrl -Connection $Connection
                                if(-not $createdSite){
                                    throw "Sitio no creado con author del sitio"
                                }
                                Write-Host "Created site $siteUrl !" -ForegroundColor Green
                            }else{
                                throw "Sitio no creado con author del sitio"
                            }
                        }else{
                            throw "Sitio no creado con author del sitio"
                        }
                    }catch{
                        Write-Host "Error al crear el sitio con el author de Contoso. Reintentando con los admin de la coleccion de sitios..."

                        $admins = Get-PnPUser -Connection $connection | Where-Object {
                            $_.IsSiteAdmin -eq $true -and
                            $_.IsShareByEmailGuestUser -eq $false -and
                            $_.LoginName -notmatch "rolemanager" -and
                            $_.LoginName -notmatch "spo-grid" -and
                            $_.Title -ne "SharePoint System" -and
                            $_.PrincipalType -eq "User"
                        }
                        $adminsSafe = $admins | Where-Object { $_ -ne $null }
                        # 1) Candidatos con "admin" en Email, Title o LoginName (insensible a mayúsculas)
                        $candidatosAdmin = $adminsSafe | Where-Object {
                            ($_.Email      -match "(?i)admin") -or
                            ($_.Title      -match "(?i)admin") -or
                            ($_.LoginName  -match "(?i)admin")
                        }

                        # 2) Elegimos el testferido:
                        #    - Si hay candidatos con "admin", tomamos el primero ordenado por Title, luego Email.
                        #    - Si no hay, tomamos el "siguiente" mejor: el primero ordenado alfabéticamente por Title, luego por Email.
                        $chosen = $null
                        if ($candidatosAdmin -and $candidatosAdmin.Count -gt 0) {
                            $chosen = $candidatosAdmin | Sort-Object Title, Email | Select-Object -First 1
                        } else {
                            $chosen = $adminsSafe        | Sort-Object Title, Email | Select-Object -First 1
                        }

                        if (-not $chosen) {
                            Write-Warning "No se encontró ningún administrador que cumpla los criterios."
                            throw "No se ha podido incluir ningun administrador en el sitio."
                        }
                        # 3) Obtener el UPN del usuario elegido
                        $chosenUPN = Get-UserUPN -user $chosen
                        New-PnPSite -Connection $connection -Type CommunicationSite -SiteDesign Blank -Title $Department.Label -Url $siteUrl -Description $Department.Label -Owner $chosenUPN -Lcid $lcid -Wait| Out-Null
                        $createdSite = Get-PnPTenantSite -Url $siteUrl -Connection $Connection
                        if(-not $createdSite){
                            throw "Sitio no creado con ningun admin"
                            Write-Host "Sitio no creado con ningun admin..."

                        }
                        Write-Host "Created site $siteUrl !" -ForegroundColor Green
                    }
                    try{
                        $editorId = $Body["Editor"].LookupId
                        $user = Get-PnPUser -Identity $editorId -Connection $Connection
                        $siteOwner = $user.LoginName -replace '^.*\|', ''
                        Add-PnPSiteCollectionAdmin -Owners $siteOwner -Connection $connection
                    }catch{
                        Write-Warning "No se ha podido añadir al author como administrador de la coleccion"
                    }
                   
                }
                
                Set-PnPTenantSite -Url $siteUrl -SharingCapability ExistingExternalUserSharingOnly -Owners "c:0t.c|tenant|$($adminGroup.Id)"  -Connection $adminConnection
                
                Write-Host $params
                $siteConnection = Get-ALMAwarePnPConnectionBySiteRelativeUrl -siteRelativeUrl $siteRelativeUrl

                #necesario para la creacion de los grupos de seguridad
                #reposo para afianzar la creacion y reconocimiento de los grupos de seguridad del Department
                Start-Sleep -Duration (New-TimeSpan -Seconds 60)
                
                if (-not $ExchangeBypass) {
                    Add-PnPGroupMember -LoginName $createdGroups["Contoso-Org_members"] -Group 4 -Connection $connection
                    Add-PnPGroupMember -LoginName $createdGroups["Contoso-Org_gestorschedulers"] -Group 4 -Connection $connection
                    Add-PnPGroupMember -LoginName $createdGroups["Contoso-Org_schedulers"] -Group 4 -Connection $connection
                    Add-PnPGroupMember -LoginName $createdGroups["Contoso-Org_asistentemembers"] -Group 4 -Connection $connection
                }
                                
                Write-Host "Aplicando plantilla PnP..." -ForegroundColor Green

                try {
                   
                    Invoke-PnPSiteTemplate -path ./templateDepartmenttest.xml -Connection $siteConnection -Parameters $params
                }
                catch {
                    Write-Host "ERROR aplicando la plantilla PnP" -ForegroundColor Red

                    # Mensaje principal
                    Write-Host "Mensaje: $($_.Exception.Message)" -ForegroundColor Yellow

                    # Tipo de excepción
                    Write-Host "Tipo: $($_.Exception.GetType().FullName)" -ForegroundColor Yellow

                    # InnerException si existe
                    if ($_.Exception.InnerException) {
                        Write-Host "InnerException: $($_.Exception.InnerException.Message)" -ForegroundColor Yellow
                    }

                    # Detalles del servidor (cuando vienen de SharePoint)
                    if ($_.Exception.ServerErrorTypeName) {
                        Write-Host "ServerErrorTypeName: $($_.Exception.ServerErrorTypeName)" -ForegroundColor Magenta
                    }
                    if ($_.Exception.ServerErrorCode) {
                        Write-Host "ServerErrorCode: $($_.Exception.ServerErrorCode)" -ForegroundColor Magenta
                    }
                    if ($_.Exception.ServerErrorTraceCorrelationId) {
                        Write-Host "CorrelationId: $($_.Exception.ServerErrorTraceCorrelationId)" -ForegroundColor Magenta
                    }
                    if ($_.Exception.ServerErrorDetails) {
                        Write-Host "ServerErrorDetails: $($_.Exception.ServerErrorDetails)" -ForegroundColor Magenta
                    }

                    # Stack completo (muy útil)
                    Write-Host "StackTrace:" -ForegroundColor DarkYellow
                    Write-Host $_.Exception.StackTrace

                    # Re-lanzar si quieres que DevOps marque el job como fallido
                    # throw
                }
                
                # Set-PnPList -Connection $siteConnection -Identity MeetingsArchivadas -DefaultSensitivityLabelForLibrary "Reservado ampliado"
                # Set-PnPList -Connection $siteConnection -Identity MeetingsFinales -DefaultSensitivityLabelForLibrary "Reservado ampliado"
                
                #Create teams

                try {
                    
                    # Incidencia 1673198
                    # prodduce error porque el bloque que llama a Create-TeamsForSite se ejecuta inmediatamente después de crear el sitio SharePoint 
                    # Create-TeamsForSite -Environment $shpHostUrl -Body $body -Bodies $bodies -Connection $connection 
                    
                    # Se opta por sustituirlo por:
                    # Esperar hasta que el sitio esté completamente disponible antes de crear el equipo de Teams
                    $maxRetries = 10
                    $retryDelay = 30
                    $siteReady = $false

                    for ($i = 1; $i -le $maxRetries; $i++) {                       
                        try {
                            $check = Get-PnPTenantSite -Url $siteUrl -Connection $adminConnection -ErrorAction Stop
                            if ($check -ne $null) {
                                Write-Host "Sitio disponible en intento $i. prodcediendo con la creación del equipo de Teams..."
                                $siteReady = $true
                                break
                            }
                        } catch {
                            Write-Host "Sitio aún no disponible (intento $i de $maxRetries). Esperando $retryDelay segundos..."
                            Start-Sleep -Seconds $retryDelay
                        }
                    }
                    
                    if (-not $siteReady) {
                        Write-Warning "El sitio $siteUrl no está disponible tras $($maxRetries * $retryDelay) segundos. Se omite la creación del equipo de Teams."
                    } else {
                        Create-TeamsForSite -Environment $shpHostUrl -Body $body -Bodies $bodies -Connection $connection                          
                    }
                }
                catch {
                    write-host "Error creating teams $($siteUrl). Exception $($_)"
                }
            }
        }
        catch {
            write-host "Error creating site $($siteUrl)"
            Write-Host $_.Exception.Message 
        }
    }
}

# Function to create security groups email enabled for the site
# works with DevOps and local development
function Create-ALMAwarePnPGroupsForSite {
    param (
        [Parameter(Mandatory = $true)]
        [string]$siteTitle,
        [Parameter(Mandatory = $true)]
        [string]$siteOwner
    )

    Write-Host "Creating groups for: $siteTitle "
    # fix for jwt discrepancy between pnp.powershell and exchangeonline
    $result = Start-Job -ArgumentList $siteTitle, $siteOwner -Scriptblock { 
        param (
            [Parameter(Mandatory = $true)]
            [string]$siteTitle,
            [Parameter(Mandatory = $true)]
            [string]$siteOwner
        )
        Import-module ..\ALM\scripts\_functions.ps1 -Force
        $ExchangeAvailable = $true
        try {
            Get-ALMAwareExchangeOnlineConnection
        }
        catch {
            Write-Warning "Exchange Online no disponible: $($_.Exception.Message). Bypass activo para $siteTitle."
            $ExchangeAvailable = $false
        }

        if (-not $ExchangeAvailable) {
            Write-Warning "Bypass Exchange activo: grupos de distribucion omitidos para $siteTitle."
            return @{}
        }

        try {
            $internalGroups = @{}
            $members = Get-DistributionGroup -Identity "$($siteTitle)_members" -ErrorAction SilentlyContinue
            if ($members.IsValid -ne $True) {
                $members = New-DistributionGroup -Name "$($siteTitle)_members" -Type Security -ManagedBy $siteOwner -MemberJoinRestriction Closed
            }
            $internalGroups.Add("Contoso-Org_members", $members.ExternalDirectoryObjectId) | Out-Null

            $gestorschedulers = Get-DistributionGroup -Identity "$($siteTitle)_gestorschedulers" -ErrorAction SilentlyContinue
            if ($gestorschedulers.IsValid -ne $True) {
                $gestorschedulers = New-DistributionGroup -Name "$($siteTitle)_gestorschedulers" -Type Security -ManagedBy $siteOwner -MemberJoinRestriction Closed
            }
            $internalGroups.Add("Contoso-Org_gestorschedulers", $gestorschedulers.ExternalDirectoryObjectId) | Out-Null

            $schedulers = Get-DistributionGroup -Identity "$($siteTitle)_schedulers" -ErrorAction SilentlyContinue
            if ($schedulers.IsValid -ne $True) {
                $schedulers = New-DistributionGroup -Name "$($siteTitle)_schedulers" -Type Security -ManagedBy $siteOwner -MemberJoinRestriction Closed
            }
            $internalGroups.Add("Contoso-Org_schedulers", $schedulers.ExternalDirectoryObjectId) | Out-Null

            $guests = Get-DistributionGroup -Identity "$($siteTitle)_guests" -ErrorAction SilentlyContinue
            if ($guests.IsValid -ne $True) {
                $guests = New-DistributionGroup -Name "$($siteTitle)_guests" -Type Security -ManagedBy $siteOwner -MemberJoinRestriction Closed
            }
            $internalGroups.Add("Contoso-Org_guests", $($guests.ExternalDirectoryObjectId)) | Out-Null

            $asistentemembers = Get-DistributionGroup -Identity "$($siteTitle)_asistentemembers" -ErrorAction SilentlyContinue
            if ($asistentemembers.IsValid -ne $True) {
                $asistentemembers = New-DistributionGroup -Name "$($siteTitle)_asistentemembers" -Type Security -ManagedBy $siteOwner -MemberJoinRestriction Closed
            }
            $internalGroups.Add("Contoso-Org_asistentemembers", $asistentemembers.ExternalDirectoryObjectId) | Out-Null

            #creating principal group for site and adding member groups to it
            $allUsersGroup = Get-DistributionGroup -Identity $siteTitle -ErrorAction SilentlyContinue
            if ($allUsersGroup.IsValid -ne $True) {
                $allUsersGroup = New-DistributionGroup -Name $siteTitle -Type Security -ManagedBy $siteOwner -MemberJoinRestriction Closed
            }

            $createdGroups = @{}
            $internalGroups.Keys | ForEach-Object {
                $group = $internalGroups[$_]
                Add-DistributionGroupMember -Identity $allUsersGroup.Guid -Member $group -BypassSecurityGroupManagerCheck -ErrorAction SilentlyContinue | Out-Null
                $createdGroups.Add($_, "c:0t.c|tenant|$($group)") | Out-Null
            }

            $createdGroups.Add("Contoso-RootGroup", $allUsersGroup.PrimarySmtpAddress) | Out-Null

            Write-Host "Created groups for site $siteTitle !" -ForegroundColor Green

            return $createdGroups
        }
        catch {
            Write-Warning "Error creando grupos Exchange para $($siteTitle): $($_.Exception.Message). Bypass Exchange activo."
            return @{}
        }
    } | Receive-Job -AutoRemoveJob -Wait

    return $result
}

# Tenant id = 805543fa-f86a-4a60-9ace-f53521807088
# spurl =  https://test.sharepoint.com 
# admin url = https://test-admin.sharepoint.com
$configSiteUrl = "/sites/Contoso"
$connection = Get-ALMAwarePnPConnectionBySiteRelativeUrl -siteRelativeUrl $configSiteUrl

$adminConnection = Get-ALMAwarePnPConnectionToAdminSite 

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

Create-ALMAwarePnPBodiesFromConfigList 