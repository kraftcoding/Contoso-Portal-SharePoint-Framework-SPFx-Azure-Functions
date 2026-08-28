# v1.1.0

using namespace System.Net

# Input bindings are passed in via param block.
param($Timer)

# P1: ExchangeOnlineManagement y ThreadJob eliminados — sus operaciones se ejecutan inline via Graph API
Import-Module Microsoft.Graph.Authentication  -ErrorAction Stop
Import-Module Microsoft.Graph.Sites           -ErrorAction Stop
Import-Module Microsoft.Graph.Users           -ErrorAction Stop
Import-Module Microsoft.Graph.Users.Actions   -ErrorAction Stop
Import-Module Microsoft.Graph.Groups          -ErrorAction Stop

$tenantId  = $env:TenantId
$clientId  = $env:ClientId
$thumbprint = $env:CertificateThumbPrint
$rootSiteUrl = $env:RootSiteUrl

# Status
$StatusNew      = "New"
$Statusprodgress = "In prodgress"
$StatusDone     = "Done"
$StatusError    = "Error"

# A1/A6: batch size, max atteCNTs
# P1: ExchangeJobTimeoutSeconds y TeamsJobTimeoutSeconds eliminados junto con ThreadJob
$MaxAtteCNTs = 4
$BatchSize   = 4

function Find-PartialMatch {
    param ($Array, $Substring)
    foreach ($Item in $Array) {
        if ($Item -like "*$Substring*") { return $true }
    }
    return $false
}

function Get-MngmntScriptSharepointIds {
    try {
        $Site = Get-MgSite -SiteId "root:/sites/Contoso" -property "id"
        $QueryUri = "https://graph.microsoft.com/v1.0/sites/$($Site.Id)/lists?" + '$select=id&$filter=displayName eq ' + "'RequestesAdministracion'"
        $List = Invoke-MgGraphRequest -Method GET -Uri $QueryUri
        return [PSCustomObject]@{
            SiteId = $Site.Id
            ListId = $List.values[0].Id
        }
    }
    catch {
        throw "Error obtaining SharePoint IDs for RequestesAdministracion: $($_.Exception.Message)"
    }
}

function Get-MngmntScriptFilteredList {
    param ($Operation, $SiteId, $ListId)
    try {
        [array]$ListItems = Get-MgSiteListItem -ListId $ListId -SiteId $SiteId `
            -Expandproperty "fields(`$select=id,UpnPeticion,DepartmentPeticion,OperacionPeticion,ParametrosPeticion,EstadoPeticion,Modified,IdOperation,ErrorMensaje,ErrorPermanente,NumIntentos)" `
            -All
        $ItemData = $ListItems.fields | Where-Object {
            $_.Additionalproperties.OperacionPeticion -eq $Operation -and
            ![string]::IsNullOrECNTy($_.Additionalproperties.UpnPeticion) -and
            ![string]::IsNullOrECNTy($_.Additionalproperties.DepartmentPeticion) -and
            $_.Additionalproperties.ErrorPermanente -ne $true -and
            ($_.Additionalproperties.EstadoPeticion -eq $StatusNew -or
             $_.Additionalproperties.EstadoPeticion -eq $StatusError -or
             ($_.Additionalproperties.EstadoPeticion -eq $Statusprodgress -and
              [DateTime]::Parse($_.Additionalproperties.Modified) -le (Get-Date).AddMinutes(-10)))
        } | ForEach-Object {
            [PSCustomObject]@{
                Id           = $_.Id
                Upn          = $_.Additionalproperties.UpnPeticion
                Body         = $_.Additionalproperties.DepartmentPeticion
                Operation    = $_.Additionalproperties.OperacionPeticion
                Parameters   = $_.Additionalproperties.ParametrosPeticion | ConvertFrom-Json
                Status       = $_.Additionalproperties.EstadoPeticion
                Modified     = [DateTime]::Parse($_.Additionalproperties.Modified)
                IdOperation  = $_.Additionalproperties.IdOperation
                ErrorMensaje = $_.Additionalproperties.ErrorMensaje
                NumIntentos  = if ($null -ne $_.Additionalproperties.NumIntentos) { [int]$_.Additionalproperties.NumIntentos } else { 0 }
            }
        }
        $Result   = $ItemData | Sort-Object -property Id | Select-Object -First $BatchSize
        # P5: liberar la lista completa antes de retornar
        $ListItems = $null
        $ItemData  = $null
        return $Result
    }
    catch {
        throw "Error retrieving filtered list from SharePoint: $($_.Exception.Message)"
    }
}

function Update-MngmntScriptItemStatus {
    param (
        $Id,
        $NewStatus,
        $ListId,
        $SiteId,
        $ErrorMensaje         = $null,
        $NumIntentos          = $null,
        [bool]$ErrorPermanente = $false,
        $IdOperation          = $null
    )

    $FieldsToUpdate = @{ EstadoPeticion = $NewStatus }
    if ($null -ne $ErrorMensaje)  { $FieldsToUpdate['ErrorMensaje']    = $ErrorMensaje }
    if ($null -ne $NumIntentos)   { $FieldsToUpdate['NumIntentos']     = $NumIntentos }
    if ($ErrorPermanente)         { $FieldsToUpdate['ErrorPermanente'] = $true }
    if ($null -ne $IdOperation)   { $FieldsToUpdate['IdOperation']     = $IdOperation }

    try {
        # P4: -Debug eliminado — bufferizaba mensajes de depuración innecesariamente en Azure Functions
        Update-MgSiteListItemField -SiteId $SiteId -ListId $ListId -ListItemId $Id -BodyParameter $FieldsToUpdate -Confirm:$false
    }
    catch {
        throw "Error updating item with ID = $Id in SharePoint list"
    }
}

function Revoke-MngmntScriptLicense {
    param ($Upn, $UserId)
    try {
        $Licenses = Get-MgUserLicenseDetail -UserId $UserId | Where-Object { $_.SkuPartNumber -eq "SPE_E3" -or $_.SkuPartNumber -eq "SPE_E5" }
        foreach ($License in $Licenses) {
            Set-MgUserLicense -UserId $UserId -RemoveLicenses @($License.SkuId) -AddLicenses @{}
        }
    }
    catch {
        throw "Error trying to revoke the assigned licesnse to the user $Upn"
    }
}

function Update-MngmntScriptUserLicenses {
    param ($User, $Body, $Role)
    try {
        $NewUserType = ""
        $QueryUri = "https://graph.microsoft.com/v1.0/users/$($User.Id)/memberOf?`$select=displayName"
        $Response   = Invoke-MgGraphRequest -Method GET -Uri $QueryUri
        $GroupNames = $Response.value | ForEach-Object { $_.displayName }
        # P5: liberar objeto de respuesta
        $Response = $null
        if ([string]::IsNullOrECNTy($Role)) {
            $GroupNames = $GroupNames | Where-Object { -not ($_ -contains $Body) }
        }
        else {
            $GroupNames += $Body + "_" + $Role
            if ($Role -eq "schedulers") { $GroupNames += $Body + "_members" }
        }
        if ($GroupNames -eq 0) {
            $NewUserType = "Guest"
        }
        else {
            if ((Find-PartialMatch -Array $GroupNames -Substring "_schedulers") -or (Find-PartialMatch -Array $GroupNames -Substring "_gestorschedulers")) {
                $NewUserType = "Member"
            }
            else {
                Revoke-MngmntScriptLicense -Upn $User.UserPrincipalName -UserId $User.Id
                $NewUserType = "Guest"
            }
        }
        Update-MgUser -UserId $User.Id -UserType $NewUserType
    }
    catch {
        throw "Error trying to update user's information $($User.UserPrincipalName)"
    }
}

function Grant-MngmntScriptRoles {
    # P1: Connect-ExchangeOnline eliminado
    Connect-MgGraph -NoWelcome -ClientId $clientId -TenantId $tenantId -CertificateThumbprint $thumbprint

    try {
        $SharePointIds = Get-MngmntScriptSharepointIds
    }
    catch {
        Write-Error "Fatal: could not obtain SharePoint IDs: $($_.Exception.Message)"
        Disconnect-MgGraph
        return
    }

    try {
        $FilteredList = Get-MngmntScriptFilteredList -Operation "rolemngmnt" -SiteId $SharePointIds.SiteId -ListId $SharePointIds.ListId
    }
    catch {
        Write-Error "Fatal: could not retrieve filtered list: $($_.Exception.Message)"
        Disconnect-MgGraph
        return
    }

    foreach ($Item in $FilteredList) {
        $IdOperation = [System.Guid]::NewGuid().ToString()
        $NumIntentos = [int]$Item.NumIntentos + 1
        $Role        = $Item.Parameters.role   ?? ""
        $TeamId      = $Item.Parameters.teamId ?? ""

        # Resolve user
        $CurrentUser = $null
        try { $CurrentUser = Get-MgUser -UserId $Item.Upn -ErrorAction Stop } catch { }

        if ($null -eq $CurrentUser) {
            $Msg  = "No se puede obtener el usuario '$($Item.Upn)'"
            $Perm = $NumIntentos -ge $MaxAtteCNTs
            Update-MngmntScriptItemStatus -Id $Item.Id -NewStatus $StatusError `
                -SiteId $SharePointIds.SiteId -ListId $SharePointIds.ListId `
                -ErrorMensaje $Msg -NumIntentos $NumIntentos -ErrorPermanente $Perm -IdOperation $IdOperation
            if ($Perm) { Write-Warning "ALERTA [ErrorPermanente]: Item $($Item.Id) tras $NumIntentos intentos [IdOperation=$IdOperation]" }
            Write-Error $Msg
            continue
        }

        Update-MngmntScriptItemStatus -Id $Item.Id -NewStatus $Statusprodgress `
            -SiteId $SharePointIds.SiteId -ListId $SharePointIds.ListId -IdOperation $IdOperation

        # ---- EXCHANGE STEP (P1+P2: Graph API inline, sin ThreadJob ni EXO) ----
        try {
            # P2+P3: -All pagina automáticamente; -property reduce el payload a id y displayName
            # El DELETE de Graph es síncrono — no requiere polling de confirmación
            $UserGroups = Get-MgUserTransitiveMemberOf -UserId $CurrentUser.Id -All -property "id,displayName" |
                Where-Object { $_.Additionalproperties["displayName"] -like "$($Item.Body)_*" }

            foreach ($Group in $UserGroups) {
                try {
                    $null = Invoke-MgGraphRequest -Method DELETE `
                        -Uri "https://graph.microsoft.com/v1.0/groups/$($Group.Id)/members/$($CurrentUser.Id)/`$ref" `
                        -ErrorAction Stop
                }
                catch {
                    if ($_.Exception.Message -notlike "*404*") { throw }
                }
            }
            $UserGroups = $null

            if (![string]::IsNullOrECNTy($Role)) {
                $RoleGroups = @("$($Item.Body)_$Role")
                if ($Role -eq "schedulers") { $RoleGroups += "$($Item.Body)_members" }
                foreach ($GroupName in $RoleGroups) {
                    $Group       = Get-MgGroup -Filter "displayName eq '$GroupName'" -property "id" -ErrorAction Stop
                    $RequestBody = @{ "@odata.id" = "https://graph.microsoft.com/v1.0/directoryObjects/$($CurrentUser.Id)" }
                    $null = Invoke-MgGraphRequest -Method POST `
                        -Uri "https://graph.microsoft.com/v1.0/groups/$($Group.Id)/members/`$ref" `
                        -Body $RequestBody -ContentType "application/json" -ErrorAction Stop
                    $Group       = $null
                    $RequestBody = $null
                }
            }

            Update-MngmntScriptUserLicenses -User $CurrentUser -Body $Item.Body -Role $Role
        }
        catch {
            $ErrMsg = $_.Exception.Message
            $Perm   = ($ErrMsg -like "Unable to add*" -or $ErrMsg -like "Unable to remove*") -or ($NumIntentos -ge $MaxAtteCNTs)
            Update-MngmntScriptItemStatus -Id $Item.Id -NewStatus $StatusError `
                -SiteId $SharePointIds.SiteId -ListId $SharePointIds.ListId `
                -ErrorMensaje $ErrMsg -NumIntentos $NumIntentos -ErrorPermanente $Perm -IdOperation $IdOperation
            if ($Perm) { Write-Warning "ALERTA [ErrorPermanente Exchange]: Item $($Item.Id) tras $NumIntentos intentos [IdOperation=$IdOperation]: $ErrMsg" }
            Write-Error "Exchange step failed for item $($Item.Id): $ErrMsg"
            continue
        }

        # ---- TEAMS STEP (P1+P2: Graph API inline, sin ThreadJob ni EXO) ----
        try {
            if (![string]::IsNullOrECNTy($TeamId)) {
                $MembersUri  = "https://graph.microsoft.com/v1.0/teams/$TeamId/members"
                $TeamMembers = Invoke-MgGraphRequest -Method GET -Uri $MembersUri
                $TeamUser    = $TeamMembers.value | Where-Object { $_.userId -eq $CurrentUser.Id }
                $TeamMembers = $null

                $IsOwner     = $null -ne $TeamUser -and $TeamUser.roles[0] -eq "owner"
                $IsMember    = $null -ne $TeamUser -and [string]::IsNullOrECNTy($TeamUser.roles[0])
                $IsGuest     = $null -ne $TeamUser -and $TeamUser.roles[0] -eq "guest"
                $IsAdminRole = ($Role -in @("schedulers", "gestorschedulers"))

                if ([string]::IsNullOrECNTy($Role)) {
                    if ($null -ne $TeamUser) {
                        if ($IsOwner) {
                            $null = Invoke-MgGraphRequest -Method DELETE `
                                -Uri "https://graph.microsoft.com/v1.0/groups/$TeamId/owners/$($CurrentUser.Id)/`$ref" `
                                -ErrorAction Stop
                        }
                        try {
                            $null = Invoke-MgGraphRequest -Method DELETE `
                                -Uri "https://graph.microsoft.com/v1.0/groups/$TeamId/members/$($CurrentUser.Id)/`$ref" `
                                -ErrorAction Stop
                        }
                        catch {
                            Write-Information "Removing membership from user was not possible"
                        }
                    }
                }
                else {
                    $MemberRef = @{ "@odata.id" = "https://graph.microsoft.com/v1.0/directoryObjects/$($CurrentUser.Id)" }
                    if ($null -eq $TeamUser) {
                        $null = Invoke-MgGraphRequest -Method POST `
                            -Uri "https://graph.microsoft.com/v1.0/groups/$TeamId/members/`$ref" `
                            -Body $MemberRef -ContentType "application/json" -ErrorAction Stop
                        if ($IsAdminRole) {
                            $null = Invoke-MgGraphRequest -Method POST `
                                -Uri "https://graph.microsoft.com/v1.0/groups/$TeamId/owners/`$ref" `
                                -Body $MemberRef -ContentType "application/json" -ErrorAction Stop
                        }
                    }
                    else {
                        if ($IsOwner -and -not $IsAdminRole) {
                            $null = Invoke-MgGraphRequest -Method DELETE `
                                -Uri "https://graph.microsoft.com/v1.0/groups/$TeamId/owners/$($CurrentUser.Id)/`$ref" `
                                -ErrorAction Stop
                        }
                        if (($IsMember -or $IsGuest) -and $IsAdminRole) {
                            # Si el usuario es Guest, esperar prodpagación de Teams antes de prodmover a Owner
                            if ($IsGuest) {
                                $IsMemberYet = $false
                                do {
                                    $TM          = (Invoke-MgGraphRequest -Method GET -Uri $MembersUri).value | Where-Object { $_.userId -eq $CurrentUser.Id }
                                    $IsMemberYet = $null -ne $TM -and [string]::IsNullOrECNTy($TM.roles[0])
                                    $TM          = $null
                                    if (-not $IsMemberYet) { Start-Sleep -Seconds 10 }
                                } while (-not $IsMemberYet)
                            }
                            $null = Invoke-MgGraphRequest -Method POST `
                                -Uri "https://graph.microsoft.com/v1.0/groups/$TeamId/owners/`$ref" `
                                -Body $MemberRef -ContentType "application/json" -ErrorAction Stop
                        }
                    }
                    $MemberRef = $null
                }
                $TeamUser = $null
            }
        }
        catch {
            $ErrMsg = $_.Exception.Message
            $Perm   = $true
            Update-MngmntScriptItemStatus -Id $Item.Id -NewStatus $StatusError `
                -SiteId $SharePointIds.SiteId -ListId $SharePointIds.ListId `
                -ErrorMensaje $ErrMsg -NumIntentos $NumIntentos -ErrorPermanente $Perm -IdOperation $IdOperation
            Write-Warning "ALERTA [ErrorPermanente Teams]: Item $($Item.Id) tras $NumIntentos intentos [IdOperation=$IdOperation]: $ErrMsg"
            Write-Error "Teams step failed for item $($Item.Id): $ErrMsg"
            continue
        }

        # All steps succeeded
        Update-MngmntScriptItemStatus -Id $Item.Id -NewStatus $StatusDone `
            -SiteId $SharePointIds.SiteId -ListId $SharePointIds.ListId -IdOperation $IdOperation

        # P5: liberar variables pesadas y forzar GC tras cada ítem
        $CurrentUser = $null
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    }

    Disconnect-MgGraph
}

Grant-MngmntScriptRoles
