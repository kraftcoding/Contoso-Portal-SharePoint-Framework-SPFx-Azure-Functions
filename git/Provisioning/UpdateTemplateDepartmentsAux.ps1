Import-module ..\ALM\scripts\_functions.ps1 -Force
Import-module ./CreateTeamsFunctions.ps1 -Force


# ALM_ENVIRONMENT -> devcdf, test, prod

#DEV

# Tenant id = contoso-test.onmicrosoft.com 
# spurl =  https://contoso-test.sharepoint.com 
# admin url = https://contoso-test-admin.sharepoint.com

#test

# Tenant id = test.onmicrosoft.com 
# spurl =  https://test.sharepoint.com 
# admin url = https://test-admin.sharepoint.com

# prod

# Tenant id = politicaterritorial.onmicrosoft.com 
# spurl =  https://politicaterritorial.sharepoint.com 
# admin url = https://politicaterritorial-admin.sharepoint.com



# Custom Function to Check if Site Collection Exists in Given URL
function CheckSiteExists($SiteURL) {
    return (Get-PnPTenantSite -Url $SiteURL -Connection $adminConnection -ErrorAction SilentlyContinue) -ne $null    
}

# Function to get and create bodies added to config list
# works with DevOps and local development
function Create-ALMAwarePnPBodiesFromConfigList {
    
    Write-Host "Getting bodies from config list:"
    
    $bodies = Get-PnPListItem -Connection $connection -List ConfiguracionDepartments -Query "<View><ViewFields><FieldRef Name='Title'/><FieldRef Name='TipoDepartment'/><FieldRef Name='Department'/><FieldRef Name='TipoMembresia'/><FieldRef Name='IdentificadorConferencia'/></ViewFields><Query><Where><BeginsWith><FieldRef Name='ContentTypeId' /><Value Type='ContentTypeId'>0x0120D520000639E2BEA58C8844A67C464647ED1640</Value></BeginsWith></Where></Query></View>"
    $shpHostUrl = $env:ALM_TENANTADMINURL -replace "-admin", ""

    $adminGroup = Get-PnPAzureADGroup -Identity "contoso-administration" -Connection $connection
    $supportGroup = Get-PnPAzureADGroup -Identity "contoso-support" -Connection $connection
    $ocpGroup = Get-PnPAzureADGroup -Identity "contoso-operations" -Connection $connection

    foreach ($body in $bodies) {

        try {
            $bodyTitle = $body['Title']

            if($bodyTitle -eq "mji-gr-ia"){
                $siteRelativeUrl = "/sites/$bodyTitle"
            
                $siteUrl = New-Object System.Uri(([System.Uri]$shpHostUrl), $siteRelativeUrl)
              
                $Department = $body['Department']
                $tipoDepartment = $body['TipoDepartment']
                    
                $createdGroups = Create-ALMAwarePnPGroupsForSite -siteTitle $bodyTitle 
                
                $params = $createdGroups.Clone()
                $params.Add("Contoso-Org_admins", "c:0t.c|tenant|$($adminGroup.Id)")
                $params.Add("Contoso-Org_support", "c:0t.c|tenant|$($supportGroup.Id)")
                $params.Add("Contoso-Org_Ocp", "c:0t.c|tenant|$($ocpGroup.Id)")
                $params.Add("DefaultBodyNameTaxValue", "-1;#$($Department.Label)|$($Department.TermGuid)")
                $params.Add("DefaultBodyTypeTaxValue" , "-1;#$($tipoDepartment.Label)|$($tipoDepartment.TermGuid)")
    
                Set-PnPTenantSite -Url $siteUrl -SharingCapability ExistingExternalUserSharingOnly -Owners "c:0t.c|tenant|$($adminGroup.Id)"  -Connection $adminConnection
    
                Add-PnPGroupMember -LoginName $createdGroups["Contoso-Org_members"] -Group 4 -Connection $connection
                Add-PnPGroupMember -LoginName $createdGroups["Contoso-Org_gestorschedulers"] -Group 4 -Connection $connection
                Add-PnPGroupMember -LoginName $createdGroups["Contoso-Org_schedulers"] -Group 4 -Connection $connection
                Add-PnPGroupMember -LoginName $createdGroups["Contoso-Org_asistentemembers"] -Group 4 -Connection $connection
    
                Write-Host $params
                $siteConnection = Get-ALMAwarePnPConnectionBySiteRelativeUrl -siteRelativeUrl $siteRelativeUrl
                Invoke-PnPSiteTemplate -path ./templateDepartmenttest.xml -Connection $siteConnection -Parameters $params
            } else {
                Write-Host "Skipping " + $bodyTitle;
            }
            
         
        }
        catch {
            write-host "Error updating site $($siteUrl)"
            Write-Host $_.Exception.Message 
        }
    }
}

# Function to create security groups email enabled for the site
# works with DevOps and local development
function Create-ALMAwarePnPGroupsForSite {
    param (
        [Parameter(Mandatory = $true)]
        [string]$siteTitle
    )

    Write-Host "Getting groups for: $siteTitle "
   
    try {

        # Obtención de los grupos relacionados con un Department
        $GraphGroups = Get-GroupsByBody -bodyName $siteTitle -Connection $Connection

        # Relacion entre grupos para la plantilla y sufijos 
        $TableGroup = @{
            "Contoso-Org_members"          = $siteTitle + "_members"
            "Contoso-Org_gestorschedulers" = $siteTitle + "_gestorschedulers"
            "Contoso-Org_schedulers"       = $siteTitle + "_schedulers"
            "Contoso-Org_guests"         = $siteTitle + "_guests"
            "Contoso-Org_asistentemembers" = $siteTitle + "_asistentemembers" 
        }

        $templateGroups = @{}

        # Selección de grupos por role
        $TableGroup.GetEnumerator() | ForEach-Object {
            $group = $GraphGroups[$_.Value]
            if ($null -ne $group) {
                $templateGroups.Add($_.Name, "c:0t.c|tenant|$($group.id)") | Out-Null
            }
        }
        
        # Grupo general del Department
        $allUsersGroup = $GraphGroups[$siteTitle];
        if ($null -ne $allUsersGroup) {
            $templateGroups.Add("Contoso-RootGroup", "$($allUsersGroup.mail)") | Out-Null
        }

        return $templateGroups
       
    }
    catch {
        write-host "Error getting groups for site $($siteTitle):"
        Write-Host $_.Exception.Message 
        throw $_
    }
   

}

function Get-GroupsByBody {
    param (
        [Parameter(Mandatory = $true)]
        [string]$bodyName,
        [Parameter(Mandatory = $true)]
        [object]$Connection
    )

    $res = @{}
    $groups = Invoke-PnPGraphMethod -Url "groups?`$filter=startswith(displayName,'$bodyName')" -Method Get -Connection $Connection 
    foreach ($group in $groups.value) {
        $res[$group.displayName] = $group
    }
    return $res

    
}



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