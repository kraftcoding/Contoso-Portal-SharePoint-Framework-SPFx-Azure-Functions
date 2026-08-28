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
function Create-ALMAwarePnPBodiesFromConfigList {
    
    Write-Host "Getting bodies from config list:"

    $auxDictionary = @{

        "CNTmd-gr-pruebauno"= "y"
    }
    
    $bodies = Get-PnPListItem -Connection $connection -List ConfiguracionDepartments -Query "<View><ViewFields><FieldRef Name='Title'/><FieldRef Name='TipoDepartment'/><FieldRef Name='Department'/><FieldRef Name='TipoMembresia'/><FieldRef Name='IdentificadorConferencia'/></ViewFields><Query><Where><BeginsWith><FieldRef Name='ContentTypeId' /><Value Type='ContentTypeId'>0x0120D520000639E2BEA58C8844A67C464647ED1640</Value></BeginsWith></Where></Query></View>"
    $shpHostUrl = $env:ALM_TENANTADMINURL -replace "-admin", ""

    foreach ($body in $bodies) {

        try {
            $bodyTitle = $body['Title']
            $siteRelativeUrl = "/sites/$bodyTitle"

            $siteUrl = New-Object System.Uri(([System.Uri]$shpHostUrl), $siteRelativeUrl)

            if (CheckSiteExists($siteUrl)) {
                # $input = Read-Host "Do you want to remove $($bodyTitle) body? (Y/N)"
                $input = $auxDictionary[$bodyTitle];
           
                if ($input -eq "Y" -or $input -eq "y" -or $input -eq "Yes") {
                    
                    Write-Host "Starting to remove body $($bodyTitle) ..."
                    
                    Write-Host "Removing groups ..."

                    Remove-PnPGroupsForSite -siteTitle $bodyTitle -Connection $adminConnection

                    Write-Host "Removing site ..."

                    Remove-PnPTenantSite -Url $siteUrl -Connection $adminConnection -Force

                    Write-Host "Removing item list ..."

                    Remove-PnPListItem -List ConfiguracionDepartments -Identity $body.Id -Connection $connection -Force

                    Write-Host "Finished"
                }
                else {
                    Write-Host "Skipping $($bodyTitle) site"
                }
            }
        }
        catch {
            write-host "Error removing body $($siteUrl)"
            Write-Host $_.Exception.Message 
        }
    }
}

# Function to create security groups email enabled for the site
# works with DevOps and local development
function Remove-PnPGroupsForSite {
    param (
        [Parameter(Mandatory = $true)]
        [string]$siteTitle,
        [Parameter(Mandatory = $true)]
        [object]$Connection
       
    )

    Write-Host "Removing groups for: $siteTitle "
   
    try {

        $formattedGroup = Get-FormattedGroupIds -bodyName $siteTitle -Connection $Connection

        $suffix = "_members", "_gestorschedulers", "_schedulers", "_guests", "_asistentemembers", ""

        $suffix | ForEach-Object { 
            $groupKey = $siteTitle + $_
            $groupId = $formattedGroup[$groupKey]
            if ($null -ne $groupId) {
                Remove-GroupById -groupId $groupId -Connection $Connection
            }
        } 
    }
    catch {
        write-host "Error removing groups for site $($siteTitle):"
        Write-Host $_.Exception.Message 
    }
    

    return $result
}

function Remove-GroupById {
    param (
        [Parameter(Mandatory = $true)]
        [string]$groupId,
        [Parameter(Mandatory = $true)]
        [object]$Connection
    )
    Invoke-PnPGraphMethod -Url "groups/$groupId" -Method Delete -Connection $Connection
}

function Get-FormattedGroupIds {
    param (
        [Parameter(Mandatory = $true)]
        [string]$bodyName,
        [Parameter(Mandatory = $true)]
        [object]$Connection
    )

    $res = @{}
    $groups = Invoke-PnPGraphMethod -Url "groups?`$filter=startswith(displayName,'$bodyName')&`$select=displayName,id" -Method Get -Connection $Connection 
    foreach ($group in $groups.value) {
        $res[$group.displayName] = $group.id
    }
    return $res

    
}
# Tenant id = 805543fa-f86a-4a60-9ace-f53521807088
# spurl =  https://test.sharepoint.com 
# admin url = https://test-admin.sharepoint.com

# Tenant id = 00000000-0000-0000-0004-000000000004
# spurl =  https://contoso-dev.sharepoint.com 
# admin url = https://contoso-dev-admin.sharepoint.com

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