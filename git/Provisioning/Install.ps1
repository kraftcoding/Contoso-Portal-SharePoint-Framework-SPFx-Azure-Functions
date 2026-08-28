#Requires -Modules @{ ModuleName="PnP.PowerShell"; ModuleVersion="2.2.0" }

[CmdletBinding()]
param (

    [Parameter(ParameterSetName = "PartialDeployment")]
    [switch]
    $HubCreation,
    [Parameter(ParameterSetName = "PartialDeployment")]
    [switch]
    $HubTemplate,
    [Parameter(ParameterSetName = "PartialDeployment")]
    [switch]
    $AzureADGroupsCreation,
    [Parameter(ParameterSetName = "PartialDeployment")]
    [switch]
    $RelatedSitesCreation,
    [Parameter(ParameterSetName = "PartialDeployment")]
    [switch]
    $RelatedSitesTemplate,
    [Parameter(ParameterSetName = "PartialDeployment")]
    [ValidateSet("All", "Tenant", "Sites")]
    $InstallApps,
    [Parameter(ParameterSetName = "All")]
    [switch]
    $All,
    [Parameter(Mandatory)]
    [ValidateSet("contoso-dev", "test", "prod", "contoso-test")]
    $Enviroment
)

'.\CustomModules\*.psm1' | Get-ChildItem -Recurse | Import-Module -Force -DisableNameChecking
'.\CustomModules\*.ps1' | Get-ChildItem -Recurse | Import-Module -Force -DisableNameChecking

function Connect-AllServices {
    try {
        Write-LogInfo $MyInvocation.MyCommand "$($PSBoundParameters.GetEnumerator() | Sort-Object -property key)" 

        $adminSiteUrl = [string]::Format("https://{0}-admin.sharepoint.com", $global:tenant)
        $landingUrl = [string]::Format("https://{0}.sharepoint.com/sites/{1}", $global:tenant, $global:landingSiteInfo.Alias)
        
        $global:adminConnection = Connect-Site -SiteUrl $adminSiteUrl
        $global:landingConnection = Connect-Site -SiteUrl $landingUrl
        $cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2($global:certificatePath, $global:certPassword) 
        $global:graphConnection = Connect-MgGraph -ClientId $global:clientId -TenantId $global:tenantId -Certificate $cert
    }
    catch {
        Write-LogException -Method $MyInvocation.MyCommand -Exception $_ -Message "General error" 
        throw $_
    } 
}
function Get-GlobalVariables {
    param(
        [Parameter(Mandatory = $true)][string] $scriptPath
    )
    
    [void] (New-Item -ItemType Directory -Force -Path "$scriptPath\Logs")
    $global:LogGeneralPath = Join-Path $scriptPath "Logs"
    $global:LogPath = Join-Path $global:LogGeneralPath "$((Get-Date -Format "yyyy-MM-dd"))_prodvisioningLog.csv"
    $global:LogGuid = [Guid]::NewGuid().Guid
    [xml]$config = Get-Content "$($scriptPath)\Config\Config-${Enviroment}.xml"


    $global:templateOrgPath = (Join-Path -Path (Get-Location).Path -ChildPath "/templateDepartmenttest.xml")
    $global:templateHubPath = (Join-Path -Path (Get-Location).Path -ChildPath "/templateHomeCNTtest.xml")
    $global:sitesToCreate = $config.configuration.sites
    $global:tenant = $config.configuration.tenant
    $global:landingSiteInfo = $global:sitesToCreate.site | Where-Object { $_.IsLanding -eq $true }
    $global:tenantId = $config.configuration.tenantId
    $global:clientId = $config.configuration.clientId
    $global:certificatePath = (Join-Path -Path (Get-Location).Path -ChildPath "Config/cert-$($Enviroment).pfx")
    $global:appsToDeployLocally = Get-ChildItem -Path "./Solutions/LocalDeployment" -Filter "*.sppkg"
    $global:appsToDeployGlobaly = Get-ChildItem -Path "./Solutions/GlobalDeployment" -Filter "*.sppkg"
    
    Write-Host "Enter the certificate password ($certificatePath):"
    $global:certPassword = Read-Host -AsSecureString
}

try {
    $scriptPath = split-path -parent $MyInvocation.MyCommand.Definition
    Get-GlobalVariables -scriptPath $scriptPath
    
    Start-Log
    Write-LogInfo $MyInvocation.MyCommand "$($PSBoundParameters.GetEnumerator() | Sort-Object -property key)"

    Connect-AllServices

    $templateParameters = $null

    if ($HubCreation -or $All) {
        Add-ContosoHubSite
    }

    if ($HubTemplate -or $All) {
        Invoke-ContosoHubSiteTemplate
    }

    if ($AzureADGroupsCreation -or $All) {
        Add-ContosoSecurityGroups
    }

    if ($RelatedSitesCreation -or $All) {
        Add-ContosoRelatedSites
    }

    if ($RelatedSitesTemplate -or $All) {
        Invoke-ContosoRelatedSitesTemplates
    }

    if ($InstallApps -or $All) {
        Add-AppsToDevelop -Scope $InstallApps
    }
}
catch {
    Write-LogException -Method $MyInvocation.MyCommand -Exception $_ -Message "General Exception"
    Write-Error $_
}
finally {
    Disconnect-MgGraph

    Stop-Log
}