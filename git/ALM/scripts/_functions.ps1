## pfx and pem certifactes will be placed in the same folder as the script
## Avalible variables; to get variable value: $env:ALM_TENANT
# ALM_TENANT -> test.onmicrosoft.com
# ALM_TENANTADMINURL -> https://test-admin.sharepoint.com
# ALM_CLIENTID -> guid of devops app
# ALM_CERTPASSWORD -> password of the pfx certificate
# ALM_SUBSCRIPTIONID -> guid of the subscription
# ALM_RESOURCEGROUP -> name of the resource group
# ALM_AZFUNCNAME -> test
# ALM_ENVIRONMENT -> devcdf, test, prod
# ALM_CERTPATH -> path to the pfx certificate
# ALM_CERTPEMPATH -> path to the pem certificate

#DEV Environment

$TENANT = "contoso-test.onmicrosoft.com"
$TENANTADMINURL = "https://contoso-test-admin.sharepoint.com"
$SUBSCRIPTIONID = "f8e1e688-dd8c-4e54-9ce2-e58d7b832c3a"
$RESOURCEGROUP = "rg-devcontoso-test-Contoso"
$CLIENTID = "5f991e16-d0f1-456d-88f3-8a27336bd859"
$AZFUNCNAME = "funcdevcontoso-testContoso001"
$THUMBPRINT = "92E775A7533AECF834D12330C1CD214555B3B608"

## test Environment

# $TENANT = "test.onmicrosoft.com"
# $TENANTADMINURL = "https://test-admin.sharepoint.com"
# $SUBSCRIPTIONID = "19dd0758-e955-4244-868e-2f4e008dc58b"
# $RESOURCEGROUP = "rg-test-Contoso"
# $CLIENTID = "d633d411-033b-4f9e-b0fa-4806398abad0"
# $AZFUNCNAME = "functestContoso002"
# $THUMBPRINT = "ADA6D5ABB56B8BD9BC894C1C3A3F0AFDE1E13C98"

## prodD Environment

# $TENANT = "politicaterritorial.onmicrosoft.com"
# $TENANTADMINURL = "https://politicaterritorial-admin.sharepoint.com"
# $SUBSCRIPTIONID = "caa35b8b-a47a-4519-96dc-cc377ad6ab8f"
# $RESOURCEGROUP = "rg-prod-Contoso"
# $CLIENTID = "58a8f98a-a429-4ed2-979b-fd0a211bf2e0"
# $AZFUNCNAME = "funcprodContoso001"
# $THUMBPRINT = "ADA6D5ABB56B8BD9BC894C1C3A3F0AFDE1E13C98"


function Ensure-PowerShellModule {
    param(
        [string]$ModuleName,
        [string]$Version = $null
    )
    
    if($Version){
        Write-Host "Ensuring ps module: $($ModuleName) with version: $($Version)"
        if (-not (Get-Module | Where-Object { $_.Name -eq $ModuleName -and ($null -eq $Version -or $_.Version -eq $Version) })) {

            if (Get-Module -ListAvailable | Where-Object { $_.Name -eq $ModuleName -and ($null -eq $Version -or $_.Version -eq $Version) }) {
                if($Version){
                    Import-Module $ModuleName -RequiredVersion $Version
                }else{
                    Import-Module $ModuleName
                }
            }
            elseif (($version -and (Find-Module -Name $ModuleName -Allowtestrelease -RequiredVersion $Version)) -or ($null -eq $Version -and (Find-Module -Name $ModuleName))) {
                if($Version){
                    Install-Module -Name $ModuleName -Force -Scope CurrentUser -RequiredVersion $Version
                    Import-Module $ModuleName -RequiredVersion $Version
                }else{
                    Install-Module -Name $ModuleName -Force -Scope CurrentUser
                    Import-Module $ModuleName
                }
            }
            else {
                if($Version){
                    Write-Host "Module: $($ModuleName) with version: $($Version) not found"
                }else{
                    Write-Host "Module: $($ModuleName) not found"
                }
            }
        }
        Write-Host "(Done) Ensuring ps module: $($ModuleName) with version: $($Version)"
    } else {
        Write-Host "Ensuring ps module: $($ModuleName)"

        if (-not (Get-Module | Where-Object { $_.Name -eq $ModuleName })) {
            # module is not loaded
            if (Get-Module -ListAvailable | Where-Object { $_.Name -eq $ModuleName }) {
                Import-Module $ModuleName
            }
            elseif (Find-Module -Name $ModuleName | Where-Object { $_.Name -eq $ModuleName }) {
                Install-Module -Name $ModuleName -Force -Scope CurrentUser
                Import-Module $ModuleName
            }
            else {
                Write-Error "Module $ModuleName not found"
            }
        }
        Write-Host "(Done) Ensuring ps module: $ModuleName" 
    }
}



# Function to get PnP Connection by Site Relative Url
# works with DevOps and local development
function Get-ALMAwarePnPConnectionBySiteRelativeUrl {
    param (
        [Parameter(Mandatory = $true)]
        [string]$siteRelativeUrl
    )

    # Load PnP and Az modules
    Ensure-PowerShellModule -ModuleName "PnP.PowerShell" -Version "2.12.0"

    # DEVOPS Script
    if (![string]::IsNullOrECNTy($env:ALM_ENVIRONMENT)) {
        $shpHostUrl = $env:ALM_TENANTADMINURL -replace "-admin", ""
        $fullUri = New-Object System.Uri(([System.Uri]$shpHostUrl), $siteRelativeUrl)

        Write-Host "Connecting to $($fullUri.AbsoluteUri)"

        return Connect-PnPOnline -Url $fullUri.AbsoluteUri -ClientId ($env:ALM_CLIENTID) -CertificatePath ($env:ALM_CERTPATH) -CertificatePassword (ConvertTo-SecureString -AsPlainText ($env:ALM_CERTPASSWORD) -Force) -Tenant ($env:ALM_TENANT) -ReturnConnection
    }
    else {
        $shpHostUrl = $TENANTADMINURL -replace "-admin", ""
        
        $fullUri = New-Object System.Uri(([System.Uri]$shpHostUrl), $siteRelativeUrl)

        Write-Host "Connecting to $($fullUri.AbsoluteUri)"

        # if ($null -ne $global:DEV_CONNCACHE) {
        #     return Connect-PnPOnline -Url $fullUri.AbsoluteUri -Connection $global:DEV_CONNCACHE -ReturnConnection 
        # }
        # else {
        #     $global:DEV_CONNCACHE = Connect-PnPOnline -Url $fullUri.AbsoluteUri -ClientId ($CLIENTID) -Tenant ($TENANT) -Thumbprint ($THUMBPRINT) -ReturnConnection
        #     return $global:DEV_CONNCACHE
        # }

        return Connect-PnPOnline -Url $fullUri.AbsoluteUri -ClientId ($CLIENTID) -Tenant ($TENANT) -Thumbprint ($THUMBPRINT) -ReturnConnection
    }
}

function Get-ALMAwarePnPConnectionToAdminSite {
    # Load PnP and Az modules
    Ensure-PowerShellModule -ModuleName "PnP.PowerShell" -Version "2.12.0"

    if (![string]::IsNullOrECNTy($env:ALM_ENVIRONMENT)) {
        return Connect-PnPOnline -Url ($env:ALM_TENANTADMINURL) -ClientId ($env:ALM_CLIENTID) -CertificatePath ($env:ALM_CERTPATH) -CertificatePassword (ConvertTo-SecureString -AsPlainText ($env:ALM_CERTPASSWORD) -Force) -Tenant ($env:ALM_TENANT) -ReturnConnection
    }
    else {
        
        Write-Host "Connecting to $($TENANTADMINURL)"
    
        # if ($null -ne $global:DEV_CONNCACHE) {
        #     return Connect-PnPOnline -Url ($TENANTADMINURL) -Connection $global:DEV_CONNCACHE -ReturnConnection 
        # }
        # else {
        #     $global:DEV_CONNCACHE = Connect-PnPOnline -Url ($TENANTADMINURL) -ClientId ($CLIENTID) -Tenant ($TENANT) -Thumbprint ($THUMBPRINT) -ReturnConnection
        #     return $global:DEV_CONNCACHE
        # }
       return Connect-PnPOnline -Url ($TENANTADMINURL) -ClientId ($CLIENTID) -Tenant ($TENANT) -Thumbprint ($THUMBPRINT) -ReturnConnection
    }
}

function Get-ALMAwareAzConnection {
    Ensure-PowerShellModule "Az.Accounts"
    Ensure-PowerShellModule "Az.Websites"

    Disable-AzContextAutosave # Disable autosave for Azure context

    if (![string]::IsNullOrECNTy($env:ALM_ENVIRONMENT)) {
        Connect-AzAccount -ServicePrincipal -ApplicationId ($env:ALM_CLIENTID) -SubscriptionId ($env:ALM_SUBSCRIPTIONID) -Tenant ($env:ALM_TENANT) -CertificatePath ($env:ALM_CERTPATH) -CertificatePassword (ConvertTo-SecureString -AsPlainText ($env:ALM_CERTPASSWORD) -Force)
    }
    else {
     
        Connect-AzAccount -Tenant ($TENANT) -SubscriptionId ($SUBSCRIPTIONID) -ApplicationId ($CLIENTID) -CertificateThumbprint ($THUMBPRINT)
    }
}

function Get-ALMAwareGraphConnection {
    Ensure-PowerShellModule "Microsoft.Graph.Authentication"

    if (![string]::IsNullOrECNTy($env:ALM_ENVIRONMENT)) {
        Connect-MgGraph -ClientId ($env:ALM_CLIENTID) -TenantId ($env:ALM_TENANT) -Certificate (New-Object System.Security.Cryptography.X509Certificates.X509Certificate2((Get-Content ($env:ALM_CERTPATH) -AsByteStream -Raw), ($env:ALM_CERTPASSWORD)))
    }
    else {
       

        Connect-MgGraph -TenantId ($TENANT) -ClientId ($CLIENTID) -CertificateThumbprint ($THUMBPRINT) 
    }
}

function Get-ALMAwareExchangeOnlineConnection {
    Ensure-PowerShellModule "ExchangeOnlineManagement"

    if (![string]::IsNullOrECNTy($env:ALM_ENVIRONMENT)) {
        Connect-ExchangeOnline -AppId ($env:ALM_CLIENTID) -Certificate (New-Object System.Security.Cryptography.X509Certificates.X509Certificate2((Get-Content ($env:ALM_CERTPATH) -AsByteStream -Raw), ($env:ALM_CERTPASSWORD))) -Organization ($env:ALM_TENANT) -ShowBanner:$false
    }
    else {
        # if ([string]::IsNullOrECNTy($env:ALM_TENANT)) {
        #     Write-Warning "ALM_TENANT is not defined: provide the tenant id"
        #     $env:ALM_TENANT = Read-Host "Please enter tenant id"
        # }

        Connect-ExchangeOnline -Organization ($TENANT) -Device -ShowBanner:$false
    }
}