

function Add-AppsToDevelop{
	param
	(
		[Parameter(Mandatory = $true)]
    	[ValidateSet("All", "Tenant", "Sites")]
    	$Scope
	)

	if(($Scope -eq "All") -or ($Scope -eq "Tenant"))
	{
		$global:appsToDeployGlobaly | ForEach-Object {
			Install-TenantAppCatalogApp -AppPath $_.FullName
		}
	}

	if(($Scope -eq "All") -or ($Scope -eq "Sites"))
	{
		$global:appsToDeployLocally | ForEach-Object {
			Update-AllSitesApp -AppId $_.BaseName
		}
	}
}


function Update-AllSitesApp {
	param
	(
		[Parameter(Mandatory = $true)][string]$AppId
	)
	
	Write-LogInfo $MyInvocation.MyCommand "$($PSBoundParameters.GetEnumerator() | Sort-Object -property key)" 

	if ($global:sitesToCreate -and $global:sitesToCreate.site -and $global:sitesToCreate.site.Count -gt 0) {

		$global:sitesToCreate.site | ForEach-Object {
			$siteConfig = $_
			$Site = Get-SiteByAlias -alias $siteConfig.Alias -Connection $global:adminConnection
			if ($null -ne $Site) {
				$siteConnection = Connect-Site -SiteUrl $Site.Url 
				$installedApp = Get-PnPApp -Identity $AppId -Connection $siteConnection
				if (($null -eq $installedApp) -or ($null -eq $installedApp.InstalledVersion)) {
					Install-PnPApp -Identity $AppId -Connection $siteConnection
				}
				else {
					Update-PnPApp -Identity $AppId -Connection $siteConnection
				}
			}
			else {
				Write-LogWarning $MyInvocation.MyCommand "Site $($siteConfig.Alias) not found exists" "Action not applied" 
			}
		} 
	}
	else {
		Write-LogInfo $MyInvocation.MyCommand "There are no sites to update app"
	}
}

function Install-TenantAppCatalogApp{
    param
	(
		[Parameter(Mandatory = $true)][string]$AppPath
	)
    try {
        Add-PnPApp -Path $AppPath -Overwrite -Publish -Connection $global:adminConnection
    } catch {
        Write-LogException -Method $MyInvocation.MyCommand -Exception $_ -Message "Error on deploy frontend app $($AppPath)."
        throw $_
    }
}