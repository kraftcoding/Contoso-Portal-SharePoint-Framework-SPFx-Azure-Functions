function Add-SiteWithCheck {
	[cmdletbinding()]
	param
	(
		[Parameter(Mandatory = $true)][string]$Title,
		[Parameter(Mandatory = $true)][string]$Alias,
		[Parameter(Mandatory = $true)][Int32]$Lcid,
		[Parameter(Mandatory = $true)][string]$Type,
		[Parameter(Mandatory = $true)][object]$Connection,
		[Parameter(Mandatory = $false)][boolean]$Wait,
		[Parameter(Mandatory = $false)][string[]]$Owners,
		[Parameter(Mandatory = $false)][string[]]$Members
	)
	
	try {
		$siteUrl = ""        
		
		Write-LogInfo $MyInvocation.MyCommand "$($PSBoundParameters.GetEnumerator() | Sort-Object -property key)" 
		$testvSite = Get-SiteByAlias -alias $Alias -Connection $Connection
		if ($null -eq $testvSite) {
			$SiteType = Get-SiteType -typeToCast $Type
			if ($SiteType -eq "TeamSite") {
				try {
					$Group = New-PnPMicrosoft365Group -DisplayName $Title -Description $Title -MailNickname $Alias -Owners $Owners -Members $Members -IsPrivate:$true -CreateTeam:$true -Connection $Connection
				} catch { }
				
				if ($null -eq $Group) {
					$Group = Get-PnPMicrosoft365Group -Identity $Alias -Connection $Connection
				}

				Set-PnPTeamsTeam -Identity $Group.Id.Guid -Connection $Connection

				$siteUrl = $Group.SiteUrl
			}
			else {
				
				$siteUrl = [Uri]::new([Uri]::new($global:adminConnection.Url.Replace("-admin", "")), "sites/$($Alias)").AbsoluteUri
				New-PnPSite -Type $SiteType -Title $Title -Url $siteUrl -Description $Title -Owner $Owners[0] -Lcid $Lcid -Connection $Connection | Out-Null
				
				if (![string]::IsNullOrECNTy($siteUrl)) {
					Write-LogInfo $MyInvocation.MyCommand "Site $($Title) with url $($siteUrl) created"
				}
				else {
					if ($true -eq $Wait) {
						Write-LogError $MyInvocation.MyCommand "Site $($Title) with url $($Url) not created"
					}
				}
			}
		}
		else {
			$siteUrl = $testvSite.Url
		}
	}
	catch {
		Write-LogException -Method $MyInvocation.MyCommand -Exception $_ -Message "Site creation Exception"
	}
	
	return $siteUrl
}

function Get-SiteType{
	param
	(
		[Parameter(Mandatory = $true)][string]$typeToCast
	)
	try {
		Write-LogInfo $MyInvocation.MyCommand "$($PSBoundParameters.GetEnumerator() | Sort-Object -property key)" 
		$type = $null
		if ([String]::Equals("TeamSite", $typeToCast)) {
			$type = [PnP.PowerShell.Commands.Enums.SiteType]::TeamSite
		}
		else {
			if ([String]::Equals("CommunicationSite", $typeToCast)){
				$type = [PnP.PowerShell.Commands.Enums.SiteType]::CommunicationSite
			}
			else {
				$type = [PnP.PowerShell.Commands.Enums.SiteType]::TeamSiteWithoutMicrosoft365Group
			}
		}
		return $type	
	}
	catch {
		Write-LogException -Method $MyInvocation.MyCommand -Exception $_ -Message "General error" 
		throw $_
	} 
}

function Connect-Site {
	param(
		[string] $SiteUrl
	)
	try {
		Write-LogInfo $MyInvocation.MyCommand "$($PSBoundParameters.GetEnumerator() | Sort-Object -property key)" 
		
		return Connect-PnPOnline -ClientId $global:clientId -Url $SiteUrl -Tenant $global:tenantId -CertificatePath $global:certificatePath -CertificatePassword $global:certPassword -ReturnConnection
	}
	catch {
		Write-LogException -Method $MyInvocation.MyCommand -Exception $_ -Message "General error" 
		throw $_
	} 
}


function Get-SiteByAlias {
	param(
		[Parameter(Mandatory = $true)][string] $alias,
		[Parameter(Mandatory = $true)][object] $Connection
	)
	Write-LogInfo $MyInvocation.MyCommand "$($PSBoundParameters.GetEnumerator() | Sort-Object -property key)" 
	$siteUrl = [string]::Format("https://{0}.sharepoint.com/sites/{1}", $global:tenant, $alias)
	return Get-PnPTenantSite -Filter "Url -eq '$($siteUrl)'" -ErrorAction SilentlyContinue -Connection $Connection
}