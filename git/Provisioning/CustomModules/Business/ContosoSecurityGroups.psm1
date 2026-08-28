function Add-ContosoSecurityGroups {
	Write-LogInfo $MyInvocation.MyCommand "$($PSBoundParameters.GetEnumerator() | Sort-Object -property key)" 
	try {
		
		$global:sitesToCreate.site | Where-Object { $_.IsOrg -eq $true } | ForEach-Object {
			$groupInfo = $_
			$siteCreatedGroups = Add-ContosoSiteSecurityGroups -groupInfo $groupInfo
		}       
	}
	catch {
		Write-LogException -Method $MyInvocation.MyCommand -Exception $_ -Message "General error" 
	}
}

function Add-ContosoSiteSecurityGroups {
	[cmdletbinding()]
	param
	(
		[Parameter(Mandatory = $true)][object]$groupInfo
	)
	Write-LogInfo $MyInvocation.MyCommand "$($PSBoundParameters.GetEnumerator() | Sort-Object -property key)" 
	try {
		$createdGroups = @{}
		$parentGroup = Add-AzureSecurityGroup -DisplayName "$($groupInfo.Alias)"
		$createdGroups.Add("Contoso-OrgAllUsers", $parentGroup.GetTemplateId()) | Out-Null

		if ($groupInfo.IsOrg -eq $true) {
			$internalGroups = @{}
			$internalGroups.Add("Contoso-Org_members", (Add-AzureSecurityGroup -DisplayName "$($groupInfo.Alias)_members")) | Out-Null
			$internalGroups.Add("Contoso-Org_gestorschedulers", (Add-AzureSecurityGroup -DisplayName "$($groupInfo.Alias)_gestorschedulers")) | Out-Null
			$internalGroups.Add("Contoso-Org_guests", (Add-AzureSecurityGroup -DisplayName "$($groupInfo.Alias)_guests")) | Out-Null
			$internalGroups.Add("Contoso-Org_schedulers", (Add-AzureSecurityGroup -DisplayName "$($groupInfo.Alias)_schedulers")) | Out-Null
			$internalGroups.Add("Contoso-Org_asistentemembers", (Add-AzureSecurityGroup -DisplayName "$($groupInfo.Alias)_asistentemembers")) | Out-Null

			$internalGroups.Keys | ForEach-Object {
				$group = $internalGroups[$_]
				Add-AADGroupMember -GroupId $parentGroup.Id -MemberId $group.Id
				$createdGroups.Add($_, $group.GetTemplateId()) | Out-Null
			}
		}  
	}
	catch {
		Write-LogException -Method $MyInvocation.MyCommand -Exception $_ -Message "General error" 
	}
	
	return $createdGroups
}