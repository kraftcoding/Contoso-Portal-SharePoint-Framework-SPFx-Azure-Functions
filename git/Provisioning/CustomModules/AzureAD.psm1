function Add-AzureSecurityGroup {
	[OutputType([AADGroutestsult])]
	[cmdletbinding()]
	param
	(
		[Parameter(Mandatory = $true)][string]$DisplayName,
		[Parameter(Mandatory = $false)][switch]$MailEnabled
	)
	
	try {
		$groupInfo = [AADGroutestsult]::new($MailEnabled.Istestsent)
		Write-LogInfo $MyInvocation.MyCommand "$($PSBoundParameters.GetEnumerator() | Sort-Object -property key)" 
		
		$testvGroups = Get-MgGroup -Search "displayName:$($DisplayName)" -ConsistencyLevel eventual
		$testvGroup = $testvGroups | Where-Object { $_.DisplayName -eq $DisplayName }
		if ($null -eq $testvGroup) {
			if ($MailEnabled) {
				$group = New-MgGroup -DisplayName $DisplayName -MailNickname $DisplayName.Replace(' ', '') -GroupType "Unified" -MailEnabled:$true -SecurityEnabled
				$groupInfo.Id = $group.Id
			}
			else {
				$group = New-MgGroup -DisplayName $DisplayName -MailNickname "NotSet" -SecurityEnabled -MailEnabled:$false
				$groupInfo.Id = $group.Id
			}
			Write-LogInfo $MyInvocation.MyCommand "Group $($DisplayName) created" 
		}
		else {
			$groupInfo.Id = $testvGroup.Id
			Write-LogInfo $MyInvocation.MyCommand "Group $($DisplayName) already at AzureAD" 
		}

		if ($global:azureGroupsOwner) {
			$user = Get-MgUser -UserId $global:azureGroupsOwner
			$groupId = $testvGroup ? $testvGroup.Id : $group.Id
			$payload = @{ '@odata.id' = "https://graph.microsoft.com/v1.0/users/$($user.Id)" }
			New-MgGroupOwnerByRef -GroupId $groupId -BodyParameter $payload
		}
	}
	catch {
		Write-LogException -Method $MyInvocation.MyCommand -Exception $_ -Message "General error" 
	}
	
	return $groupInfo
}

function Add-AADGroupMember {
	[cmdletbinding()]
	param
	(
		[Parameter(Mandatory = $true)][string]$GroupId,
		[Parameter(Mandatory = $true)][string]$MemberId
	)
	
	try {
		Write-LogInfo $MyInvocation.MyCommand "$($PSBoundParameters.GetEnumerator() | Sort-Object -property key)" 
		$groupMatch = Get-MgGroup -GroupId $GroupId 
		if ($null -ne $groupMatch) {

			$alreadyAdded = Confirm-MgGroupMemberGroup -GroupId $GroupId -GroupIds @($MemberId)
			New-MgGroupMember -GroupId $groupMatch.Id -DirectoryObjectId $MemberId
		}
		else {
			Write-LogError $MyInvocation.MyCommand "Group $($GroupDisplayName) not founded!" 
		}
		
	}
	catch {
		Write-LogException -Method $MyInvocation.MyCommand -Exception $_ -Message "General error" 
	}
}
