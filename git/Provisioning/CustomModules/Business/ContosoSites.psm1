
function Add-ContosoHubSite {
    try {
        Write-LogInfo $MyInvocation.MyCommand "$($PSBoundParameters.GetEnumerator() | Sort-Object -property key)"
        $landingSiteInfo = $global:sitesToCreate.site | Where-Object { $_.IsLanding -eq $true }
        if ($null -ne $landingSiteInfo) {
            $addedHubSite = Add-SiteWithCheck -Title $landingSiteInfo.Title `
                -Alias $landingSiteInfo.Alias `
                -Lcid $landingSiteInfo.Lcid `
                -SiteType $landingSiteInfo.Type `
                -Connection $global:adminConnection `
                -Wait $true `
                -Owners $landingSiteInfo.Owners.Split(",") `
                -Members $landingSiteInfo.Members.Split(",")

            Register-PnPHubSite -Site $addedHubSite -Connection $global:adminConnection
        }
        else {
            Write-LogError $MyInvocation.MyCommand "Not landing site is configured" 
        }
    }
    catch {
        Write-LogException -Method $MyInvocation.MyCommand -Exception $_ -Message "Error creating landing site" 
        throw $_
    } 
}

function Invoke-ContosoHubSiteTemplate {
    try {
        Write-LogInfo $MyInvocation.MyCommand "$($PSBoundParameters.GetEnumerator() | Sort-Object -property key)"
        $landingSiteInfo = $global:sitesToCreate.site | Where-Object { $_.IsLanding -eq $true }
        if ($null -ne $landingSiteInfo) {
            $Site = Get-SiteByAlias -alias $landingSiteInfo.Alias -Connection $global:adminConnection
            $siteConnection = Connect-Site -SiteUrl $Site.Url 	

            Invoke-PnPSiteTemplate -Path $global:templateHubPath `
                -Connection $siteConnection
        }
        else {
            Write-LogError $MyInvocation.MyCommand "Not landing site is configured" 
        }
    }
    catch {
        Write-LogException -Method $MyInvocation.MyCommand -Exception $_ -Message "Error invoking template landing site" 
        throw $_
    } 
}

function Add-ContosoRelatedSites {
    try {
        Write-LogInfo $MyInvocation.MyCommand "$($PSBoundParameters.GetEnumerator() | Sort-Object -property key)"
        $relatedSitesToCreate = $global:sitesToCreate.site | Where-Object { $_.IsOrg -eq $true }

        $landingSiteInfo = $global:sitesToCreate.site | Where-Object { $_.IsLanding -eq $true }
        if ($null -ne $landingSiteInfo) {
            $hubSite = Get-SiteByAlias -alias $landingSiteInfo.Alias -Connection $global:adminConnection
        }

        $relatedSitesToCreate | ForEach-Object {
            $siteInfo = $_
            $addedSite = Add-SiteWithCheck -Title $siteInfo.Title `
                -Alias $siteInfo.Alias `
                -Lcid $siteInfo.Lcid `
                -Type $siteInfo.Type `
                -Connection $global:adminConnection `
                -Wait $true `
                -Owners $siteInfo.Owners.Split(",") `
                -Members $siteInfo.Members.Split(",")

            if ($null -ne $hubSite) {
                Add-PnPHubSiteAssociation -Site $addedSite -HubSite $hubSite.Url -Connection $global:adminConnection
            }
            else {
                Write-LogError $MyInvocation.MyCommand "Hub site not founded. Unable to associate sites" 
            }
        }
    
    }
    catch {
        Write-LogException -Method $MyInvocation.MyCommand -Exception $_ -Message "Error creating related sites" 
        throw $_
    } 
}

function Invoke-ContosoRelatedSitesTemplates {
    try {
        Write-LogInfo $MyInvocation.MyCommand "$($PSBoundParameters.GetEnumerator() | Sort-Object -property key)"
        $relatedSitesToCreate = $global:sitesToCreate.site | Where-Object { $_.IsOrg -eq $true }

        $relatedSitesToCreate | ForEach-Object {
            $siteInfo = $_
            $Site = Get-SiteByAlias -alias $siteInfo.Alias -Connection $global:adminConnection
            $siteConnection = Connect-Site -SiteUrl $Site.Url 	
            $parameters = Add-ContosoSiteSecurityGroups -groupInfo $siteInfo
            $templateParameters = $parameters.Clone()

            Invoke-PnPSiteTemplate -Path $global:templateOrgPath `
                -Connection $siteConnection `
                -Parameters $templateParameters
        }
    
    }
    catch {
        Write-LogException -Method $MyInvocation.MyCommand -Exception $_ -Message "Error setting site template" 
        throw $_
    }
}
