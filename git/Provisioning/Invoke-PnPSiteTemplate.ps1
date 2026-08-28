Connect-PnPOnline -Url "https://contoso.sharepoint.com/sites/CNTmd-cs-Contoso" -ClientId "00000000-0000-0000-0001-000000000001"   -Tenant "contoso.onmicrosoft.com"   -Thumbprint "AAAA1111BBBB2222CCCC3333DDDD4444EEEE5555"
Invoke-PnPSiteTemplate `
  -Path "plantilla-completa.xml" `
  -ClearNavigation `
  -OverwriteSystempropertyBagValues `
  -IgnoreDuplicateDataRowErrors `
  -Verbose
