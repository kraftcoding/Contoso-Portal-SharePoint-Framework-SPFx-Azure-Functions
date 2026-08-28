Connect-PnPOnline -Url "https://test.sharepoint.com/sites/Contoso" -ClientId "d633d411-033b-4f9e-b0fa-4806398abad0" -Tenant "Contoso-test.gob.es" -Thumbprint "ADA6D5ABB56B8BD9BC894C1C3A3F0AFDE1E13C98"

Get-PnPSiteTemplate `
  -Out "plantilla-completa.xml" `
  -Handlers All `
  -IncludeAllClientSidePages `
  -PersistBrandingFiles `
  -PersistMultiLanguageResources
