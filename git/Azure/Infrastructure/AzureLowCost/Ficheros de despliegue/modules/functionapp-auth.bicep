param functionAppName string
param aadClientId string
param aadTenantId string
param certThumbprint string

resource fn 'Microsoft.Web/sites@2021-03-01' existing = {
  name: functionAppName
}

resource auth 'Microsoft.Web/sites/config@2021-03-01' = {
  name: 'authsettingsV2'
  parent: fn
  properties: {
    globalValidation: {
      requireAuthentication: true
      excludedPaths: ['/api/noauth/*']
      unauthenticatedClientAction: 'RedirectToLoginPage'
      redirectToProvider: 'azureactivedirectory'
    }
    identityProviders: {
      azureActiveDirectory: {
        enabled: true
        isAutoProvisioned: false
        registration: {
          clientId: aadClientId
          clientSecretCertificateIssuer: 'https://sts.windows.net/${aadTenantId}/'
          clientSecretCertificateThumbprint: certThumbprint
        }
      }
    }
    login: {
      tokenStore: {
        enabled: false
      }
    }
  }
}
