param name string
param location string
param planId string
param settings array
param cors array

resource fn 'Microsoft.Web/sites@2021-03-01' = {
  name: name
  location: location
  kind: 'functionapp'
  identity: {
    type: 'SystemAssigned'
  }
  properties: {
    httpsOnly: true
    serverFarmId: planId
    siteConfig: {
      minTlsVersion: '1.2'
      appSettings: settings
      cors: {
        allowedOrigins: cors
      }
    }
  }
}

output id string = fn.id
output principalId string = fn.identity.principalId
output tenantId string = fn.identity.tenantId
