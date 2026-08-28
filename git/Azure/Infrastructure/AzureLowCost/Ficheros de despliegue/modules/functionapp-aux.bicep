param name string
param location string
param planId string

@description('App settings base (sin secretos)')
param baseAppSettings array

@description('Connection string de Storage desde Key Vault')
param storageConnectionString string

@description('Nombre de la storage account')
param storageAccountName string

@description('Nombre del content share para esta Function Aux')
param contentShareName string

@description('Instrumentation Key de App Insights')
param appInsightsKey string

@description('Connection string de App Insights')
param appInsightsConnectionString string

@description('Thumbprint del certificado DevOps')
param certThumbprint string

resource fnAux 'Microsoft.Web/sites@2023-12-01' = {
  name: name
  location: location
  kind: 'functionapp'
  identity: {
    type: 'SystemAssigned'
  }
  properties: {
    serverFarmId: planId
    httpsOnly: true
    siteConfig: {
      powerShellVersion: '7.4'
      netFrameworkVersion: 'v8.0'
      minTlsVersion: '1.2'
      appSettings: baseAppSettings
    }
  }
}

resource fnAux_config 'Microsoft.Web/sites/config@2018-11-01' = {
  name: 'web'
  parent: fnAux
  properties: {
    appSettings: concat(baseAppSettings, [
      {
        name: 'AzureWebJobsStorage'
        value: storageConnectionString
      }
      {
        name: 'AzureWebJobsStorage__accountName'
        value: storageAccountName
      }
      {
        name: 'WEBSITE_CONTENTAZUREFILECONNECTIONSTRING'
        value: storageConnectionString
      }
      {
        name: 'WEBSITE_CONTENTSHARE'
        value: contentShareName
      }
      {
        name: 'APPINSIGHTS_INSTRUMENTATIONKEY'
        value: appInsightsKey
      }
      {
        name: 'APPLICATIONINSIGHTS_CONNECTION_STRING'
        value: appInsightsConnectionString
      }
      {
        name: 'WEBSITE_LOAD_CERTIFICATES'
        value: certThumbprint
      }
      {
        name: 'CertificateThumbPrint'
        value: certThumbprint
      }
    ])
  }
}

output id string = fnAux.id
output principalId string = fnAux.identity.principalId
output tenantId string = fnAux.identity.tenantId
