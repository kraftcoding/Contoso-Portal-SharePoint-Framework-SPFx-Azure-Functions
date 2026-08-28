param logName string
param appInsightsName string
param location string

resource log 'Microsoft.OperationalInsights/workspaces@2021-12-01-preview' = {
  name: logName
  location: location
  properties: {
    retentionInDays: 30
    features: {
      searchVersion: 1
    }
    sku: {
      name: 'PerGB2018'
    }
  }
}

resource appi 'Microsoft.Insights/components@2020-02-02' = {
  name: appInsightsName
  location: location
  kind: 'web'
  properties: {
    Application_Type: 'web'
    WorkspaceResourceId: log.id
  }
}

output instrumentationKey string = reference(appi.id, '2015-05-01').InstrumentationKey
output connectionString string = 'InstrumentationKey=${reference(appi.id, '2015-05-01').InstrumentationKey}'
