param param_envtoken string
param param_apptoken string
param param_azcounter string
param param_location string = resourceGroup().location
param param_storageAccountType string = 'Standard_LRS'


module naming './modules/naming.bicep' = {
  name: 'naming'
  params: {
    env: param_envtoken
    app: param_apptoken
    counter: param_azcounter
  }
}

module identity './modules/identity.bicep' = {
  name: 'identity'
  params: {
    name: naming.outputs.identityName
    location: param_location
  }
}

module kv './modules/keyvault.bicep' = {
  name: 'kv'
  params: {
    name: naming.outputs.vaultName
    location: param_location
  }
}

module storage './modules/storage.bicep' = {
  name: 'storage'
  params: {
    name: naming.outputs.storageName
    location: param_location
    sku: param_storageAccountType
  }
}

module plan './modules/appplan.bicep' = {
  name: 'plan'
  params: {
    name: naming.outputs.appPlanName
    location: param_location
  }
}

module monitoring './modules/monitoring.bicep' = {
  name: 'monitoring'
  params: {
    logName: naming.outputs.logAnalyticsName
    appInsightsName: naming.outputs.appInsightsName
    location: param_location
  }
}

module fn './modules/functionapp.bicep' = {
  name: 'fn'
  params: {
    name: naming.outputs.fnAppName
    location: param_location
    planId: plan.outputs.id
    settings: [
      {
        name: 'AzureWebJobsStorage'
        value: storage.outputs.connectionString
      }
      {
        name: 'APPINSIGHTS_INSTRUMENTATIONKEY'
        value: monitoring.outputs.instrumentationKey
      }
    ]
    cors: []
  }
}

module kvAccess './modules/keyvault-access.bicep' = {
  name: 'kvAccess'
  params: {
    vaultId: kv.outputs.id
    policies: [
      {
        tenantId: identity.outputs.tenantId
        objectId: identity.outputs.principalId
        permissions: {
          secrets: ['get','list','set']
          certificates: ['get','list','create','import','update']
          keys: ['get','list','create','import','update']
        }
      }
    ]
  }
}

output functionName string = naming.outputs.fnAppName
output keyVaultName string = naming.outputs.vaultName

@secure()
param param_certificatePasswordForPfx string

param param_downloadCertificates bool = false
param param_aadClientId string
param param_aadTenantId string

module certs './modules/certificates.bicep' = {
  name: 'certs'
  params: {
    vaultId: kv.outputs.id
    vaultName: kv.outputs.nameOut
    deploymentIdentityId: identity.outputs.id
    location: param_location
    subjectPrefix: '${param_envtoken}-${param_apptoken}-${param_azcounter}'
    certificatePasswordForPfx: param_certificatePasswordForPfx
    downloadCertificates: param_downloadCertificates
  }
}

module auth './modules/functionapp-auth.bicep' = {
  name: 'fn-auth'
  params: {
    functionAppName: naming.outputs.fnAppName
    aadClientId: param_aadClientId
    aadTenantId: param_aadTenantId
    certThumbprint: certs.outputs.webApiCertThumbprint
  }
}

module fnAux './modules/functionapp-aux.bicep' = {
  name: 'fnAux'
  dependsOn: [
    certs
    kvAccess
  ]
  params: {
    name: '${naming.outputs.fnAppName}aux'
    location: param_location
    planId: plan.outputs.id

    baseAppSettings: [
      {
        name: 'FUNCTIONS_EXTENSION_VERSION'
        value: '~4'
      }
      {
        name: 'FUNCTIONS_WORKER_RUNTIME'
        value: 'powershell'
      }
      {
        name: 'FUNCTIONS_WORKER_RUNTIME_VERSION'
        value: '7.4'
      }
      {
        name: 'TenantId'
        value: param_aadTenantId
      }
      {
        name: 'ClientId'
        value: param_aadClientId
      }
      {
        name: 'WEBSITE_RUN_FROM_PACKAGE'
        value: 1
      }
    ]

    storageConnectionString: storage.outputs.connectionString
    storageAccountName: naming.outputs.storageName
    contentShareName: 'funcshare${naming.outputs.fnAppName}aux'

    appInsightsKey: monitoring.outputs.instrumentationKey
    appInsightsConnectionString: monitoring.outputs.connectionString

    certThumbprint: certs.outputs.devOpsCertThumbprint
  }
}
