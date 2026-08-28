@description('DevOps Azure AD app registration application object id: resource group onwer and keyvault access will be granted')
param param_devopsAppObjectId string = ''

@description('DevOps Azure AD app registration client id: aux authentication')
param param_devopsClientId string = ''

@description('Azure AD app registration application id.')
param param_aadClientId string = ''

@description('Azure AD app registration tenant id.')
param param_aadTenantId string = ''

@description('Azure AD app registration tenant name *.onmicosoft.com')
param param_aadTenantOrganization string = ''

@description('Default SharePoint Online site for the app. (server relative url: /sites/root)')
param param_appRootSiteUrl string = ''

@description('Default SharePoint Online host (used for CORS)')
param param_spHost string = ''

@description('Must be service principal Id of Azure App Service. https://learn.microsoft.com/en-us/azure/app-service/deploy-resource-manager-template#deploy-web-app-certificate-from-key-vault')
param param_AzureAppServiceSPId string = ''

@description('Location for all resources.')
param param_location string = resourceGroup().location

@description('Location for Application Insights')
param param_appInsightsLocation string = resourceGroup().location

@description('Update tag for certificate creation')
param param_forceUpdateTagValue string = utcNow()

@description('Storage Account type')
@allowed([
  'Standard_LRS'
  'Standard_GRS'
  'Standard_RAGRS'
])
param param_storageAccountType string = 'Standard_LRS'

@description('Specifies the Azure Function hosting plan SKU.')
@allowed([
  'EP1'
  'EP2'
  'EP3'
])
param param_functionAppPlanSku string = 'EP1'

@description('Specifies the maximum number of workers that the function app can scale out to.')
param param_maxFuncWorkerCount int = 5

@secure()
@description('Password used to encrypt the certificates private key.')
param param_certificatePasswordForPfx string
@description('Include certificates in the output of the template.')
param param_downloadCertificates bool = false

// app params
@description('Application default locale to be used.')
param param_defaultLocale string = 'es-ES'

@description('Application mailbox user principal name.')
param param_mailboxUserId string = ''

@secure()
@description('WebService Portal SIM user.')
param param_portalSimUser string = ''

@secure()
@description('WebService Portal SIM password.')
param param_portalSimPassword string = ''

@description('WebService Portal SIM service.')
param param_portalSimService string = ''

@description('WebService Portal SIM service enabled or not')
param param_portalSimEnabled bool = true

@description('Taxonomy service root site id.')
param param_taxonomyRootSiteId string = ''

@description('Taxonomy service term group id.')
param param_taxonomyAppTermGroup string = ''

@description('Is production environment (changes exception handling and logging).')
param param_isProductionEnvironment bool = true

@secure()
@description('WebService Portal inside user.')
param param_insideUser string = ''

@secure()
@description('WebService Portal inside password.')
param param_insidePassword string = ''

// use this parameters to set specific resource names
param param_vaultName string = ''
param param_appPlanName string = ''
param param_fnAppName string = ''
param param_strAccName string = ''
param param_appInsightsName string = ''
param param_deploymentIdentityName string = ''
param param_logAnalyticsName string = ''

// networing resources
@description('The name of the virtual network for virtual network integration.')
param param_vnetName string = ''
@description('The name of the virtual network subnet to be associated with the Azure Function app.')
param param_functionSubnetName string = 'snet-func'
@description('The name of the virtual network subnet used for allocating IP addresses for private endpoints.')
param param_privateEndpointSubnetName string = 'snet-pe'
@description('The IP adddress space used for the virtual network.')
//param param_vnetAddressPrefix string = '10.100.0.0/16'
param param_vnetAddressPrefixPrivateEndpoint string = '10.117.32.144/28'
param param_vnetAddressPrefixFunction string = '10.117.32.192/26'
@description('The IP address space used for the Azure Function integration subnet.')
//param param_functionSubnetAddressPrefix string = '10.100.0.0/24'
param param_functionSubnetAddressPrefix string = '10.117.32.192/26'
@description('The IP address space used for the private endpoints.')
//param param_privateEndpointSubnetAddressPrefix string = '10.100.1.0/24'
param param_privateEndpointSubnetAddressPrefix string = '10.117.32.144/28'

// use this tokens to create automatic names
param param_apptoken string = ''
param param_envtoken string = ''
param param_azcounter string = ''

param baseTime string = utcNow('u')

// resources names
var var_vaultName = (!empty(param_vaultName))
  ? param_vaultName
  : 'kv${param_envtoken}${param_apptoken}${param_azcounter}'
var var_appPlanName = (!empty(param_appPlanName))
  ? param_appPlanName
  : 'asp${param_envtoken}${param_apptoken}${param_azcounter}'
var var_fnAppName = (!empty(param_fnAppName))
  ? param_fnAppName
  : 'func${param_envtoken}${param_apptoken}${param_azcounter}'
var var_strAccName = (!empty(param_strAccName))
  ? param_strAccName
  : toLower('st${param_envtoken}${param_apptoken}${param_azcounter}')
var var_appInsightsName = (!empty(param_appInsightsName))
  ? param_appInsightsName
  : 'appi${param_envtoken}${param_apptoken}${param_azcounter}'
var var_deploymentIdentityName = (!empty(param_deploymentIdentityName))
  ? param_deploymentIdentityName
  : 'id${param_envtoken}${param_apptoken}${param_azcounter}'
var var_logAnalyticsName = (!empty(param_logAnalyticsName))
  ? param_logAnalyticsName
  : 'log${param_envtoken}${param_apptoken}${param_azcounter}'
var var_vnetName = (!empty(param_vnetName))
  ? param_vnetName
  : 'vnet${param_envtoken}${param_apptoken}${param_azcounter}'

var var_functionContentShareName = toLower('funcshare${var_fnAppName}')

var var_privateStorageFileDnsZoneName = 'privatelink.file.${environment().suffixes.storage}'
var var_privateEndpointStorageFileName = '${var_strAccName}-file-private-endpoint'
var var_privateStorageTableDnsZoneName = 'privatelink.table.${environment().suffixes.storage}'
var var_privateEndpointStorageTableName = '${var_strAccName}-table-private-endpoint'
var var_privateStorageBlobDnsZoneName = 'privatelink.blob.${environment().suffixes.storage}'
var var_privateEndpointStorageBlobName = '${var_strAccName}-blob-private-endpoint'
var var_privateStorageQueueDnsZoneName = 'privatelink.queue.${environment().suffixes.storage}'
var var_privateEndpointStorageQueueName = '${var_strAccName}-queue-private-endpoint'

var var_secretStorCNName = 'secretAzStorageConnectionString'
var var_secretAppInsightsKeyName = 'secretAppInsightsKey'
var var_secretAppInsightsCNName = 'secretAppInsightsConnectionString'
var var_secretPortalSimUser = 'secretPortalSimUser'
var var_secretPortalSimPassword = 'secretPortalSimPassword'
var var_secretInsideUser = 'secretInsideUser'
var var_secretInsidePassword = 'secretInsidePassword'
var var_certSubjectName = 'CN=${param_envtoken}-${param_apptoken}-${param_azcounter}'
var var_webApiCertName = 'cert-webapi-${param_envtoken}-${param_apptoken}-${param_azcounter}'
var var_certValidityInMonths = 120
var var_stsCertName = 'cert-sts-${param_envtoken}-${param_apptoken}-${param_azcounter}'
var var_devOpsCertName = 'cert-devops-${param_envtoken}-${param_apptoken}-${param_azcounter}'

var baseAppEnvSettings = [
  {
    name: 'FUNCTIONS_EXTENSION_VERSION'
    value: '~4'
  }
  {
    name: 'FUNCTIONS_WORKER_RUNTIME'
    value: 'dotnet-isolated'
  }
  {
    name: 'WEBSITE_CONTENTSHARE'
    value: var_functionContentShareName
  }
  {
    name: 'WEBSITE_RUN_FROM_PACKAGE'
    value: 1
  }
  {
    name: 'WEBSITE_USE_PLACEHOLDER_DOTNETISOLATED'
    value: 1
  }
  {
    name: 'ClientId'
    value: param_aadClientId
  }
  {
    name: 'TenantId'
    value: param_aadTenantId
  }
  {
    name: 'RootSiteUrl'
    value: '${param_spHost}${param_appRootSiteUrl}'
  }
  {
    name: 'DefaultLocale'
    value: param_defaultLocale
  }
  {
    name: 'GraphMailboxUserId'
    value: param_mailboxUserId
  }
  {
    name: 'TaxonomyRootSiteId'
    value: param_taxonomyRootSiteId
  }
  {
    name: 'TaxonomyAppTermGroup'
    value: param_taxonomyAppTermGroup
  }
  {
    name: 'IsProductionEnvironment'
    value: param_isProductionEnvironment
  }
  {
    name: 'WEBSITE_VNET_ROUTE_ALL'
    value: '1'
  }
  {
    name: 'WEBSITE_CONTENTOVERVNET'
    value: '1'
  }
]
var keyVaultAccessPolicies = [
  {
    tenantId: deployAcc.properties.tenantId
    objectId: deployAcc.properties.principalId
    permissions: {
      secrets: [
        'all'
      ]
      keys: [
        'all'
      ]
      certificates: [
        'all'
      ]
    }
  }
  {
    tenantId: reference(fnApp.id, '2018-11-01', 'Full').identity.tenantId
    objectId: param_devopsAppObjectId
    permissions: {
      secrets: [
        'all'
      ]
      keys: [
        'all'
      ]
      certificates: [
        'all'
      ]
    }
  }
  {
    tenantId: reference(fnApp.id, '2018-11-01', 'Full').identity.tenantId
    objectId: reference(fnApp.id, '2018-11-01', 'Full').identity.principalId
    permissions: {
      secrets: [
        'get'
      ]
      certificates: [
        'get'
      ]
    }
  }
  {
    tenantId: reference(fnApp.id, '2018-11-01', 'Full').identity.tenantId
    objectId: param_AzureAppServiceSPId
    permissions: {
      secrets: [
        'get'
      ]
      certificates: [
        'get'
      ]
    }
  }
]

resource deployAcc 'Microsoft.ManagedIdentity/userAssignedIdentities@2018-11-30' = {
  name: var_deploymentIdentityName
  location: param_location
}

@description('This is the built-in Key Vault Administrator role')
resource keyVaultAdminRoleDefinition 'Microsoft.Authorization/roleDefinitions@2022-04-01' existing = {
  scope: resourceGroup()
  name: '00482a5a-887f-4fb3-b363-3b7fe8e74483'
}

@description('This is the built-in Owner role')
resource ownerRoleDefinition 'Microsoft.Authorization/roleDefinitions@2022-04-01' existing = {
  scope: resourceGroup()
  name: '8e3af657-a8ff-443c-a75c-2fe8c4bcb635'
}

resource deployAccAsKeyVaultAdmin 'Microsoft.Authorization/roleAssignments@2022-04-01' = {
  name: guid(resourceGroup().id, deployAcc.id, keyVaultAdminRoleDefinition.id)
  scope: resourceGroup()
  properties: {
    roleDefinitionId: keyVaultAdminRoleDefinition.id
    principalId: deployAcc.properties.principalId
    principalType: 'ServicePrincipal'
  }
}

resource devopsAccAsOwner 'Microsoft.Authorization/roleAssignments@2022-04-01' = {
  name: guid(resourceGroup().id, param_devopsAppObjectId, ownerRoleDefinition.id)
  scope: resourceGroup()
  properties: {
    roleDefinitionId: ownerRoleDefinition.id
    principalId: param_devopsAppObjectId
    principalType: 'ServicePrincipal'
  }
}

resource storAcc 'Microsoft.Storage/storageAccounts@2023-01-01' = {
  name: var_strAccName
  location: param_location
  sku: {
    name: param_storageAccountType
  }
  kind: 'StorageV2'
  properties: {
    supportsHttpsTrafficOnly: true
    minimumTlsVersion: 'TLS1_2'
    allowCrossTenantReplication: false
    allowBlobPublicAccess: false
    publicNetworkAccess: 'Disabled'
    networkAcls: {
      bypass: 'None'
      defaultAction: 'Deny'
    }
  }
}
var storageCN = 'DefaultEndpointsProtocol=https;AccountName=${var_strAccName};EndpointSuffix=${environment().suffixes.storage};AccountKey=${storAcc.listKeys().keys[0].value}'

resource share 'Microsoft.Storage/storageAccounts/fileServices/shares@2022-05-01' = {
  name: '${var_strAccName}/default/${var_functionContentShareName}'
  dependsOn: [
    storAcc
  ]
}

resource storAccDefaultBlob 'Microsoft.Storage/storageAccounts/blobServices@2021-09-01' existing = {
  name: 'default'
  parent: storAcc
}

resource storAccDefaultQueue 'Microsoft.Storage/storageAccounts/queueServices@2021-09-01' existing = {
  name: 'default'
  parent: storAcc
}

resource storAccDefaultFile 'Microsoft.Storage/storageAccounts/fileServices@2021-09-01' existing = {
  name: 'default'
  parent: storAcc
}

resource storAccDefaultTable 'Microsoft.Storage/storageAccounts/tableServices@2021-09-01' existing = {
  name: 'default'
  parent: storAcc
}

resource appPlan 'Microsoft.Web/serverfarms@2018-02-01' = {
  name: var_appPlanName
  location: param_location
  sku: {
    tier: 'ElasticPremium'
    name: param_functionAppPlanSku
    size: param_functionAppPlanSku
    family: 'EP'
  }
  kind: 'elastic'
  properties: {
    maximumElasticWorkerCount: param_maxFuncWorkerCount
  }
}

resource fnApp 'Microsoft.Web/sites@2021-03-01' = {
  name: var_fnAppName
  location: param_location
  kind: 'functionapp'
  identity: {
    type: 'SystemAssigned'
  }
  properties: {
    httpsOnly: true
    serverFarmId: appPlan.id
    siteConfig: {
      minTlsVersion: '1.2'
      appSettings: concat(baseAppEnvSettings, [
        {
          name: 'AzureWebJobsStorage'
          value: storageCN
        }
        {
          name: 'WEBSITE_CONTENTAZUREFILECONNECTIONSTRING'
          value: storageCN
        }
      ])
      cors: {
        allowedOrigins: [
          param_spHost
        ]
      }
      use32BitWorkerProcess: false
      netFrameworkVersion: 'v8.0'
      functionsRuntimeScaleMonitoringEnabled: true
    }
  }
  dependsOn: [
    share
    privateStorageFileDnsZoneLink
    privateEndpointStorageFilePrivateDnsZoneGroup
    privateStorageBlobDnsZoneLink
    privateEndpointStorageBlobPrivateDnsZoneGroup
    privateStorageTableDnsZoneLink
    privateEndpointStorageTablePrivateDnsZoneGroup
    privateStorageQueueDnsZoneLink
    privateEndpointStorageQueuePrivateDnsZoneGroup
  ]
}

resource fnApp_configs 'Microsoft.Web/sites/config@2018-11-01' = {
  parent: fnApp
  name: 'web'
  properties: {
    appSettings: concat(baseAppEnvSettings, [
      {
        name: 'AzureWebJobsStorage'
        value: '@Microsoft.KeyVault(SecretUri=${vault_secret_storageCN.properties.secretUri})'
      }
      {
        name: 'WEBSITE_CONTENTAZUREFILECONNECTIONSTRING'
        value: '@Microsoft.KeyVault(SecretUri=${vault_secret_storageCN.properties.secretUri})'
      }
      {
        name: 'StorageConnectionString'
        value: '@Microsoft.KeyVault(SecretUri=${vault_secret_storageCN.properties.secretUri})'
      }
      {
        name: 'WEBSITE_LOAD_CERTIFICATES'
        value: createAddWebApiCertificate.properties.outputs.webApiCertThumbprint
      }
      {
        name: 'CertificateThumbPrint'
        value: createAddWebApiCertificate.properties.outputs.webApiCertThumbprint
      }
      {
        name: 'APPINSIGHTS_INSTRUMENTATIONKEY'
        value: '@Microsoft.KeyVault(SecretUri=${vault_secret_appInsightsKey.properties.secretUri})'
      }
      {
        name: 'APPLICATIONINSIGHTS_CONNECTION_STRING'
        value: '@Microsoft.KeyVault(SecretUri=${vault_secret_appInsightsCN.properties.secretUri})'
      }
      {
        name: 'PortalSimUser'
        value: '@Microsoft.KeyVault(SecretUri=${vault_secret_portalSimUser.properties.secretUri})'
      }
      {
        name: 'PortalSimPassword'
        value: '@Microsoft.KeyVault(SecretUri=${vault_secret_portalSimPassword.properties.secretUri})'
      }
      {
        name: 'PortalSimService'
        value: param_portalSimService
      }
      {
        name: 'PortalSimEnabled'
        value: param_portalSimEnabled
      }
      {
        name: 'InsideUser'
        value: '@Microsoft.KeyVault(SecretUri=${vault_secret_insideUser.properties.secretUri})'
      }
      {
        name: 'InsidePassword'
        value: '@Microsoft.KeyVault(SecretUri=${vault_secret_insidePassword.properties.secretUri})'
      }
    ])
  }
}

resource fnApp_auth 'Microsoft.Web/sites/config@2021-03-01' = {
  name: 'authsettingsV2'
  kind: 'string'
  parent: fnApp
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
          clientId: param_aadClientId
          clientSecretCertificateIssuer: 'https://sts.windows.net/${param_aadTenantId}/'
          clientSecretCertificateThumbprint: createAddWebApiCertificate.properties.outputs.webApiCertThumbprint
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

resource fnApp_addCertificate 'Microsoft.Web/certificates@2019-08-01' = {
  name: var_webApiCertName
  location: param_location
  properties: {
    keyVaultId: vault.id
    keyVaultSecretName: var_webApiCertName
    serverFarmId: appPlan.id
    password: ''
  }
  dependsOn: [
    createAddWebApiCertificate
  ]
}

var baseAppEnvSettingsAux = [
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
    value: param_devopsClientId
  }
  {
    name: 'RootSiteUrl'
    value: '${param_spHost}${param_appRootSiteUrl}'
  }
  {
    name: 'Organization'
    value: param_aadTenantOrganization
  }
  {
    name: 'WEBSITE_RUN_FROM_PACKAGE'
    value: 1
  }
  {
    name: 'WEBSITE_VNET_ROUTE_ALL'
    value: '1'
  }
  {
    name: 'WEBSITE_CONTENTOVERVNET'
    value: '1'
  }
]

resource shareAux 'Microsoft.Storage/storageAccounts/fileServices/shares@2022-05-01' = {
  name: '${var_strAccName}/default/${var_functionContentShareName}aux'
  dependsOn: [
    storAcc
  ]
}

resource fnAppAux 'Microsoft.Web/sites@2023-12-01' = {
  name: '${var_fnAppName}aux'
  location: param_location
  kind: 'functionapp'
  identity: {
    type: 'SystemAssigned'
  }
  properties: {
    //publicNetworkAccess: 'Disabled'
    serverFarmId: appPlan.id
    httpsOnly: true
    siteConfig: {
      powerShellVersion: '7.4'
      netFrameworkVersion: 'v8.0'
      minTlsVersion: '1.2'
      appSettings: baseAppEnvSettingsAux
    }
  }
  dependsOn: [
    privateStorageFileDnsZoneLink
    privateEndpointStorageFilePrivateDnsZoneGroup
    privateStorageBlobDnsZoneLink
    privateEndpointStorageBlobPrivateDnsZoneGroup
    privateStorageTableDnsZoneLink
    privateEndpointStorageTablePrivateDnsZoneGroup
    privateStorageQueueDnsZoneLink
    privateEndpointStorageQueuePrivateDnsZoneGroup
  ]
}

resource fnAppAux_configs 'Microsoft.Web/sites/config@2018-11-01' = {
  parent: fnAppAux
  name: 'web'
  properties: {
    appSettings: concat(baseAppEnvSettingsAux, [
      {
        name: 'AzureWebJobsStorage'
        value: '@Microsoft.KeyVault(SecretUri=${vault_secret_storageCN.properties.secretUri})'
      }
      {
        name: 'AzureWebJobsStorage__accountName'
        value: storAcc.name
      }
      {
        name: 'WEBSITE_CONTENTAZUREFILECONNECTIONSTRING'
        value: '@Microsoft.KeyVault(SecretUri=${vault_secret_storageCN.properties.secretUri})'
      }
      {
        name: 'WEBSITE_CONTENTSHARE'
        value: '${var_functionContentShareName}aux'
      }
      {
        name: 'APPINSIGHTS_INSTRUMENTATIONKEY'
        value: '@Microsoft.KeyVault(SecretUri=${vault_secret_appInsightsKey.properties.secretUri})'
      }
      {
        name: 'APPLICATIONINSIGHTS_CONNECTION_STRING'
        value: '@Microsoft.KeyVault(SecretUri=${vault_secret_appInsightsCN.properties.secretUri})'
      }
      {
        name: 'WEBSITE_LOAD_CERTIFICATES'
        value: createDevOpsCertificate.properties.outputs.devopsCertThumbprint
      }
      {
        name: 'CertificateThumbPrint'
        value: createDevOpsCertificate.properties.outputs.devopsCertThumbprint
      }
    ])
  }
  dependsOn: [
    shareAux
    addAccessToAux
    networkConfigAux
  ]
}

resource fnAppAux_addCertificate 'Microsoft.Web/certificates@2019-08-01' = {
  name: var_devOpsCertName
  location: param_location
  properties: {
    keyVaultId: vault.id
    keyVaultSecretName: var_devOpsCertName
    serverFarmId: appPlan.id
    password: ''
  }
  dependsOn: [
    createDevOpsCertificate
  ]
}

resource addAccessToAux 'Microsoft.KeyVault/vaults/accessPolicies@2023-07-01' = {
  name: 'add'
  parent: vault
  properties: {
    accessPolicies: [
      {
        tenantId: reference(fnAppAux.id, '2018-11-01', 'Full').identity.tenantId
        objectId: reference(fnAppAux.id, '2018-11-01', 'Full').identity.principalId
        permissions: {
          secrets: [
            'get'
          ]
          certificates: [
            'get'
          ]
        }
      }
    ]
  }
}

resource networkConfigAux 'Microsoft.Web/sites/networkConfig@2022-03-01' = {
  parent: fnAppAux
  name: 'virtualNetwork'
  properties: {
    subnetResourceId: resourceId('Microsoft.Network/virtualNetworks/subnets', var_vnetName, param_functionSubnetName)
    swiftSupported: true
  }
  dependsOn: [
    vnet
  ]
}

resource logAnalytics 'Microsoft.OperationalInsights/workspaces@2021-12-01-preview' = {
  name: var_logAnalyticsName
  location: param_appInsightsLocation
  properties: any({
    retentionInDays: 30
    features: {
      searchVersion: 1
    }
    sku: {
      name: 'PerGB2018'
    }
  })
}

resource appInsights 'Microsoft.Insights/components@2020-02-02' = {
  kind: 'web'
  name: var_appInsightsName
  location: param_appInsightsLocation
  tags: {
    'hidden-link:${fnApp.id}': 'Resource'
    'hidden-link:${fnAppAux.id}': 'Resource'
    displayName: 'AppInsightsComponent'
  }
  properties: {
    Application_Type: 'web'
    WorkspaceResourceId: logAnalytics.id
  }
}

resource storAccDiagSetting 'Microsoft.Insights/diagnosticSettings@2021-05-01-preview' = {
  name: '${var_strAccName}-logs'
  scope: storAcc
  properties: {
    workspaceId: logAnalytics.id
    metrics: [
      {
        category: 'Transaction'
        enabled: true
      }
    ]
  }
}

resource storAccBlobDiagSetting 'Microsoft.Insights/diagnosticSettings@2021-05-01-preview' = {
  name: '${var_strAccName}-blob-logs'
  scope: storAccDefaultBlob
  properties: {
    workspaceId: logAnalytics.id
    logs: [
      {
        categoryGroup: 'audit'
        enabled: true
      }
    ]
    metrics: [
      {
        category: 'Transaction'
        enabled: true
      }
    ]
  }
}

resource storAccQueueDiagSetting 'Microsoft.Insights/diagnosticSettings@2021-05-01-preview' = {
  name: '${var_strAccName}-queue-logs'
  scope: storAccDefaultQueue
  properties: {
    workspaceId: logAnalytics.id
    logs: [
      {
        categoryGroup: 'audit'
        enabled: true
      }
    ]
    metrics: [
      {
        category: 'Transaction'
        enabled: true
      }
    ]
  }
}

resource storAccFileDiagSetting 'Microsoft.Insights/diagnosticSettings@2021-05-01-preview' = {
  name: '${var_strAccName}-file-logs'
  scope: storAccDefaultFile
  properties: {
    workspaceId: logAnalytics.id
    logs: [
      {
        categoryGroup: 'audit'
        enabled: true
      }
    ]
    metrics: [
      {
        category: 'Transaction'
        enabled: true
      }
    ]
  }
}

resource storAccTableDiagSetting 'Microsoft.Insights/diagnosticSettings@2021-05-01-preview' = {
  name: '${var_strAccName}-table-logs'
  scope: storAccDefaultTable
  properties: {
    workspaceId: logAnalytics.id
    logs: [
      {
        categoryGroup: 'audit'
        enabled: true
      }
    ]
    metrics: [
      {
        category: 'Transaction'
        enabled: true
      }
    ]
  }
}

resource appPlanDiagSetting 'Microsoft.Insights/diagnosticSettings@2021-05-01-preview' = {
  name: '${var_appPlanName}-logs'
  scope: appPlan
  properties: {
    workspaceId: logAnalytics.id
    metrics: [
      {
        category: 'AllMetrics'
        enabled: true
      }
    ]
  }
}

resource azfuncDiagSetting 'Microsoft.Insights/diagnosticSettings@2021-05-01-preview' = {
  name: '${var_fnAppName}-logs'
  scope: fnApp
  properties: {
    workspaceId: logAnalytics.id
    logs: [
      {
        category: 'FunctionAppLogs'
        enabled: true
      }
    ]
    metrics: [
      {
        category: 'AllMetrics'
        enabled: true
      }
    ]
  }
}

resource keyVaultDiagSetting 'Microsoft.Insights/diagnosticSettings@2021-05-01-preview' = {
  name: '${var_vaultName}-logs'
  scope: vault
  properties: {
    workspaceId: logAnalytics.id
    logs: [
      {
        categoryGroup: 'audit'
        enabled: true
      }
    ]
    metrics: [
      {
        category: 'AllMetrics'
        enabled: true
      }
    ]
  }
}

resource vault 'Microsoft.KeyVault/vaults@2020-04-01-preview' = {
  name: var_vaultName
  location: param_location
  properties: {
    sku: {
      family: 'A'
      name: 'standard'
    }
    tenantId: subscription().tenantId
    enablePurgeProtection: true
    enableSoftDelete: true
    accessPolicies: keyVaultAccessPolicies
  }
  dependsOn: [
    deployAccAsKeyVaultAdmin
    devopsAccAsOwner
  ]
}

resource vault_secret_storageCN 'Microsoft.KeyVault/vaults/secrets@2020-04-01-preview' = {
  parent: vault
  name: var_secretStorCNName
  properties: {
    value: storageCN
    attributes: {
      exp: dateTimeToEpoch(dateTimeAdd(baseTime, 'P5Y'))
    }
  }
}

resource vault_secret_appInsightsKey 'Microsoft.KeyVault/vaults/secrets@2020-04-01-preview' = {
  parent: vault
  name: var_secretAppInsightsKeyName
  properties: {
    value: reference(appInsights.id, '2015-05-01').InstrumentationKey
    attributes: {
      exp: dateTimeToEpoch(dateTimeAdd(baseTime, 'P5Y'))
    }
  }
}

resource vault_secret_appInsightsCN 'Microsoft.KeyVault/vaults/secrets@2020-04-01-preview' = {
  parent: vault
  name: var_secretAppInsightsCNName
  properties: {
    value: 'InstrumentationKey=${reference(appInsights.id, '2015-05-01').InstrumentationKey}'
    attributes: {
      exp: dateTimeToEpoch(dateTimeAdd(baseTime, 'P5Y'))
    }
  }
}

resource vault_secret_portalSimUser 'Microsoft.KeyVault/vaults/secrets@2020-04-01-preview' = {
  parent: vault
  name: var_secretPortalSimUser
  properties: {
    value: param_portalSimUser
    attributes: {
      exp: dateTimeToEpoch(dateTimeAdd(baseTime, 'P5Y'))
    }
  }
}

resource vault_secret_portalSimPassword 'Microsoft.KeyVault/vaults/secrets@2020-04-01-preview' = {
  parent: vault
  name: var_secretPortalSimPassword
  properties: {
    value: param_portalSimPassword
    attributes: {
      exp: dateTimeToEpoch(dateTimeAdd(baseTime, 'P5Y'))
    }
  }
}

resource vault_secret_insideUser 'Microsoft.KeyVault/vaults/secrets@2020-04-01-preview' = {
  parent: vault
  name: var_secretInsideUser
  properties: {
    value: param_insideUser
    attributes: {
      exp: dateTimeToEpoch(dateTimeAdd(baseTime, 'P5Y'))
    }
  }
}

resource vault_secret_insidePassword 'Microsoft.KeyVault/vaults/secrets@2020-04-01-preview' = {
  parent: vault
  name: var_secretInsidePassword
  properties: {
    value: param_insidePassword
    attributes: {
      exp: dateTimeToEpoch(dateTimeAdd(baseTime, 'P5Y'))
    }
  }
}

resource createAddWebApiCertificate 'Microsoft.Resources/deploymentScripts@2020-10-01' = {
  name: 'createAddWebApiCertificate'
  location: param_location
  identity: {
    type: 'UserAssigned'
    userAssignedIdentities: {
      '${deployAcc.id}': {}
    }
  }
  kind: 'AzurePowerShell'
  properties: {
    forceUpdateTag: param_forceUpdateTagValue
    azPowerShellVersion: '6.4'
    timeout: 'PT30M'
    arguments: '-vaultName ${var_vaultName} -certificateName ${var_webApiCertName} -subjectName ${var_certSubjectName} -validityInMonths ${var_certValidityInMonths} -outputPrefix webApi -outputCerts $${param_downloadCertificates}'
    environmentVariables: [
      {
        name: 'pfxExportPassword'
        secureValue: param_certificatePasswordForPfx
      }
    ]
    scriptContent: loadTextContent('support-deployment-keyvaultcert.ps1')
    cleanupPreference: 'OnSuccess'
    retentionInterval: 'P1D'
  }
  dependsOn: [
    vault
    deployAccAsKeyVaultAdmin
  ]
}

resource createSTSCertificate 'Microsoft.Resources/deploymentScripts@2020-10-01' = {
  name: 'createSTSCertificate'
  location: param_location
  identity: {
    type: 'UserAssigned'
    userAssignedIdentities: {
      '${deployAcc.id}': {}
    }
  }
  kind: 'AzurePowerShell'
  properties: {
    forceUpdateTag: param_forceUpdateTagValue
    azPowerShellVersion: '6.4'
    timeout: 'PT30M'
    arguments: '-vaultName ${var_vaultName} -certificateName ${var_stsCertName} -subjectName ${var_certSubjectName} -validityInMonths ${var_certValidityInMonths} -outputPrefix sts -outputCerts $${param_downloadCertificates}'
    environmentVariables: [
      {
        name: 'pfxExportPassword'
        secureValue: param_certificatePasswordForPfx
      }
    ]
    scriptContent: loadTextContent('support-deployment-keyvaultcert.ps1')
    cleanupPreference: 'OnSuccess'
    retentionInterval: 'P1D'
  }
  dependsOn: [
    vault
    deployAccAsKeyVaultAdmin
  ]
}

resource createDevOpsCertificate 'Microsoft.Resources/deploymentScripts@2020-10-01' = {
  name: 'createDevOpsCertificate'
  location: param_location
  identity: {
    type: 'UserAssigned'
    userAssignedIdentities: {
      '${deployAcc.id}': {}
    }
  }
  kind: 'AzurePowerShell'
  properties: {
    forceUpdateTag: param_forceUpdateTagValue
    azPowerShellVersion: '6.4'
    timeout: 'PT30M'
    arguments: '-vaultName ${var_vaultName} -certificateName ${var_devOpsCertName} -subjectName ${var_certSubjectName} -validityInMonths ${var_certValidityInMonths} -outputPrefix devops -outputCerts $${param_downloadCertificates}'
    environmentVariables: [
      {
        name: 'pfxExportPassword'
        secureValue: param_certificatePasswordForPfx
      }
    ]
    scriptContent: loadTextContent('support-deployment-keyvaultcert.ps1')
    cleanupPreference: 'OnSuccess'
    retentionInterval: 'P1D'
  }
  dependsOn: [
    vault
    deployAccAsKeyVaultAdmin
  ]
}

resource vnet 'Microsoft.Network/virtualNetworks@2022-05-01' = {
  name: var_vnetName
  location: param_location
  properties: {
    addressSpace: {
      addressPrefixes: [
        param_vnetAddressPrefixPrivateEndpoint
        param_vnetAddressPrefixFunction
      ]
    }
    subnets: [
      {
        name: param_functionSubnetName
        properties: {
          privateEndpointNetworkPolicies: 'Enabled'
          privateLinkServiceNetworkPolicies: 'Enabled'
          delegations: [
            {
              name: 'webapp'
              properties: {
                serviceName: 'Microsoft.Web/serverFarms'
              }
            }
          ]
          addressPrefix: param_functionSubnetAddressPrefix
        }
      }
      {
        name: param_privateEndpointSubnetName
        properties: {
          privateEndpointNetworkPolicies: 'Disabled'
          privateLinkServiceNetworkPolicies: 'Enabled'
          addressPrefix: param_privateEndpointSubnetAddressPrefix
        }
      }
    ]
  }
}

resource privateStorageFileDnsZone 'Microsoft.Network/privateDnsZones@2020-06-01' = {
  name: var_privateStorageFileDnsZoneName
  location: 'global'
}

resource privateStorageBlobDnsZone 'Microsoft.Network/privateDnsZones@2020-06-01' = {
  name: var_privateStorageBlobDnsZoneName
  location: 'global'
}

resource privateStorageQueueDnsZone 'Microsoft.Network/privateDnsZones@2020-06-01' = {
  name: var_privateStorageQueueDnsZoneName
  location: 'global'
}

resource privateStorageTableDnsZone 'Microsoft.Network/privateDnsZones@2020-06-01' = {
  name: var_privateStorageTableDnsZoneName
  location: 'global'
}

resource privateStorageFileDnsZoneLink 'Microsoft.Network/privateDnsZones/virtualNetworkLinks@2020-06-01' = {
  parent: privateStorageFileDnsZone
  name: '${var_privateStorageFileDnsZoneName}-link'
  location: 'global'
  properties: {
    registrationEnabled: false
    virtualNetwork: {
      id: vnet.id
    }
  }
}

resource privateStorageBlobDnsZoneLink 'Microsoft.Network/privateDnsZones/virtualNetworkLinks@2020-06-01' = {
  parent: privateStorageBlobDnsZone
  name: '${var_privateStorageBlobDnsZoneName}-link'
  location: 'global'
  properties: {
    registrationEnabled: false
    virtualNetwork: {
      id: vnet.id
    }
  }
}

resource privateStorageTableDnsZoneLink 'Microsoft.Network/privateDnsZones/virtualNetworkLinks@2020-06-01' = {
  parent: privateStorageTableDnsZone
  name: '${var_privateStorageTableDnsZoneName}-link'
  location: 'global'
  properties: {
    registrationEnabled: false
    virtualNetwork: {
      id: vnet.id
    }
  }
}

resource privateStorageQueueDnsZoneLink 'Microsoft.Network/privateDnsZones/virtualNetworkLinks@2020-06-01' = {
  parent: privateStorageQueueDnsZone
  name: '${var_privateStorageQueueDnsZoneName}-link'
  location: 'global'
  properties: {
    registrationEnabled: false
    virtualNetwork: {
      id: vnet.id
    }
  }
}

resource privateEndpointStorageFilePrivateDnsZoneGroup 'Microsoft.Network/privateEndpoints/privateDnsZoneGroups@2022-05-01' = {
  parent: privateEndpointStorageFile
  name: 'filePrivateDnsZoneGroup'
  properties: {
    privateDnsZoneConfigs: [
      {
        name: 'config'
        properties: {
          privateDnsZoneId: privateStorageFileDnsZone.id
        }
      }
    ]
  }
}

resource privateEndpointStorageBlobPrivateDnsZoneGroup 'Microsoft.Network/privateEndpoints/privateDnsZoneGroups@2022-05-01' = {
  parent: privateEndpointStorageBlob
  name: 'blobPrivateDnsZoneGroup'
  properties: {
    privateDnsZoneConfigs: [
      {
        name: 'config'
        properties: {
          privateDnsZoneId: privateStorageBlobDnsZone.id
        }
      }
    ]
  }
}

resource privateEndpointStorageTablePrivateDnsZoneGroup 'Microsoft.Network/privateEndpoints/privateDnsZoneGroups@2022-05-01' = {
  parent: privateEndpointStorageTable
  name: 'tablePrivateDnsZoneGroup'
  properties: {
    privateDnsZoneConfigs: [
      {
        name: 'config'
        properties: {
          privateDnsZoneId: privateStorageTableDnsZone.id
        }
      }
    ]
  }
}

resource privateEndpointStorageQueuePrivateDnsZoneGroup 'Microsoft.Network/privateEndpoints/privateDnsZoneGroups@2022-05-01' = {
  parent: privateEndpointStorageQueue
  name: 'queuePrivateDnsZoneGroup'
  properties: {
    privateDnsZoneConfigs: [
      {
        name: 'config'
        properties: {
          privateDnsZoneId: privateStorageQueueDnsZone.id
        }
      }
    ]
  }
}

resource privateEndpointStorageFile 'Microsoft.Network/privateEndpoints@2022-05-01' = {
  name: var_privateEndpointStorageFileName
  location: param_location
  properties: {
    subnet: {
      id: resourceId('Microsoft.Network/virtualNetworks/subnets', var_vnetName, param_privateEndpointSubnetName)
    }
    privateLinkServiceConnections: [
      {
        name: 'MyStorageFilePrivateLinkConnection'
        properties: {
          privateLinkServiceId: storAcc.id
          groupIds: [
            'file'
          ]
        }
      }
    ]
  }
  dependsOn: [
    vnet
  ]
}

resource privateEndpointStorageBlob 'Microsoft.Network/privateEndpoints@2022-05-01' = {
  name: var_privateEndpointStorageBlobName
  location: param_location
  properties: {
    subnet: {
      id: resourceId('Microsoft.Network/virtualNetworks/subnets', var_vnetName, param_privateEndpointSubnetName)
    }
    privateLinkServiceConnections: [
      {
        name: 'MyStorageBlobPrivateLinkConnection'
        properties: {
          privateLinkServiceId: storAcc.id
          groupIds: [
            'blob'
          ]
        }
      }
    ]
  }
  dependsOn: [
    vnet
  ]
}

resource privateEndpointStorageTable 'Microsoft.Network/privateEndpoints@2022-05-01' = {
  name: var_privateEndpointStorageTableName
  location: param_location
  properties: {
    subnet: {
      id: resourceId('Microsoft.Network/virtualNetworks/subnets', var_vnetName, param_privateEndpointSubnetName)
    }
    privateLinkServiceConnections: [
      {
        name: 'MyStorageTablePrivateLinkConnection'
        properties: {
          privateLinkServiceId: storAcc.id
          groupIds: [
            'table'
          ]
        }
      }
    ]
  }
  dependsOn: [
    vnet
  ]
}

resource privateEndpointStorageQueue 'Microsoft.Network/privateEndpoints@2022-05-01' = {
  name: var_privateEndpointStorageQueueName
  location: param_location
  properties: {
    subnet: {
      id: resourceId('Microsoft.Network/virtualNetworks/subnets', var_vnetName, param_privateEndpointSubnetName)
    }
    privateLinkServiceConnections: [
      {
        name: 'MyStorageQueuePrivateLinkConnection'
        properties: {
          privateLinkServiceId: storAcc.id
          groupIds: [
            'queue'
          ]
        }
      }
    ]
  }
  dependsOn: [
    vnet
  ]
}

resource networkConfig 'Microsoft.Web/sites/networkConfig@2022-03-01' = {
  parent: fnApp
  name: 'virtualNetwork'
  properties: {
    subnetResourceId: resourceId('Microsoft.Network/virtualNetworks/subnets', var_vnetName, param_functionSubnetName)
    swiftSupported: true
  }
  dependsOn: [
    vnet
  ]
}

resource graph_billing_account 'Microsoft.GraphServices/accounts@2023-04-13' = {
  name: 'graph_billing_account'
  location: 'global'
  properties: {
    appId: param_aadClientId
  }
}

output azFuncName string = fnApp.name
output azFuncAuxName string = fnAppAux.name
output azFuncUri string = 'https://${fnApp.properties.defaultHostName}'
output keyVaultName string = vault.name

output webApiPublicCertBase64 string = (param_downloadCertificates)
  ? createAddWebApiCertificate.properties.outputs.webApiPublicCertBase64
  : ''
output webApiPublicCertFileName string = '${var_webApiCertName}.cer'
output webApiPrivateCertBase64 string = (param_downloadCertificates)
  ? createAddWebApiCertificate.properties.outputs.webApiPrivateCertBase64
  : ''
output webApiPrivateCertFileName string = '${var_webApiCertName}.pfx'
output webApiCertThumbprint string = createAddWebApiCertificate.properties.outputs.webApiCertThumbprint
output webApiCertName string = createAddWebApiCertificate.properties.outputs.webApiCertName
output webApiCertExpires string = createAddWebApiCertificate.properties.outputs.webApiCertExpires

output stsPublicCertBase64 string = (param_downloadCertificates)
  ? createSTSCertificate.properties.outputs.stsPublicCertBase64
  : ''
output stsPublicCertFileName string = '${var_stsCertName}.cer'
output stsPrivateCertBase64 string = (param_downloadCertificates)
  ? createSTSCertificate.properties.outputs.stsPrivateCertBase64
  : ''
output stsPrivateCertFileName string = '${var_stsCertName}.pfx'
output stsCertThumbprint string = createSTSCertificate.properties.outputs.stsCertThumbprint
output stsCertName string = createSTSCertificate.properties.outputs.stsCertName
output stsCertExpires string = createSTSCertificate.properties.outputs.stsCertExpires

output devOpsPublicCertBase64 string = (param_downloadCertificates)
  ? createDevOpsCertificate.properties.outputs.devopsPublicCertBase64
  : ''
output devOpsPublicCertFileName string = '${var_devOpsCertName}.cer'
output devOpsPrivateCertBase64 string = (param_downloadCertificates)
  ? createDevOpsCertificate.properties.outputs.devopsPrivateCertBase64
  : ''
output devOpsPrivateCertFileName string = '${var_devOpsCertName}.pfx'
output devOpsCertThumbprint string = createDevOpsCertificate.properties.outputs.devopsCertThumbprint
output devOpsCertName string = createDevOpsCertificate.properties.outputs.devopsCertName
output devOpsCertExpires string = createDevOpsCertificate.properties.outputs.devopsCertExpires
