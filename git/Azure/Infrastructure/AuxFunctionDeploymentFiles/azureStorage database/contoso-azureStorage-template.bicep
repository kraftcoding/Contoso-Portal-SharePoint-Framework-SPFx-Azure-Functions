param param_storageAccounts_db_name string = 'dbdevnttdexmscoopera003test'

resource storageAccounts_db_name_resource 'Microsoft.Storage/storageAccounts@2024-01-01' = {
  name: param_storageAccounts_db_name
  location: 'westeurope'
  sku: {
    name: 'Standard_RAGRS'
    tier: 'Standard'
  }
  kind: 'StorageV2'
  properties: {
    dnsEndpointType: 'Standard'
    defaultToOAuthAuthentication: false
    publicNetworkAccess: 'Enabled'
    allowCrossTenantReplication: false
    minimumTlsVersion: 'TLS1_2'
    allowBlobPublicAccess: false
    allowSharedKeyAccess: true
    largeFileSharesState: 'Enabled'
    networkAcls: {
      bypass: 'AzureServices'
      virtualNetworkRules: []
      ipRules: []
      defaultAction: 'Allow'
    }
    supportsHttpsTrafficOnly: true
    encryption: {
      requireInfrastructureEncryption: false
      services: {
        file: {
          keyType: 'Account'
          enabled: true
        }
        blob: {
          keyType: 'Account'
          enabled: true
        }
      }
      keySource: 'Microsoft.Storage'
    }
    accessTier: 'Hot'
  }
}

resource storageAccounts_db_name_default 'Microsoft.Storage/storageAccounts/blobServices@2024-01-01' = {
  parent: storageAccounts_db_name_resource
  name: 'default'
  sku: {
    name: 'Standard_RAGRS'
    tier: 'Standard'
  }
  properties: {
    cors: {
      corsRules: []
    }
    deleteRetentionPolicy: {
      allowPermanentDelete: false
      enabled: true
      days: 7
    }
    containerDeleteRetentionPolicy: {
      enabled: true
      days: 7
    }
  }
}

resource Microsoft_Storage_storageAccounts_fileServices_storageAccounts_db_name_default 'Microsoft.Storage/storageAccounts/fileServices@2024-01-01' = {
  parent: storageAccounts_db_name_resource
  name: 'default'
  sku: {
    name: 'Standard_RAGRS'
    tier: 'Standard'
  }
  properties: {
    protocolSettings: {
      smb: {}
    }
    cors: {
      corsRules: []
    }
    shareDeleteRetentionPolicy: {
      enabled: true
      days: 7
    }
  }
}

resource Microsoft_Storage_storageAccounts_queueServices_storageAccounts_db_name_default 'Microsoft.Storage/storageAccounts/queueServices@2024-01-01' = {
  parent: storageAccounts_db_name_resource
  name: 'default'
  properties: {
    cors: {
      corsRules: []
    }
  }
}

resource Microsoft_Storage_storageAccounts_tableServices_storageAccounts_db_name_default 'Microsoft.Storage/storageAccounts/tableServices@2024-01-01' = {
  parent: storageAccounts_db_name_resource
  name: 'default'
  properties: {
    cors: {
      corsRules: []
    }
  }
}

resource storageAccounts_db_name_default_Actas 'Microsoft.Storage/storageAccounts/tableServices/tables@2024-01-01' = {
  parent: Microsoft_Storage_storageAccounts_tableServices_storageAccounts_db_name_default
  name: 'Actas'
  properties: {}
  dependsOn: [
    storageAccounts_db_name_resource
  ]
}

resource storageAccounts_db_name_default_Convocatorias 'Microsoft.Storage/storageAccounts/tableServices/tables@2024-01-01' = {
  parent: Microsoft_Storage_storageAccounts_tableServices_storageAccounts_db_name_default
  name: 'Convocatorias'
  properties: {}
  dependsOn: [
    storageAccounts_db_name_resource
  ]
}

resource storageAccounts_db_name_default_DocumentosEvento 'Microsoft.Storage/storageAccounts/tableServices/tables@2024-01-01' = {
  parent: Microsoft_Storage_storageAccounts_tableServices_storageAccounts_db_name_default
  name: 'DocumentosEvento'
  properties: {}
  dependsOn: [
    storageAccounts_db_name_resource
  ]
}

resource storageAccounts_db_name_default_OrdenesDelDia 'Microsoft.Storage/storageAccounts/tableServices/tables@2024-01-01' = {
  parent: Microsoft_Storage_storageAccounts_tableServices_storageAccounts_db_name_default
  name: 'OrdenesDelDia'
  properties: {}
  dependsOn: [
    storageAccounts_db_name_resource
  ]
}

resource storageAccounts_db_name_default_OrganosCoopera 'Microsoft.Storage/storageAccounts/tableServices/tables@2024-01-01' = {
  parent: Microsoft_Storage_storageAccounts_tableServices_storageAccounts_db_name_default
  name: 'OrganosCoopera'
  properties: {}
  dependsOn: [
    storageAccounts_db_name_resource
  ]
}

resource storageAccounts_db_name_default_OrganosReoico 'Microsoft.Storage/storageAccounts/tableServices/tables@2024-01-01' = {
  parent: Microsoft_Storage_storageAccounts_tableServices_storageAccounts_db_name_default
  name: 'OrganosReoico'
  properties: {}
  dependsOn: [
    storageAccounts_db_name_resource
  ]
}

resource storageAccounts_db_name_default_RelacionAreaSectorialMinisterio 'Microsoft.Storage/storageAccounts/tableServices/tables@2024-01-01' = {
  parent: Microsoft_Storage_storageAccounts_tableServices_storageAccounts_db_name_default
  name: 'RelacionAreaSectorialMinisterio'
  properties: {}
  dependsOn: [
    storageAccounts_db_name_resource
  ]
}

resource storageAccounts_db_name_default_TareasAprobacionActa 'Microsoft.Storage/storageAccounts/tableServices/tables@2024-01-01' = {
  parent: Microsoft_Storage_storageAccounts_tableServices_storageAccounts_db_name_default
  name: 'TareasAprobacionActa'
  properties: {}
  dependsOn: [
    storageAccounts_db_name_resource
  ]
}

resource storageAccounts_db_name_default_TareasAsistencia 'Microsoft.Storage/storageAccounts/tableServices/tables@2024-01-01' = {
  parent: Microsoft_Storage_storageAccounts_tableServices_storageAccounts_db_name_default
  name: 'TareasAsistencia'
  properties: {}
  dependsOn: [
    storageAccounts_db_name_resource
  ]
}

resource storageAccounts_db_name_default_TareasCertificacion 'Microsoft.Storage/storageAccounts/tableServices/tables@2024-01-01' = {
  parent: Microsoft_Storage_storageAccounts_tableServices_storageAccounts_db_name_default
  name: 'TareasCertificacion'
  properties: {}
  dependsOn: [
    storageAccounts_db_name_resource
  ]
}

resource storageAccounts_db_name_default_TareasDelegacion 'Microsoft.Storage/storageAccounts/tableServices/tables@2024-01-01' = {
  parent: Microsoft_Storage_storageAccounts_tableServices_storageAccounts_db_name_default
  name: 'TareasDelegacion'
  properties: {}
  dependsOn: [
    storageAccounts_db_name_resource
  ]
}

resource storageAccounts_db_name_default_TareasModificacionActa 'Microsoft.Storage/storageAccounts/tableServices/tables@2024-01-01' = {
  parent: Microsoft_Storage_storageAccounts_tableServices_storageAccounts_db_name_default
  name: 'TareasModificacionActa'
  properties: {}
  dependsOn: [
    storageAccounts_db_name_resource
  ]
}

resource storageAccounts_db_name_default_Votaciones 'Microsoft.Storage/storageAccounts/tableServices/tables@2024-01-01' = {
  parent: Microsoft_Storage_storageAccounts_tableServices_storageAccounts_db_name_default
  name: 'Votaciones'
  properties: {}
  dependsOn: [
    storageAccounts_db_name_resource
  ]
}
