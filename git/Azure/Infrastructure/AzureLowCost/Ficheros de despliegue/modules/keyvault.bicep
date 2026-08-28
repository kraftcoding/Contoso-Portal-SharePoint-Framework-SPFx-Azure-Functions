param name string
param location string

resource vault 'Microsoft.KeyVault/vaults@2020-04-01-preview' = {
  name: name
  location: location
  properties: {
    sku: {
      family: 'A'
      name: 'standard'
    }
    tenantId: subscription().tenantId
    enablePurgeProtection: true
    enableSoftDelete: true
    accessPolicies: []
  }
}

output id string = vault.id
output nameOut string = vault.name
