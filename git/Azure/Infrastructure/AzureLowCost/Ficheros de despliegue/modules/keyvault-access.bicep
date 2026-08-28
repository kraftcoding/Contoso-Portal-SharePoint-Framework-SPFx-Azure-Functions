param vaultId string
param policies array

resource vault 'Microsoft.KeyVault/vaults@2020-04-01-preview' existing = {
  name: last(split(vaultId, '/'))
}

resource accessPolicies 'Microsoft.KeyVault/vaults/accessPolicies@2023-07-01' = {
  name: 'add'
  parent: vault
  properties: {
    accessPolicies: policies
  }
}
