param name string
param location string

resource identity 'Microsoft.ManagedIdentity/userAssignedIdentities@2018-11-30' = {
  name: name
  location: location
}

output principalId string = identity.properties.principalId
output tenantId string = identity.properties.tenantId
output id string = identity.id
