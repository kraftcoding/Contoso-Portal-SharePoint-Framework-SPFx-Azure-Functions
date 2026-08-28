param name string
param location string

resource plan 'Microsoft.Web/serverfarms@2018-02-01' = {
  name: name
  location: location
  sku: {
    name: 'Y1'
    tier: 'Dynamic'
  }
}

output id string = plan.id
