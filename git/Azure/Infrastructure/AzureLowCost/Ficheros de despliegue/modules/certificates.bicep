param vaultId string
param vaultName string
param deploymentIdentityId string
param location string
param subjectPrefix string
param validityInMonths int = 120

@secure()
param certificatePasswordForPfx string

param downloadCertificates bool = false

param forceUpdateTag string = utcNow()

param webApiCertBaseName string = 'cert-webapi'
param stsCertBaseName string = 'cert-sts'
param devOpsCertBaseName string = 'cert-devops'

var certSubjectName = 'CN=${subjectPrefix}'
var webApiCertName = '${webApiCertBaseName}-${subjectPrefix}'
var stsCertName = '${stsCertBaseName}-${subjectPrefix}'
var devOpsCertName = '${devOpsCertBaseName}-${subjectPrefix}'

resource vault 'Microsoft.KeyVault/vaults@2020-04-01-preview' existing = {
  name: last(split(vaultId, '/'))
}

resource createWebApiCert 'Microsoft.Resources/deploymentScripts@2020-10-01' = {
  name: 'createWebApiCertificate'
  location: location
  identity: {
    type: 'UserAssigned'
    userAssignedIdentities: {
      '${deploymentIdentityId}': {}
    }
  }
  kind: 'AzurePowerShell'
  properties: {
    forceUpdateTag: forceUpdateTag
    azPowerShellVersion: '6.4'
    timeout: 'PT30M'
    arguments: '-vaultName ${vaultName} -certificateName ${webApiCertName} -subjectName ${certSubjectName} -validityInMonths ${validityInMonths} -outputPrefix webApi -outputCerts $${downloadCertificates}'
    environmentVariables: [
      {
        name: 'pfxExportPassword'
        secureValue: certificatePasswordForPfx
      }
    ]
    scriptContent: loadTextContent('../support-deployment-keyvaultcert.ps1')
    cleanupPreference: 'OnSuccess'
    retentionInterval: 'P1D'
  }
}

resource createStsCert 'Microsoft.Resources/deploymentScripts@2020-10-01' = {
  name: 'createSTSCertificate'
  location: location
  identity: {
    type: 'UserAssigned'
    userAssignedIdentities: {
      '${deploymentIdentityId}': {}
    }
  }
  kind: 'AzurePowerShell'
  properties: {
    forceUpdateTag: forceUpdateTag
    azPowerShellVersion: '6.4'
    timeout: 'PT30M'
    arguments: '-vaultName ${vaultName} -certificateName ${stsCertName} -subjectName ${certSubjectName} -validityInMonths ${validityInMonths} -outputPrefix sts -outputCerts $${downloadCertificates}'
    environmentVariables: [
      {
        name: 'pfxExportPassword'
        secureValue: certificatePasswordForPfx
      }
    ]
    scriptContent: loadTextContent('../support-deployment-keyvaultcert.ps1')
    cleanupPreference: 'OnSuccess'
    retentionInterval: 'P1D'
  }
}

resource createDevOpsCert 'Microsoft.Resources/deploymentScripts@2020-10-01' = {
  name: 'createDevOpsCertificate'
  location: location
  identity: {
    type: 'UserAssigned'
    userAssignedIdentities: {
      '${deploymentIdentityId}': {}
    }
  }
  kind: 'AzurePowerShell'
  properties: {
    forceUpdateTag: forceUpdateTag
    azPowerShellVersion: '6.4'
    timeout: 'PT30M'
    arguments: '-vaultName ${vaultName} -certificateName ${devOpsCertName} -subjectName ${certSubjectName} -validityInMonths ${validityInMonths} -outputPrefix devops -outputCerts $${downloadCertificates}'
    environmentVariables: [
      {
        name: 'pfxExportPassword'
        secureValue: certificatePasswordForPfx
      }
    ]
    scriptContent: loadTextContent('../support-deployment-keyvaultcert.ps1')
    cleanupPreference: 'OnSuccess'
    retentionInterval: 'P1D'
  }
}

output webApiCertThumbprint string = createWebApiCert.properties.outputs.webApiCertThumbprint
output stsCertThumbprint string = createStsCert.properties.outputs.stsCertThumbprint
output devOpsCertThumbprint string = createDevOpsCert.properties.outputs.devopsCertThumbprint
