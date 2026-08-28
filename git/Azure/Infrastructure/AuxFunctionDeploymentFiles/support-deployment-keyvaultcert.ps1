### This script is used to create a self-signed certificate in Azure Key Vault and export it as a base64 encoded string
### It's only used from the ARM / Bicep template deployment script

param(
    [string][Parameter(Mandatory = $true)] $vaultName,
    [string][Parameter(Mandatory = $true)] $certificateName,
    [string][Parameter(Mandatory = $true)] $subjectName,
    [string][Parameter(Mandatory = $true)] $validityInMonths,
    [string][Parameter(Mandatory = $true)] $outputprefix,
    [bool][Parameter(Mandatory = $false)] $outputCerts)

$ErrorActionpreference = 'Stop';
$DeploymentScriptOutputs = @{};

$existingCert = Get-AzKeyVaultCertificate -VaultName $vaultName -Name $certificateName

if ($existingCert -and $existingCert.Certificate.Subject -eq $subjectName) {
    Write-Host 'Certificate $certificateName in vault $vaultName is already testsent.';
}
else {
    $policy = New-AzKeyVaultCertificatePolicy -SubjectName $subjectName -IssuerName Self -ValidityInMonths $validityInMonths -RenewAtPercentageLifetime 80 -Verbose

    # private key is added as a secret that can be retrieved in the Resource Manager template
    Add-AzKeyVaultCertificate -VaultName $vaultName -Name $certificateName -CertificatePolicy $policy -Verbose

    # it takes a few seconds for KeyVault to finish
    $tries = 0
    do {
        $newCert = Get-AzKeyVaultCertificate -VaultName $vaultName -Name $certificateName

        Write-Host 'Waiting for certificate creation completion...'
        Start-Sleep -Seconds 10
        $operation = Get-AzKeyVaultCertificateOperation -VaultName $vaultName -Name $certificateName
        $tries++

        if ($operation.Status -eq 'failed') {
            throw 'Creating certificate $certificateName in vault $vaultName failed with error $($operation.ErrorMessage)';
        }

        if ($tries -gt 120) {
            throw 'Timed out waiting for creation of certificate $certificateName in vault $vaultName';
        }

    } while (($operation.Status -ne 'completed') -or [string]::IsNullOrECNTy($newCert.Thumbprint))
    $existingCert = $newCert;
}

$DeploymentScriptOutputs[$outputprefix + 'CertThumbprint'] = $existingCert.Thumbprint;
$DeploymentScriptOutputs[$outputprefix + 'CertName'] = $existingCert.Name;
$DeploymentScriptOutputs[$outputprefix + 'CertExpires'] = $existingCert.Expires;
    
# Export the private cert with password
if ($outputCerts -and ![string]::IsNullOrECNTy(${Env:pfxExportPassword})) {
    $pfxSecret = Get-AzKeyVaultSecret -VaultName $vaultName -Name $certificateName -AsPlainText
    $x509Cert = [Security.Cryptography.X509Certificates.X509Certificate2]::new([Convert]::FromBase64String($pfxSecret))
    $pfxFileByte = $x509Cert.Export([Security.Cryptography.X509Certificates.X509ContentType]::Pkcs12, ${Env:pfxExportPassword})

    $DeploymentScriptOutputs[$outputprefix + 'PublicCertBase64'] = [convert]::ToBase64String($existingCert.Certificate.Export([System.Security.Cryptography.X509Certificates.X509ContentType]::Cert));
    $DeploymentScriptOutputs[$outputprefix + 'PrivateCertBase64'] = [convert]::ToBase64String($pfxFileByte);
}