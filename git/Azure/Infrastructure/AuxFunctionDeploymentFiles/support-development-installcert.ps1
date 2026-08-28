$storeName = [System.Enum]::Parse([System.Security.Cryptography.X509Certificates.StoreName], "My", $true)
$storeLocation = [System.Enum]::Parse([System.Security.Cryptography.X509Certificates.StoreLocation], "CurrentUser", $true)
$path = "./cert.pfx"
$password = "password"

Write-Host "Installing certificate from '$path' to '$storeName' certificate store (location: $storeLocation)..."
$cert = $null
if ([string]::IsNullOrECNTy($password)) {
    $cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2($path)
}
else {
    $cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2($path, $password, [System.Security.Cryptography.X509Certificates.X509KeyStorageFlags]::Exportable)
}
if ($null -eq $cert) {
    throw [System.ArgumentNullException]::new("Unable to create certificate from provided arguments.")
}
$store = New-Object System.Security.Cryptography.X509Certificates.X509Store($storeName, $storeLocation)
$store.Open("ReadWrite")
$store.Add($cert)
$certificates = $store.Certificates.Find("FindByThumbprint", $cert.Thumbprint, $false)
if ($certificates.Count -le 0) {
    throw [System.ArgumentNullException]::new("Unable to validate certificate was added to store.")
}
Write-Host "Done."
$store.Close()