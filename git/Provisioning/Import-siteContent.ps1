# ================================
# CONFIGURACIÓN
# ================================
$SiteUrl = "https://contoso.sharepoint.com/sites/CNTmd-cs-Contoso"
$ImportPath = "C:\Users\Iker\Documents\service-account@contoso.local\repo\ContosoDEV\prodvisioning\Local\CNTmd-cs-Contoso (content export)"

if (!(Test-Path $ImportPath)) {
    Write-Host "❌ No existe la carpeta de importación: $ImportPath"
    exit
}

# ================================
# CONEXIÓN
# ================================
Connect-PnPOnline `
    -Url $SiteUrl `
    -ClientId "00000000-0000-0000-0001-000000000001" `
    -Tenant "contoso.onmicrosoft.com" `
    -Thumbprint "AAAA1111BBBB2222CCCC3333DDDD4444EEEE5555"

# ================================
# FUNCIÓN: SUBIR CARPETAS Y ARCHIVOS
# ================================
function Upload-FolderREST {
    param(
        [string]$LocalPath,
        [string]$LibraryRootUrl,   # /sites/.../Biblioteca
        [string]$RelativeFolder    # carpeta dentro de la biblioteca
    )

    # Construir URL REST correcta
    $targetUrl = "$LibraryRootUrl/$RelativeFolder".TrimEnd("/")
    $encodedUrl = $targetUrl -replace " ", "%20"

    # Comprodbar si existe
    $apiCheck = "$($SiteUrl)/_api/web/GetFolderByServerRelativeUrl('$encodedUrl')"
    $folderExists = Invoke-PnPSteststMethod -Url $apiCheck -Method Get -ErrorAction SilentlyContinue

    if (-not $folderExists) {
        $parent = Split-Path $encodedUrl -Parent
        $name = Split-Path $encodedUrl -Leaf

        Write-Host "📁 Creando carpeta: $encodedUrl"

        $apiCreate = "$($SiteUrl)/_api/web/GetFolderByServerRelativeUrl('$parent')/Folders/add('$name')"
        Invoke-PnPSteststMethod -Url $apiCreate -Method Post | Out-Null
    }

    # Subir archivos
    Get-ChildItem -Path $LocalPath -File | ForEach-Object {
        Write-Host "  📄 Subiendo archivo: $($_.Name)"
        Add-PnPFile -Path $_.FullName -Folder $encodedUrl -ErrorAction Stop
    }

    # Subcarpetas
    Get-ChildItem -Path $LocalPath -Directory | ForEach-Object {
        Upload-FolderREST `
            -LocalPath $_.FullName `
            -LibraryRootUrl $LibraryRootUrl `
            -RelativeFolder "$RelativeFolder/$($_.Name)"
    }
}

# ================================
# IMPORTAR BIBLIOTECAS
# ================================
$localLibs = Get-ChildItem -Path $ImportPath -Directory

foreach ($lib in $localLibs) {

    Write-Host "📦 Importando biblioteca: $($lib.Name)"

    # Obtener la biblioteca real del sitio destino
    $spLib = Get-PnPList -Identity $lib.Name -ErrorAction Stop

    # Ruta real en SharePoint
    $libraryRoot = $spLib.RootFolder.ServerRelativeUrl

    Upload-FolderREST `
        -LocalPath $lib.FullName `
        -LibraryRootUrl $libraryRoot `
        -RelativeFolder ""
}

Write-Host "==============================="
Write-Host "IMPORTACIÓN COMPLETADA"
Write-Host "==============================="