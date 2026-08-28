# ================================
# CONFIGURACIÓN
# ================================
$SiteUrl = "https://test.sharepoint.com/sites/CNTmd-cs-Contoso"
$ExportPath = "C:\SharePointExport"

if (!(Test-Path $ExportPath)) {
    New-Item -ItemType Directory -Path $ExportPath | Out-Null
}

# ================================
# CONEXIÓN
# ================================
Connect-PnPOnline `
    -Url $SiteUrl `
    -ClientId "d633d411-033b-4f9e-b0fa-4806398abad0" `
    -Tenant "Contoso-test.gob.es" `
    -Thumbprint "ADA6D5ABB56B8BD9BC894C1C3A3F0AFDE1E13C98"

# ================================
# FUNCIÓN: DESCARGA RECURSIVA VIA REST (FUNCIONA SIEMtest)
# ================================
function Download-FolderREST {
    param(
        [string]$ServerRelativeUrl,
        [string]$LocalPath
    )

    if (!(Test-Path $LocalPath)) {
        New-Item -ItemType Directory -Path $LocalPath | Out-Null
    }

    # Llamada REST para obtener carpetas y archivos
    $api = "$($SiteUrl)/_api/web/GetFolderByServerRelativeUrl('$ServerRelativeUrl')?`$expand=Folders,Files"
    $result = Invoke-PnPSteststMethod -Url $api -Method Get

    # Descargar archivos
    foreach ($file in $result.Files) {
        $fileUrl = $file.ServerRelativeUrl
        $fileName = $file.Name
        Write-Host "  📄 Archivo: $fileUrl"

        Get-PnPFile -Url $fileUrl -Path $LocalPath -FileName $fileName -AsFile -Force
    }

    # Recorrer carpetas (incluye Document Sets)
    foreach ($folder in $result.Folders) {
        if ($folder.Name -notin @("Forms")) {
            $subFolderUrl = $folder.ServerRelativeUrl
            $subFolderLocal = Join-Path $LocalPath $folder.Name

            Write-Host "📁 Carpeta / Document Set: $subFolderUrl"

            Download-FolderREST -ServerRelativeUrl $subFolderUrl -LocalPath $subFolderLocal
        }
    }
}

# ================================
# EXPORTAR TODAS LAS BIBLIOTECAS DE Documents
# ================================
$libs = Get-PnPList | Where-Object { $_.BaseTemplate -eq 101 -and $_.Hidden -eq $false }

foreach ($lib in $libs) {
    Write-Host "Descargando biblioteca: $($lib.Title)"

    $libPath = Join-Path $ExportPath $lib.Title
    $rootUrl = $lib.RootFolder.ServerRelativeUrl

    Download-FolderREST -ServerRelativeUrl $rootUrl -LocalPath $libPath
}

# ================================
# EXPORTAR LISTAS EN CSV
# ================================
$lists = Get-PnPList | Where-Object {
    $_.BaseTemplate -ne 101 -and $_.Hidden -eq $false -and $_.Title -ne "Páginas del sitio"
}

foreach ($list in $lists) {
    Write-Host "Exportando lista: $($list.Title)"

    $items = Get-PnPListItem -List $list.Title -PageSize 5000 -ErrorAction SilentlyContinue

    if ($items -eq $null -or $items.Count -eq 0) {
        Write-Host "  -> Lista vacía, se omite."
        continue
    }

    $exportable = $items | ForEach-Object {
        $_.FieldValues | Where-Object { $_.Key -notmatch "^_" }
    }

    $csvPath = Join-Path $ExportPath ($list.Title + ".csv")
    $exportable | Export-Csv -Path $csvPath -NoTypeInformation -Force
}

# ================================
# EXPORTAR PÁGINAS MODERNAS
# ================================
$pagesList = Get-PnPList | Where-Object { $_.Title -eq "Páginas del sitio" }

if ($pagesList) {
    $pagesPath = Join-Path $ExportPath "SitePages"
    New-Item -ItemType Directory -Path $pagesPath -Force | Out-Null

    Write-Host "Exportando páginas desde: $($pagesList.Title)"

    $pages = Get-PnPListItem -List $pagesList.Title -ErrorAction SilentlyContinue

    foreach ($page in $pages) {
        $jsonPath = Join-Path $pagesPath ($page.FieldValues.FileLeafRef + ".json")
        Write-Host "  -> Página: $($page.FieldValues.FileLeafRef)"
        $page.FieldValues | ConvertTo-Json -Depth 10 | Out-File $jsonPath -Encoding UTF8
    }
}

Write-Host "==============================="
Write-Host "EXPORTACIÓN COMPLETADA"
Write-Host "Carpeta: $ExportPath"
Write-Host "==============================="