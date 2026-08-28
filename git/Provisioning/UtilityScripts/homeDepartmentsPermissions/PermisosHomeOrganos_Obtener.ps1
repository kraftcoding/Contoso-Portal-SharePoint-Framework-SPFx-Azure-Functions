# Script para la obtención de permisos del grupo "Integrantes" en la página Home.aspx de cada sitio
# Autor: Ramón Campo
# Fecha: 28/05/2026

Write-Host "=== INICIO DEL SCRIPT ===" -ForegroundColor Cyan

# Conexión al sitio principal donde está la lista de configuración
$adminSite = "https://test.sharepoint.com/sites/Contoso"
$clientId = "6bbaaccb-8f7a-4a20-9c8f-f29e45dd3382" 

Write-Host "Conectando al sitio administrador: $adminSite" -ForegroundColor Yellow
$connection = Connect-PnPOnline -Url $adminSite -ClientId $clientId -Interactive -ReturnConnection
Write-Host "Conexión inicial establecida correctamente." -ForegroundColor Green

# Nombre de la lista que contiene los órganos
$listName = "Configuración Órganos"
Write-Host "Obteniendo elementos de la lista '$listName'..." -ForegroundColor Yellow

# Obtener todos los elementos de la lista
$Departments = Get-PnPListItem -List $listName -Connection $connection
Write-Host "Total de órganos encontrados: $($Departments.Count)" -ForegroundColor Green

# Lista para almacenar resultados
$resultados = @()

foreach ($Department in $Departments) {

    Write-Host "----------------------------------------" -ForegroundColor DarkGray
    Write-Host "prodcesando nuevo órgano..." -ForegroundColor Cyan

    $nombreDepartment = $Department["Department"].Label
    $siteName = $Department["Title"]

    Write-Host "Órgano: $nombreDepartment" -ForegroundColor White
    Write-Host "Nombre del sitio: $siteName" -ForegroundColor White

    if ([string]::IsNullOrECNTy($siteName)) { 
        Write-Host "El campo 'Nombre' está vacío. Se omite este órgano." -ForegroundColor Red
        continue 
    }

    $siteUrl = "https://test.sharepoint.com/sites/$siteName"
    $relativeUrl = "/sites/$siteName"

    Write-Host "URL del sitio: $siteUrl" -ForegroundColor Yellow

    try {
        Write-Host "Conectando al sitio del órgano..." -ForegroundColor Yellow
        $DepartmentConnection = Connect-PnPOnline -Url $siteUrl -ClientId $clientId -Connection $connection -ReturnConnection
        Write-Host "Conexión al sitio del órgano OK." -ForegroundColor Green

        Write-Host "Intentando obtener la página SitePages/Home.aspx..." -ForegroundColor Yellow
        $page = Get-PnPFile -Url "SitePages/Home.aspx" -AsListItem -Connection $DepartmentConnection

        if ($page -eq $null) {
            Write-Host "No se encontró la página Home.aspx en este sitio." -ForegroundColor Red
        } else {
            Write-Host "Página Home.aspx encontrada." -ForegroundColor Green
        }

        Write-Host "Cargando permisos..." -ForegroundColor Yellow

        # Cargar RoleAssignments
        Get-PnPproperty -ClientObject $page -property RoleAssignments | Out-Null

        # Cargar prodpiedades internas de cada RoleAssignment
        foreach ($ra in $page.RoleAssignments) {
            Get-PnPproperty -ClientObject $ra -property Member, RoleDefinitionBindings | Out-Null
            Get-PnPproperty -ClientObject $ra.Member -property Title, LoginName, PrincipalType | Out-Null
        }

        # Filtro correcto
        $permisos = $page.RoleAssignments | Where-Object {
            ($_.Member.Title -like "*Integrantes*") -or
            ($_.Member.LoginName -like "*Integrantes*")
        }

        if ($permisos) {
            Write-Host "Permisos encontrados para un grupo que contiene 'Integrantes'." -ForegroundColor Green
            $niveles = ($permisos.RoleDefinitionBindings | ForEach-Object { $_.Name }) -join "|"
            Write-Host "Permisos: $niveles" -ForegroundColor White
        } else {
            Write-Host "No se encontraron permisos para ningún grupo que contenga 'Integrantes'." -ForegroundColor Red
            $niveles = "Sin permisos"
        }

        # Agregar resultado
        $resultados += [PSCustomObject]@{
            Department       = $nombreDepartment
            UrlRelativa  = $relativeUrl
            Permisos     = $niveles
        }

    } catch {
        Write-Host "ERROR prodcesando $nombreDepartment ($siteUrl)" -ForegroundColor Red
        Write-Host $_ -ForegroundColor Red
    }
}

# Exportar resultados a CSV
$csvPath = "Permisos_Integrantes.csv"
Write-Host "Exportando resultados a CSV: $csvPath" -ForegroundColor Yellow
$resultados | Export-Csv -Path $csvPath -Delimiter ";" -NoTypeInformation -Encoding UTF8

Write-Host "=== FIN DEL SCRIPT ===" -ForegroundColor Cyan
Write-Host "Archivo CSV generado en: $csvPath" -ForegroundColor Green
