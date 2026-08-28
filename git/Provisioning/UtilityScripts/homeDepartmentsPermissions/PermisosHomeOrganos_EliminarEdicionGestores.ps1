
# Script para la eliminación del permiso Edit del grupo "Integrantes" en la página Home.aspx de cada sitio de Órgano.
# Autor Ramón Campo
# Fecha: 29/05/2026

Write-Host "=== INICIO SCRIPT ELIMINAR PERMISO EDITAR ===" -ForegroundColor Cyan

# Parámetros
$adminSite = "https://test.sharepoint.com/sites/Contoso"
$clientId  = "6bbaaccb-8f7a-4a20-9c8f-f29e45dd3382"
$csvPath   = "EliminarPermisoEditar.csv"
$logPath   = "Log_EliminarPermisoEditar_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"

# Conexión admin
try {
    Write-Host "[INFO] Conectando a Admin..." -ForegroundColor Yellow

    $connection = Connect-PnPOnline `
        -Url $adminSite `
        -ClientId $clientId `
        -Interactive `
        -ReturnConnection `
        -ErrorAction Stop

    Write-Host "[OK] Conectado a Admin" -ForegroundColor Green
}
catch {
    Write-Host "[ERROR] $($_.Exception.Message)" -ForegroundColor Red
    exit
}

# Cargar CSV
$sites = Import-Csv -Path $csvPath -Delimiter ";"

$resultados = @()

foreach ($site in $sites) {

    Write-Host "----------------------------------------" -ForegroundColor DarkGray

    $Department = $site.Department
    $siteUrl = "https://test.sharepoint.com$($site.UrlRelativa)"

    Write-Host "[INFO] Órgano: $Department" -ForegroundColor Cyan
    Write-Host "[INFO] URL: $siteUrl" -ForegroundColor Gray

    $estado = "OK"
    $detalle = ""

    try {

        # Conexión site
        $conn = Connect-PnPOnline `
            -Url $siteUrl `
            -ClientId $clientId `
            -Connection $connection `
            -ReturnConnection `
            -ErrorAction Stop

        Write-Host "[OK] Conectado al site" -ForegroundColor Green

        # Obtener página
        $page = Get-PnPFile `
            -Url "SitePages/Home.aspx" `
            -AsListItem `
            -Connection $conn `
            -ErrorAction Stop

        Write-Host "[OK] Página ID: $($page.Id)" -ForegroundColor Green

        # Cargar permisos
        Get-PnPproperty -ClientObject $page -property RoleAssignments -Connection $conn | Out-Null

        # Filtrar SOLO grupos que contengan "Integrantes"
        $grupoRA = @()

        foreach ($ra in $page.RoleAssignments) {

            Get-PnPproperty -ClientObject $ra -property Member, RoleDefinitionBindings -Connection $conn | Out-Null
            Get-PnPproperty -ClientObject $ra.Member -property Title, LoginName, PrincipalType -Connection $conn | Out-Null

            # 👇 SOLO grupos (SharePointGroup)
            if ($ra.Member.PrincipalType -eq "SharePointGroup") {

                if ($ra.Member.Title -like "*Integrantes*") {
                    $grupoRA += $ra
                }
            }
        }

        if ($grupoRA.Count -eq 0) {
            $estado = "SIN_GRUPO"
            $detalle = "No hay grupos Integrantes"
            Write-Host "[WARN] $detalle" -ForegroundColor Yellow
        }
        else {

            Write-Host "[OK] Grupos encontrados: $($grupoRA.Count)" -ForegroundColor Green

            foreach ($ra in $grupoRA) {

                $groupName = $ra.Member.Title
                $roles = $ra.RoleDefinitionBindings | ForEach-Object { $_.Name }

                Write-Host "[INFO] Grupo: $groupName" -ForegroundColor Cyan
                Write-Host "[INFO] Permisos: $($roles -join ', ')" -ForegroundColor Yellow

                if ($roles -contains "Editar") {

                    Write-Host "[INFO] Eliminando permiso Editar..." -ForegroundColor DarkGray

                    # Quitar solo "Editar"
                    Set-PnPListItemPermission `
                        -List "SitePages" `
                        -Identity $page.Id `
                        -Group $groupName `
                        -RemoveRole "Editar" `
                        -Connection $conn `
                        -ErrorAction Stop

                    $detalle = "Editar eliminado de grupo $groupName"
                    Write-Host "[OK] $detalle" -ForegroundColor Green
                }
                else {
                    $estado = "SIN_EDITAR"
                    $detalle = "Grupo sin permiso Editar"
                    Write-Host "[INFO] $detalle" -ForegroundColor DarkYellow
                }
            }
        }

    }
    catch {
        $estado = "ERROR"
        $detalle = $_.Exception.Message
        Write-Host "[ERROR] $detalle" -ForegroundColor Red
    }

    $resultados += [PSCustomObject]@{
        Fecha  = Get-Date
        Department = $Department
        Url    = $siteUrl
        Estado = $estado
        Detalle= $detalle
    }
}

# Export log
$resultados | Export-Csv `
    -Path $logPath `
    -Delimiter ";" `
    -NoTypeInformation `
    -Encoding UTF8

Write-Host "=== FIN SCRIPT ===" -ForegroundColor Cyan
Write-Host "[OK] Log: $logPath" -ForegroundColor Green