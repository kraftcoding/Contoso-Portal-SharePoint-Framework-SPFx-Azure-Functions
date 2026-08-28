# Conexión al sitio
Connect-PnPOnline `
    -Url "https://contoso.sharepoint.com/sites/CNTmd-cs-Contoso" `
    -ClientId "00000000-0000-0000-0001-000000000001" `
    -Tenant "contoso.onmicrosoft.com" `
    -Thumbprint "AAAA1111BBBB2222CCCC3333DDDD4444EEEE5555"

# Listas del sistema que NO deben tocarse
$systemLists = @(
    "Style Library",
    "Site Assets",
    "Site Pages",
    "Form Templates",
    "Master Page Gallery",
    "Theme Gallery",
    "Solution Gallery",
    "TaxonomyHiddenList",
    "User Information List",
    "Workflow Tasks",
    "Workflow History",
    "Composed Looks",
    "List Template Gallery",
    "Web Part Gallery"
)

# Obtener todas las listas y bibliotecas del sitio
$lists = Get-PnPList

foreach ($list in $lists) {

    # Saltar listas del sistema
    if ($systemLists -contains $list.Title) {
        Write-Host "⏭ Saltando lista del sistema: $($list.Title)"
        continue
    }

    Write-Host "🔍 Revisando: $($list.Title)"

    # Si la lista NO hereda permisos, restaurarlos
    if (-not $list.HasUniqueRoleAssignments) {
        Write-Host "✔ $($list.Title) ya hereda permisos"
    }
    else {
        Write-Host "🔧 Restaurando permisos heredados en: $($list.Title)"
        Set-PnPList -Identity $list -ResetRoleInheritance
        Write-Host "✅ Permisos heredados restaurados en: $($list.Title)"
    }
}

Write-Host "---------------------------------------------"
Write-Host "🏁 prodceso completado. Todas las listas y bibliotecas están desbloqueadas."
Write-Host "---------------------------------------------"