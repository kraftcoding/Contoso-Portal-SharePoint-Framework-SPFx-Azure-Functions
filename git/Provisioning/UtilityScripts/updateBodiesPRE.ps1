Import-module .\_functions.ps1 -Force

# for local development will be ignored if the script is executed in devops
$global:DEV_SHPHOSTURL = "https://test.sharepoint.com";

$testBodies = @(
    @{
        "SiteRelativeUrl"    = "/sites/mapa-co-adr"
        "TemplateParameters" = @{
            "Contoso-Org_schedulers"       = "c:0t.c|tenant|718d6317-fbb4-48eb-abf3-5a4ce3266447"
            "Contoso-Org_gestorschedulers" = "c:0t.c|tenant|3d6405b1-c29d-466d-9b70-0df91581df94"
            "Contoso-Org_members"          = "c:0t.c|tenant|38982af0-65d0-4a01-90ab-bdc958f1e347"
            "Contoso-Org_guests"         = "c:0t.c|tenant|0209e18b-bc28-444d-a4cc-3b25374eefd4"
            "DefaultBodyNameTaxValue"    = "-1;#Conferencia Sectorial de Agricultura y Desarrollo Rural|ff0e1394-324a-4a85-ac22-6a82417a5b61"
        }
    },
    @{
        "SiteRelativeUrl"    = "/sites/maetd-cs-ae"
        "TemplateParameters" = @{
            "Contoso-Org_schedulers"       = "c:0t.c|tenant|a915bd17-dde5-44df-a639-461d5aa97178"
            "Contoso-Org_gestorschedulers" = "c:0t.c|tenant|dca9fc50-a603-4857-bc64-8ab259902087"
            "Contoso-Org_members"          = "c:0t.c|tenant|253e38a1-177d-459f-97f2-b5eba10f0695"
            "Contoso-Org_guests"         = "c:0t.c|tenant|93e3ca42-c964-4c21-9201-29added332b5"
            "DefaultBodyNameTaxValue"    = "-1;#Conferencia Sectorial de Agricultura y Desarrollo Rural|ff0e1394-324a-4a85-ac22-6a82417a5b61"
        }
    },
    @{
        "SiteRelativeUrl"    = "/sites/demomvp-cs-ae"
        "TemplateParameters" = @{
            "Contoso-Org_schedulers"       = "c:0t.c|tenant|3711453f-86f6-46b0-bbd8-e8db6966a92a"
            "Contoso-Org_gestorschedulers" = "c:0t.c|tenant|b4cde7cb-3ff4-4456-a678-93c9a3335b50"
            "Contoso-Org_members"          = "c:0t.c|tenant|ad521ce0-d816-41a3-bc16-a0a43bb0b175"
            "Contoso-Org_guests"         = "c:0t.c|tenant|51679630-2682-4498-99b6-3910473baeb5"
            "DefaultBodyNameTaxValue"    = "-1;#Comisión sectorial para DEMO MVP|08cd74e6-6193-4045-830d-d36819854b1f"
        }
    }
)

try{
    Set-PnPTraceLog -On -WriteToConsole -Level Debug
}catch {
    write-host "Launching on PnP.Powershell new version..."
    Write-Host $_.Exception.Message
    try{
        Start-PnPTraceLog -WriteToConsole -Level Debug
    }catch {
        write-host "Launching on PnP.Powershell OLD version..."
        Write-Host $_.Exception.Message
    }
}

foreach ($testBody in $testBodies) {
    try {
        write-host "prodvisioning site $($testBody.SiteRelativeUrl) ..."
        
        $connection = Get-PnPConnectionBySiteRelativeUrl -siteRelativeUrl $testBody.SiteRelativeUrl
        Invoke-PnPSiteTemplate -path ./../templateDepartmenttest.xml -Parameters $testBody.TemplateParameters -Connection $connection
        
        write-host "Site $($testBody.SiteRelativeUrl) prodvisioned"
    }
    catch {
        write-host "Error prodvisioning site $($testBody.SiteRelativeUrl)"
        Write-Host $_.Exception.Message
    }
}
