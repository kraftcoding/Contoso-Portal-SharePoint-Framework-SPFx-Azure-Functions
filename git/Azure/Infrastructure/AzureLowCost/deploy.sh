#!/bin/bash

# ================================
# CONFIGURACIÓN
# ================================
RG_NAME="rg-dev-contoso"
LOCATION="westeurope"
TEMPLATE_FILE="infra/main.bicep"
PARAMS_FILE="infra/params/dev.json"

# ================================
# MENÚ DE ACCIONES
# ================================
echo "======================================"
echo "   SCRIPT DE DESPLIEGUE DE INFRA"
echo "======================================"
echo ""
echo "1) Crear/actualizar entorno"
echo "2) Eliminar entorno y recrear"
echo "3) Salir"
echo ""
read -p "Selecciona una opción: " OPTION

# ================================
# FUNCIÓN: CREAR RESOURCE GROUP
# ================================
create_rg() {
    echo "🔍 Comprodbando si existe el Resource Group '$RG_NAME'..."
    RG_EXISTS=$(az group exists --name $RG_NAME)

    if [ "$RG_EXISTS" = "false" ]; then
        echo "📦 Creando Resource Group '$RG_NAME'..."
        az group create --name $RG_NAME --location $LOCATION
    else
        echo "✔️ El Resource Group ya existe"
    fi
}

# ================================
# FUNCIÓN: ELIMINAR RESOURCE GROUP
# ================================
delete_rg() {
    echo "⚠️ Eliminando Resource Group '$RG_NAME'..."
    az group delete --name $RG_NAME --yes --no-wait
    echo "⏳ Eliminación iniciada. Espera unos minutos antes de recrear."
}

# ================================
# FUNCIÓN: DESPLEGAR INFRA
# ================================
deploy_infra() {
    echo "🚀 Iniciando despliegue..."
    az deployment group create \
        --resource-group $RG_NAME \
        --template-file $TEMPLATE_FILE \
        --parameters @$PARAMS_FILE

    echo "🎉 Despliegue completado"
}

# ================================
# LÓGICA PRINCIPAL
# ================================
case $OPTION in

    1)
        create_rg
        deploy_infra
        ;;

    2)
        delete_rg
        echo "⏳ Espera a que el RG se elimine y vuelve a ejecutar el script"
        ;;

    3)
        echo "👋 Saliendo..."
        exit 0
        ;;

    *)
        echo "❌ Opción no válida"
        ;;
esac