########################################################################
# ORQUESTADOR DIARIO - HEALTH CHECK & DASHBOARD EJECUTIVO
# Autor: Victor Manuel Segundo Marquez
########################################################################

$RutaScripts = "C:\Semana9-Automatizacion\Scripts\OFICIALES"
Set-Location -Path $RutaScripts

# 1. Ejecutar diagnóstico del sistema
& "$RutaScripts\HealthCheck.ps1"

# 2. Generar Dashboard Ejecutivo en Excel con los datos recién creados
& "$RutaScripts\HealthCheckDashboard.ps1"
