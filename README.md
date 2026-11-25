# Semana 9 - Automatizacion HealthCheck

Repositorio con los scripts de monitoreo, dashboard y procedimiento de rollback.

## Archivos del repositorio
* **HealthCheck.ps1**: Ejecuta el diagnostico principal del sistema.
* **HealthCheckDashboard.ps1**: Genera la vista o reporte visual de los resultados.
* **Rollback.ps1**: Revierte los cambios y limpia los archivos generados en caso de fallo o reinicio.

## Como ejecutarlos
1. Abrir **PowerShell** en la carpeta del proyecto.
2. Ejecutar el script principal:
   \\\powershell
   .\HealthCheck.ps1
   \\\
3. Ejecutar el dashboard:
   \\\powershell
   .\HealthCheckDashboard.ps1
   \\\

## Procedimiento de Rollback
* **Mediante script:** Ejecutar \.\Rollback.ps1\ para limpiar los archivos generados.
* **Mediante Git:** Ejecutar \git log --oneline\ para ver el historial y \git revert HEAD\ para deshacer el ultimo cambio guardado.
