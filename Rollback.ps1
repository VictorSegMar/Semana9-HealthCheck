# Rollback.ps1 - Script de reversion para HealthCheck
Write-Host "Iniciando proceso de Rollback..." -ForegroundColor Yellow

# Elimina archivos temporales o reportes generados en la carpeta (ajusta si usas otra ruta)
Get-ChildItem -Path . -Include *.html, *.csv, *.log -Recurse | Remove-Item -Force -ErrorAction SilentlyContinue

Write-Host "Rollback completado: entorno restaurado al estado inicial." -ForegroundColor Green
