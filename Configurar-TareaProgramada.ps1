########################################################################
# CONFIGURADOR DE TAREA PROGRAMADA - HEALTH CHECK ENTERPRISE
# Autor: Victor Manuel Segundo Marquez
#
# Instrucciones:
# Cambia el valor de $HoraEjecucion (formato 24h "HH:mm") y ejecuta
# este script como Administrador para actualizar el horario.
########################################################################

$HoraEjecucion = "15:50"   # <-- EDITA AQUÍ LA HORA DESEADA (Ej. "08:00", "14:30", "18:00")
$NombreTarea   = "HealthCheck_Enterprise_Diario"
$RutaScript    = "C:\Semana9-Automatizacion\Scripts\OFICIALES\Run-HealthCheckSuite.ps1"

# 1. Validar que se ejecuta como Administrador
$EsAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $EsAdmin) {
    Write-Host "[ERROR] Debes ejecutar este script en PowerShell como Administrador." -ForegroundColor Red
    return
}

# 2. Definir acción, nuevo horario y permisos
$Accion        = New-ScheduledTaskAction -Execute "powershell.exe" -Argument "-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$RutaScript`""
$Disparador    = New-ScheduledTaskTrigger -Daily -At $HoraEjecucion
$Principal     = New-ScheduledTaskPrincipal -UserId "$env:USERDOMAIN\$env:USERNAME" -LogonType S4U -RunLevel Highest
$Configuracion = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable

# 3. Registrar o actualizar la tarea existente
Register-ScheduledTask `
    -TaskName $NombreTarea `
    -Description "Ejecuta diariamente HealthCheck.ps1 y genera HealthCheckDashboard.xlsx a las $HoraEjecucion hrs" `
    -Action $Accion `
    -Trigger $Disparador `
    -Principal $Principal `
    -Settings $Configuracion `
    -Force | Out-Null

$Info = Get-ScheduledTask -TaskName $NombreTarea | Get-ScheduledTaskInfo
Write-Host "[OK] Tarea '$NombreTarea' actualizada correctamente." -ForegroundColor Green
Write-Host "Próxima ejecución programada: $($Info.NextRunTime)" -ForegroundColor Cyan
