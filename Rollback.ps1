########################################################################
# ROLLBACK & RECUPERACIÓN ENTERPRISE - HEALTH CHECK
# Autor: Victor Manuel Segundo Marquez
#
# Objetivo General:
# Proveer un mecanismo seguro, auditado y reversible para limpiar,
# corregir o restaurar los reportes generados por el Health Check,
# evitando pérdida accidental de información histórica.
########################################################################

########################################################################
# MÓDULO 1: VARIABLES GLOBALES Y PARÁMETROS DE CONFIGURACIÓN
# Propósito: Centralizar las rutas de trabajo, el formato de marca de
# tiempo y la política máxima de retención de respaldos en disco.
########################################################################

$RutaReportes = "C:\Semana9-Automatizacion\Reportes"          # Carpeta donde se generan CSV, XLSX y Logs
$RutaBackups  = "C:\Semana9-Automatizacion\Backups_Rollback"  # Carpeta destino para respaldos comprimidos (.zip)
$MaxBackups   = 10                                            # Número máximo de respaldos ZIP a conservar

########################################################################
# MÓDULO 2: FUNCIONES AUXILIARES DE AUDITORÍA, VALIDACIÓN Y RESPALDO
# Propósito: Encapsular tareas críticas reutilizables (bitácora, control
# de archivos abiertos en Excel y creación de paquetes ZIP).
########################################################################

# Función: Write-RollbackLog
# Registra cada operación de reversión o restauración en el archivo de log oficial.
function Write-RollbackLog {
    param([string]$Mensaje)
    
    # Si la carpeta de reportes fue eliminada o no existe, la crea para guardar el log
    if (-not (Test-Path $RutaReportes)) { 
        New-Item -ItemType Directory -Path $RutaReportes -Force | Out-Null 
    }
    
    $LogPath = Join-Path $RutaReportes "HealthCheck_Log.txt"
    "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') [ROLLBACK] $Mensaje" | Add-Content -Path $LogPath -Encoding utf8
}

# Función: Test-ArchivoBloqueado
# Intenta abrir un archivo en modo exclusivo; si falla, significa que está abierto (ej. en Excel).
function Test-ArchivoBloqueado {
    param([string]$RutaArchivo)
    
    if (-not (Test-Path $RutaArchivo)) { return $false }
    try {
        $stream = [System.IO.File]::Open($RutaArchivo, 'Open', 'ReadWrite', 'None')
        $stream.Close()
        return $false # Archivo libre
    } catch {
        return $true  # Archivo bloqueado por otro proceso
    }
}

# Función: Invoke-RespaldoPreventivo
# Empaqueta todo el contenido actual de Reportes en un ZIP antes de borrar o modificar datos,
# y aplica la política de retención eliminando respaldos excedentes antiguos.
function Invoke-RespaldoPreventivo {
    # Crea el directorio de backups si es la primera vez que se ejecuta
    if (-not (Test-Path $RutaBackups)) {
        New-Item -ItemType Directory -Path $RutaBackups -Force | Out-Null
    }
    
    # Verifica si existen archivos para respaldar
    $Archivos = Get-ChildItem -Path $RutaReportes -File -ErrorAction SilentlyContinue
    if (-not $Archivos -or $Archivos.Count -eq 0) { return $null }

        # Genera la marca de tiempo exacta al momento de respaldar
    $FechaCorte = Get-Date -Format "yyyyMMdd_HHmmss"
    $ArchivoZip = Join-Path $RutaBackups "Backup_PreRollback_$FechaCorte.zip"
    Compress-Archive -Path "$RutaReportes\*" -DestinationPath $ArchivoZip -Force

    # Política de retención: ordena del más reciente al más antiguo y borra lo que supere $MaxBackups
    Get-ChildItem -Path $RutaBackups -Filter "Backup_PreRollback_*.zip" |
        Sort-Object LastWriteTime -Descending |
        Select-Object -Skip $MaxBackups |
        Remove-Item -Force -ErrorAction SilentlyContinue

    return $ArchivoZip
}

# Función: Pause-VolverAlMenu
# Detiene la pantalla para que el usuario lea el resultado antes de regresar al menú principal.
function Pause-VolverAlMenu {
    Write-Host ""
    Read-Host "Presiona [ENTER] para volver al menú principal" | Out-Null
}



########################################################################
# MÓDULO 3: BUCLE DE NAVEGACIÓN, INTERFAZ Y VALIDACIÓN DE BLOQUEOS
# Propósito: Mantener el menú activo hasta que el usuario elija salir,
# permitiendo cancelar cualquier acción y regresar al inicio.
########################################################################

$Continuar = $true

while ($Continuar) {
    Clear-Host
    Write-Host "==========================================================" -ForegroundColor Cyan
    Write-Host "     SISTEMA DE ROLLBACK & RECUPERACIÓN - HEALTH CHECK    " -ForegroundColor Cyan
    Write-Host "==========================================================" -ForegroundColor Cyan
    Write-Host "1. Limpieza total (Crea Backup ZIP y vacía carpeta Reportes)"
    Write-Host "2. Deshacer última ejecución (Elimina último registro del Histórico)"
    Write-Host "3. Restaurar reportes desde un Backup ZIP existente"
    Write-Host "4. Salir"
    Write-Host "==========================================================" -ForegroundColor Cyan

    $Opcion = Read-Host "Selecciona una opción (1-4)"

    # Verifica que HealthCheck_Dashboard.xlsx no esté bloqueado antes de opciones 1 o 3
    $ExcelPath = Join-Path $RutaReportes "HealthCheck_Dashboard.xlsx"
    if ($Opcion -in @("1", "3") -and (Test-ArchivoBloqueado -RutaArchivo $ExcelPath)) {
        Write-Host "`n[ERROR] El archivo 'HealthCheck_Dashboard.xlsx' está abierto en Excel." -ForegroundColor Red
        Write-Host "Ciérralo primero antes de continuar." -ForegroundColor Yellow
        Pause-VolverAlMenu
        continue
    }



    ####################################################################
    # MÓDULO 4: EJECUCIÓN DE MODOS DE RECUPERACIÓN CON CONFIRMACIÓN
    # Propósito: Solicitar confirmación explícita en cada operación
    # permitiendo cancelar y volver al menú principal en todo momento.
    ####################################################################

    switch ($Opcion) {

        # --------------------------------------------------------------
        # OPCIÓN 1: LIMPIEZA TOTAL CON CONFIRMACIÓN PREVIA
        # --------------------------------------------------------------
        "1" {
            if (-not (Test-Path $RutaReportes)) {
                Write-Host "`n[INFO] No existe la carpeta de reportes. Nada que limpiar." -ForegroundColor Yellow
                Pause-VolverAlMenu
                continue
            }

            $Archivos = @(Get-ChildItem -Path $RutaReportes -File)
            if ($Archivos.Count -eq 0) {
                Write-Host "`n[INFO] La carpeta de reportes ya está vacía." -ForegroundColor Green
                Pause-VolverAlMenu
                continue
            }

            Write-Host "`n--- CONFIRMACIÓN DE LIMPIEZA TOTAL ---" -ForegroundColor Yellow
            Write-Host "Se respaldarán en ZIP y luego se eliminarán $($Archivos.Count) archivo(s) de Reportes."
            $Confirmar = Read-Host "¿Deseas continuar? (S = Sí, ejecutar / C = Cancelar y volver al menú)"

            if ($Confirmar -notin @("S", "s", "SI", "si", "Sí", "sí")) {
                Write-Host "`n[CANCELADO] Operación abortada. Volviendo al menú principal..." -ForegroundColor Gray
                Start-Sleep -Seconds 1
                continue
            }

            $Zip = Invoke-RespaldoPreventivo
            Write-Host "`n[OK] Respaldo creado en: $Zip" -ForegroundColor Green
            
            foreach ($Arch in $Archivos) {
                Remove-Item -Path $Arch.FullName -Force
                Write-Host "  -> Eliminado: $($Arch.Name)" -ForegroundColor Gray
            }

            Write-RollbackLog "Limpieza total ejecutada ($($Archivos.Count) archivos). Respaldo: $Zip"
            Write-Host "`n[COMPLETADO] Entorno limpio y listo para nueva ejecución." -ForegroundColor Cyan
            Pause-VolverAlMenu
        }

        # --------------------------------------------------------------
        # OPCIÓN 2: ROLLBACK PARCIAL CON CONFIRMACIÓN DEL REGISTRO
        # --------------------------------------------------------------
        "2" {
            $HistPath = Join-Path $RutaReportes "HealthCheck_History.csv"
            if (-not (Test-Path $HistPath)) {
                Write-Host "`n[ERROR] No se encontró el archivo HealthCheck_History.csv." -ForegroundColor Red
                Pause-VolverAlMenu
                continue
            }

            $Historico = @(Import-Csv -Path $HistPath -Encoding UTF8)
            if ($Historico.Count -le 1) {
                Write-Host "`n[ADVERTENCIA] El histórico solo tiene $($Historico.Count) registro. Usa la Opción 1 para limpieza total." -ForegroundColor Yellow
                Pause-VolverAlMenu
                continue
            }

            $UltimoRegistro = $Historico[-1]
            Write-Host "`n--- CONFIRMACIÓN DE ROLLBACK PARCIAL ---" -ForegroundColor Yellow
            Write-Host "Último registro detectado -> Fecha: $($UltimoRegistro.Fecha) | Score: $($UltimoRegistro.Score)"
            $Confirmar = Read-Host "¿Deseas eliminar este registro? (S = Sí, eliminar / C = Cancelar y volver al menú)"

            if ($Confirmar -notin @("S", "s", "SI", "si", "Sí", "sí")) {
                Write-Host "`n[CANCELADO] No se modificó el histórico. Volviendo al menú principal..." -ForegroundColor Gray
                Start-Sleep -Seconds 1
                continue
            }

            $Zip = Invoke-RespaldoPreventivo
            $HistoricoRecortado = $Historico[0..($Historico.Count - 2)]
            $HistoricoRecortado | Export-Csv -Path $HistPath -NoTypeInformation -Encoding UTF8

            Write-RollbackLog "Se revirtió el último registro del histórico (Fecha: $($UltimoRegistro.Fecha)). Respaldo: $Zip"
            Write-Host "`n[OK] Se eliminó del histórico la ejecución del: $($UltimoRegistro.Fecha)" -ForegroundColor Green
            Write-Host "Registros restantes en histórico: $($HistoricoRecortado.Count)" -ForegroundColor Cyan
            Write-Host "Nota: Ejecuta .\HealthCheckDashboard.ps1 para regenerar el Excel con este cambio." -ForegroundColor Yellow
            Pause-VolverAlMenu
        }

        # --------------------------------------------------------------
        # OPCIÓN 3: RESTAURACIÓN DESDE BACKUP ZIP CON OPCIÓN DE VOLVER
        # --------------------------------------------------------------
        "3" {
            if (-not (Test-Path $RutaBackups)) {
                Write-Host "`n[ERROR] No existe la carpeta de respaldos ($RutaBackups)." -ForegroundColor Red
                Pause-VolverAlMenu
                continue
            }

            $Backups = @(Get-ChildItem -Path $RutaBackups -Filter "Backup_PreRollback_*.zip" | Sort-Object LastWriteTime -Descending)
            if ($Backups.Count -eq 0) {
                Write-Host "`n[INFO] No hay archivos de respaldo disponibles para restaurar." -ForegroundColor Yellow
                Pause-VolverAlMenu
                continue
            }

            Write-Host "`nRespaldos disponibles:" -ForegroundColor Cyan
            Write-Host "  [0] <-- Cancelar y volver al menú principal" -ForegroundColor Yellow
            for ($i = 0; $i -lt $Backups.Count; $i++) {
                Write-Host "  [$($i + 1)] $($Backups[$i].Name)  ($(Get-Date $Backups[$i].LastWriteTime -Format 'dd/MM/yyyy HH:mm:ss'))"
            }

            $Sel = Read-Host "`nIngresa el número del respaldo a restaurar (0 para cancelar, 1-$($Backups.Count))"
            
            if ($Sel -eq "0") {
                Write-Host "`n[CANCELADO] Volviendo al menú principal..." -ForegroundColor Gray
                Start-Sleep -Seconds 1
                continue
            }

            $Idx = 0
            if (-not [int]::TryParse($Sel, [ref]$Idx) -or $Idx -lt 1 -or $Idx -gt $Backups.Count) {
                Write-Host "`n[ERROR] Selección inválida." -ForegroundColor Red
                Pause-VolverAlMenu
                continue
            }

            $ZipElegido = $Backups[$Idx - 1]
            $Confirmar = Read-Host "¿Confirmas reemplazar los reportes actuales con '$($ZipElegido.Name)'? (S = Sí / C = Cancelar)"
            
            if ($Confirmar -notin @("S", "s", "SI", "si", "Sí", "sí")) {
                Write-Host "`n[CANCELADO] Restauración abortada. Volviendo al menú principal..." -ForegroundColor Gray
                Start-Sleep -Seconds 1
                continue
            }

            if (-not (Test-Path $RutaReportes)) { New-Item -ItemType Directory -Path $RutaReportes -Force | Out-Null }
            Get-ChildItem -Path $RutaReportes -File | Remove-Item -Force -ErrorAction SilentlyContinue
            Expand-Archive -Path $ZipElegido.FullName -DestinationPath $RutaReportes -Force

            Write-RollbackLog "Entorno restaurado desde el respaldo: $($ZipElegido.Name)"
            Write-Host "`n[OK] Reportes restaurados exitosamente desde $($ZipElegido.Name)." -ForegroundColor Green
            Pause-VolverAlMenu
        }

        # --------------------------------------------------------------
        # OPCIÓN 4: SALIR DEL SCRIPT
        # --------------------------------------------------------------
        "4" {
            Write-Host "`nSaliendo del sistema de Rollback. ¡Hasta pronto!" -ForegroundColor Cyan
            $Continuar = $false
        }

        # --------------------------------------------------------------
        # OPCIÓN INVÁLIDA: AVISO Y REGRESO AL MENÚ
        # --------------------------------------------------------------
        default {
            Write-Host "`n[ADVERTENCIA] Opción no válida. Elige un número del 1 al 4." -ForegroundColor Yellow
            Start-Sleep -Seconds 1
        }
    }
}
