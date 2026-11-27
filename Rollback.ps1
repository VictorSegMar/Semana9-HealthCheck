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
$FechaCorte   = Get-Date -Format "yyyyMMdd_HHmmss"            # Marca de tiempo única para nombrar cada respaldo
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

    # Comprime todos los archivos actuales en un paquete ZIP con fecha y hora
    $ArchivoZip = Join-Path $RutaBackups "Backup_PreRollback_$FechaCorte.zip"
    Compress-Archive -Path "$RutaReportes\*" -DestinationPath $ArchivoZip -Force

    # Política de retención: ordena del más reciente al más antiguo y borra lo que supere $MaxBackups
    Get-ChildItem -Path $RutaBackups -Filter "Backup_PreRollback_*.zip" |
        Sort-Object LastWriteTime -Descending |
        Select-Object -Skip $MaxBackups |
        Remove-Item -Force -ErrorAction SilentlyContinue

    return $ArchivoZip
}


########################################################################
# MÓDULO 3: INTERFAZ DE USUARIO Y VALIDACIÓN PREVENTIVA DE BLOQUEOS
# Propósito: Mostrar las opciones disponibles al operador y frenar el
# proceso si el archivo Excel está abierto para evitar errores parciales.
########################################################################

Clear-Host
Write-Host "==========================================================" -ForegroundColor Cyan
Write-Host "     SISTEMA DE ROLLBACK & RECUPERACIÓN - HEALTH CHECK    " -ForegroundColor Cyan
Write-Host "==========================================================" -ForegroundColor Cyan
Write-Host "1. Limpieza total (Crea Backup ZIP y vacía carpeta Reportes)"
Write-Host "2. Deshacer última ejecución (Elimina último registro del Histórico)"
Write-Host "3. Restaurar reportes desde un Backup ZIP existente"
Write-Host "4. Cancelar / Salir"
Write-Host "==========================================================" -ForegroundColor Cyan

$Opcion = Read-Host "Selecciona una opción (1-4)"

# Verifica que HealthCheck_Dashboard.xlsx no esté bloqueado antes de ejecutar opciones 1 o 3
$ExcelPath = Join-Path $RutaReportes "HealthCheck_Dashboard.xlsx"
if ($Opcion -in @("1", "3") -and (Test-ArchivoBloqueado -RutaArchivo $ExcelPath)) {
    Write-Host "`n[ERROR] El archivo 'HealthCheck_Dashboard.xlsx' está abierto en Excel." -ForegroundColor Red
    Write-Host "Ciérralo primero y vuelve a ejecutar el Rollback." -ForegroundColor Yellow
    return
}

########################################################################
# MÓDULO 4: EJECUCIÓN DE MODOS DE RECUPERACIÓN (SWITCH PRINCIPAL)
# Propósito: Ejecutar la lógica específica según la opción elegida.
########################################################################

switch ($Opcion) {

    # ------------------------------------------------------------------
    # OPCIÓN 1: LIMPIEZA TOTAL CON RESPALDO PREVENTIVO
    # Respalda toda la carpeta Reportes en ZIP y elimina los archivos activos.
    # ------------------------------------------------------------------
    "1" {
        if (-not (Test-Path $RutaReportes)) {
            Write-Host "`n[INFO] No existe la carpeta de reportes. Nada que limpiar." -ForegroundColor Yellow
            return
        }

        $Zip = Invoke-RespaldoPreventivo
        if (-not $Zip) {
            Write-Host "`n[INFO] La carpeta de reportes ya está vacía." -ForegroundColor Green
            return
        }

        Write-Host "`n[OK] Respaldo creado en: $Zip" -ForegroundColor Green
        
        # Recorre y elimina cada archivo del directorio de reportes
        $Archivos = Get-ChildItem -Path $RutaReportes -File
        foreach ($Arch in $Archivos) {
            Remove-Item -Path $Arch.FullName -Force
            Write-Host "  -> Eliminado: $($Arch.Name)" -ForegroundColor Gray
        }

        Write-RollbackLog "Limpieza total ejecutada ($($Archivos.Count) archivos). Respaldo: $Zip"
        Write-Host "`n[COMPLETADO] Entorno limpio y listo para nueva ejecución." -ForegroundColor Cyan
    }

    # ------------------------------------------------------------------
    # OPCIÓN 2: ROLLBACK PARCIAL (ÚLTIMA EJECUCIÓN DEL HISTÓRICO)
    # Conserva las ejecuciones anteriores y elimina solo la última fila
    # agregada a HealthCheck_History.csv.
    # ------------------------------------------------------------------
    "2" {
        $HistPath = Join-Path $RutaReportes "HealthCheck_History.csv"
        if (-not (Test-Path $HistPath)) {
            Write-Host "`n[ERROR] No se encontró el archivo HealthCheck_History.csv." -ForegroundColor Red
            return
        }

        # Carga el histórico como arreglo de objetos para manipular sus filas
        $Historico = @(Import-Csv -Path $HistPath -Encoding UTF8)
        if ($Historico.Count -le 1) {
            Write-Host "`n[ADVERTENCIA] El histórico solo tiene $($Historico.Count) registro. Usa la Opción 1 para limpieza total." -ForegroundColor Yellow
            return
        }

        # Respalda antes de recortar el CSV
        $Zip = Invoke-RespaldoPreventivo
        $UltimoRegistro = $Historico[-1]
        
        # Conserva desde el primer registro (índice 0) hasta el penúltimo (Count - 2)
        $HistoricoRecortado = $Historico[0..($Historico.Count - 2)]
        $HistoricoRecortado | Export-Csv -Path $HistPath -NoTypeInformation -Encoding UTF8

        Write-RollbackLog "Se revirtió el último registro del histórico (Fecha: $($UltimoRegistro.Fecha)). Respaldo: $Zip"
        Write-Host "`n[OK] Se eliminó del histórico la ejecución del: $($UltimoRegistro.Fecha)" -ForegroundColor Green
        Write-Host "Registros restantes en histórico: $($HistoricoRecortado.Count)" -ForegroundColor Cyan
        Write-Host "Nota: Ejecuta .\HealthCheckDashboard.ps1 para regenerar el Excel con este cambio." -ForegroundColor Yellow
    }

    # ------------------------------------------------------------------
    # OPCIÓN 3: RESTAURACIÓN COMPLETA DESDE UN BACKUP ZIP
    # Lista los paquetes ZIP disponibles y descomprime el elegido.
    # ------------------------------------------------------------------
    "3" {
        if (-not (Test-Path $RutaBackups)) {
            Write-Host "`n[ERROR] No existe la carpeta de respaldos ($RutaBackups)." -ForegroundColor Red
            return
        }

        # Obtiene los respaldos disponibles ordenados del más reciente al más antiguo
        $Backups = @(Get-ChildItem -Path $RutaBackups -Filter "Backup_PreRollback_*.zip" | Sort-Object LastWriteTime -Descending)
        if ($Backups.Count -eq 0) {
            Write-Host "`n[INFO] No hay archivos de respaldo disponibles para restaurar." -ForegroundColor Yellow
            return
        }

        # Muestra el catálogo numerado de respaldos al usuario
        Write-Host "`nRespaldos disponibles:" -ForegroundColor Cyan
        for ($i = 0; $i -lt $Backups.Count; $i++) {
            Write-Host "  [$($i + 1)] $($Backups[$i].Name)  ($(Get-Date $Backups[$i].LastWriteTime -Format 'dd/MM/yyyy HH:mm:ss'))"
        }
        # Valida que la selección numérica esté dentro del rango válido
        $Sel = Read-Host "`nIngresa el número del respaldo a restaurar (1-$($Backups.Count))"
        $Idx = 0
        if (-not [int]::TryParse($Sel, [ref]$Idx) -or $Idx -lt 1 -or $Idx -gt $Backups.Count) {
            Write-Host "[ERROR] Selección inválida." -ForegroundColor Red
            return
        }
        $ZipElegido = $Backups[$Idx - 1]
        # Limpia archivos actuales y extrae el contenido íntegro del ZIP seleccionado
        if (-not (Test-Path $RutaReportes)) { New-Item -ItemType Directory -Path $RutaReportes -Force | Out-Null }
        Get-ChildItem -Path $RutaReportes -File | Remove-Item -Force -ErrorAction SilentlyContinue
        Expand-Archive -Path $ZipElegido.FullName -DestinationPath $RutaReportes -Force
        Write-RollbackLog "Entorno restaurado desde el respaldo: $($ZipElegido.Name)"
        Write-Host "`n[OK] Reportes restaurados exitosamente desde $($ZipElegido.Name)." -ForegroundColor Green
        }
    # ------------------------------------------------------------------
    # OPCIÓN 4: SALIDA SEGURA
    # Finaliza el script sin realizar ninguna modificación.
    # ------------------------------------------------------------------
    default {
    	Write-Host "`nOperación cancelada por el usuario." -ForegroundColor Gray
    }
}
