########################################################################
# Health Check Enterprise
# Autor: Victor Manuel Segundo Marquez
# Objetivo:
# Recopilar información de sistema, hardware, almacenamiento,
# red, seguridad y rendimiento para generar un reporte de salud.

########################################################################

# Define la ruta de salida del reporte.
$RutaReporte = "C:\Semana9-Automatizacion\Reportes\HealthCheck_Report.csv"

# Verifica la existencia de la carpeta de reportes.
$Directorio = Split-Path $RutaReporte

if (!(Test-Path $Directorio))
{
    New-Item -ItemType Directory -Path $Directorio -Force | Out-Null
}

########################################################################
# LOG Y CONTROL DE EJECUCIÓN
########################################################################

$LogFile = "C:\Semana9-Automatizacion\Reportes\HealthCheck_Log.txt"

$TotalHallazgos = 0

# Registra hora inicial del proceso.
$InicioScript = Get-Date

function Write-Log
{
 param(
  [string]$Tipo,
  [string]$Mensaje
 )

 Add-Content `
  -Path $LogFile `
  -Value "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') [$Tipo] $Mensaje"
}

Write-Log "INFO" "Inicio de ejecución del Health Check."


########################################################################
# INFORMACIÓN GENERAL
########################################################################

# Obtiene la fecha y hora de ejecución.
$Fecha = Get-Date

# Obtiene el nombre del equipo.
$Equipo = $env:COMPUTERNAME

# Obtiene el usuario autenticado.
$Usuario = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name


Write-Log "INFO" "Equipo: $Equipo"

Write-Log "INFO" "Usuario Ejecutor: $Usuario"

# Consulta información del sistema operativo.
$SO = Get-CimInstance Win32_OperatingSystem

# Obtiene el nombre del sistema operativo.
$SistemaOperativo = $SO.Caption

# Obtiene la versión instalada.
$VersionSO = $SO.Version

# Calcula el tiempo desde el último arranque.
$Uptime = (Get-Date) - $SO.LastBootUpTime

########################################################################
# HARDWARE
########################################################################

# Consulta información del equipo.
$ComputerSystem = Get-CimInstance Win32_ComputerSystem

# Consulta información del BIOS.
$BIOS = Get-CimInstance Win32_BIOS

# Obtiene fabricante del equipo.
$Fabricante = $ComputerSystem.Manufacturer

# Obtiene modelo del equipo.
$Modelo = $ComputerSystem.Model

# Obtiene número de serie.
$NumeroSerie = $BIOS.SerialNumber

# Obtiene fabricante del BIOS.
$BIOSFabricante = $BIOS.Manufacturer

# Obtiene versión del BIOS.
$BIOSVersion = $BIOS.SMBIOSBIOSVersion

########################################################################
# PROCESADOR
########################################################################

# Consulta información del procesador.
$CPU = Get-CimInstance Win32_Processor

# Obtiene nombre del procesador.
$CPU_Nombre = $CPU.Name

# Obtiene cantidad de núcleos físicos.
$CPU_Nucleos = $CPU.NumberOfCores

# Obtiene cantidad de procesadores lógicos.
$CPU_Logicos = $CPU.NumberOfLogicalProcessors

# Obtiene la carga actual del procesador.
$CPU_Uso = (
Get-CimInstance Win32_Processor |
Measure-Object LoadPercentage -Average
).Average

if($CPU_Uso -gt 85)
{
    Write-Log "ALERTA" "Uso elevado de CPU ($CPU_Uso%)."

    $TotalHallazgos++
}

########################################################################
# MEMORIA RAM
########################################################################

# Calcula la memoria física instalada.
$RAM_Total = [Math]::Round(
$ComputerSystem.TotalPhysicalMemory / 1GB,
2
)

# Obtiene memoria física disponible.
$RAM_Disponible = [Math]::Round(
$SO.FreePhysicalMemory / 1MB,
2
)

# Calcula porcentaje de uso de memoria.
$RAM_Uso = [Math]::Round(
(($RAM_Total - $RAM_Disponible) / $RAM_Total) * 100,
2
)

if($RAM_Uso -gt 90)
{
    Write-Log "ALERTA" "Uso elevado de memoria RAM ($RAM_Uso%)."

    $TotalHallazgos++
}

########################################################################
# ALMACENAMIENTO
########################################################################

# Consulta la unidad principal.
try
{
    $Disco = Get-CimInstance `
    Win32_LogicalDisk `
    -Filter "DeviceID='C:'" `
    -ErrorAction Stop
}
catch
{
    Write-Log "ERROR" "No fue posible obtener información del disco."

    $TotalHallazgos++
}

# Calcula capacidad total del disco.
$Disco_Total = [Math]::Round(
$Disco.Size / 1GB,
2
)

# Calcula espacio libre disponible.
$Disco_Libre = [Math]::Round(
$Disco.FreeSpace / 1GB,
2
)

# Calcula porcentaje de utilización.
$Disco_Uso = [Math]::Round(
(($Disco_Total - $Disco_Libre) / $Disco_Total) * 100,
2
)

if($Disco_Libre -lt 20)
{
    Write-Log "ALERTA" "Menos de 20 GB libres en disco."

    $TotalHallazgos++
}

########################################################################
# ESTADO DEL SSD
########################################################################

# Obtiene el estado operacional del disco.
$PhysicalDisk = Get-PhysicalDisk | Select-Object -First 1

$DiskHealth = $PhysicalDisk.HealthStatus
$DiskOperationalStatus = $PhysicalDisk.OperationalStatus

########################################################################
# RED
########################################################################

# Identifica la ruta por defecto utilizada para acceder a Internet.
$RutaActiva = Get-NetRoute `
-DestinationPrefix "0.0.0.0/0" |
Sort-Object RouteMetric |
Select-Object -First 1

# Obtiene el índice de la interfaz utilizada por la ruta principal.
$InterfaceIndex = $RutaActiva.InterfaceIndex

# Obtiene la configuración de la interfaz activa.
$NIC = Get-NetIPConfiguration |
Where-Object {
    $_.InterfaceIndex -eq $InterfaceIndex
}

# Obtiene el nombre del adaptador activo.
$AdaptadorActivo = $NIC.InterfaceAlias

# Obtiene la dirección IPv4 principal.
$IP = $NIC.IPv4Address.IPAddress

# Obtiene la puerta de enlace configurada.
$Gateway = $RutaActiva.NextHop

# Obtiene los servidores DNS asociados a la interfaz activa.
$DNS = (
Get-DnsClientServerAddress `
-InterfaceIndex $InterfaceIndex `
-AddressFamily IPv4
).ServerAddresses -join ", "

########################################################################
# CONECTIVIDAD
########################################################################

# Valida conectividad hacia Internet.
$Internet = Test-Connection `
www.microsoft.com `
-Count 2 `
-Quiet

if(!$Internet)
{
     Write-Log "ALERTA" "No existe conectividad hacia Internet."

     $TotalHallazgos++
}
########################################################################
# DOMINIO
########################################################################

# Obtiene el dominio configurado.
$Dominio = $ComputerSystem.Domain

# Valida si el equipo pertenece a dominio.
$PartOfDomain = $ComputerSystem.PartOfDomain

########################################################################
# MICROSOFT DEFENDER
########################################################################

# Consulta estado del antivirus.
try
{
    $Defender = Get-MpComputerStatus -ErrorAction Stop

    $DefenderActivo = $Defender.AntivirusEnabled

    $RealtimeProtection = $Defender.RealTimeProtectionEnabled

    $FirmaAV = $Defender.AntivirusSignatureLastUpdated

    if(!$DefenderActivo)
    {
        Write-Log "ALERTA" "Microsoft Defender se encuentra deshabilitado."

        $TotalHallazgos++
    }
}
catch
{
    Write-Log "ERROR" "No fue posible consultar Microsoft Defender."

    $TotalHallazgos++

    $DefenderActivo = "Error"

    $RealtimeProtection = "Error"

    $FirmaAV = "Error"
}

########################################################################
# FIREWALL
########################################################################

# Consulta perfiles de Firewall.
$FirewallProfiles = Get-NetFirewallProfile

# Valida que todos los perfiles estén activos.
$Firewall = ($FirewallProfiles.Enabled -notcontains $false)

if(!$Firewall)
{
    Write-Log "ALERTA" "Firewall deshabilitado."

    $TotalHallazgos++
}
########################################################################
# TPM
########################################################################

try
{
    # Consulta estado del TPM.
    $TPM = Get-Tpm -ErrorAction Stop

    $TpmPresent = $TPM.TpmPresent
    $TpmReady = $TPM.TpmReady
    Write-Log "INFO" "TPM consultado correctamente."
}
catch
{
    Write-Log "ERROR" "No fue posible consultar el TPM."

    $TotalHallazgos++

    $TpmPresent = "No Detectado"
    $TpmReady = "No Detectado"
}


########################################################################
# SECURE BOOT
########################################################################

try
{
    # Valida si Secure Boot está habilitado.
    $SecureBoot = Confirm-SecureBootUEFI -ErrorAction Stop
}
catch
{
    Write-Log "ADVERTENCIA" "Secure Boot no compatible."

    $TotalHallazgos++

    $SecureBoot = "No Compatible"
}

########################################################################
# SMBv1
########################################################################

# Verifica estado del protocolo SMBv1.
$SMB = Get-WindowsOptionalFeature `
-Online `
-FeatureName SMB1Protocol

$SMBv1 = $SMB.State

########################################################################
# BITLOCKER
########################################################################

try
{
    $BitLocker = Get-BitLockerVolume `
    -MountPoint "C:"`
    -ErrorAction Stop


    $BitLockerEstado = $BitLocker.ProtectionStatus
}
catch
{
    Write-Log "ERROR" "No fue posible consultar BitLocker."

    $TotalHallazgos++

    $BitLockerEstado = "Error"
}


########################################################################
# WINDOWS UPDATE
########################################################################

# 1. Obtiene el último parche instalado.
try
{
    $HotFix = Get-HotFix -ErrorAction Stop |
    Where-Object { $_.InstalledOn } |
    Sort-Object InstalledOn -Descending |
    Select-Object -First 1
}
catch
{
    Write-Log "ERROR" "No fue posible consultar las actualizaciones."

    $TotalHallazgos++

    $HotFix = $null
}

# 2. Obtiene parches pendientes (Críticos/Seguridad vs Opcionales).
$ParchesPendientes_CritSec    = 0
$ParchesPendientes_Opcionales = 0

try
{
    Write-Log "INFO" "Consultando actualizaciones pendientes en Windows Update..."
    $UpdateSession  = New-Object -ComObject Microsoft.Update.Session
    $UpdateSearcher = $UpdateSession.CreateUpdateSearcher()
    $SearchResult   = $UpdateSearcher.Search("IsInstalled=0 and Type='Software' and IsHidden=0")

    foreach ($Update in $SearchResult.Updates)
    {
        $Categorias = $Update.Categories | ForEach-Object { $_.Name }
        if ($Categorias -match "Security|Critical" -or $Update.MsrcSeverity -in @("Critical", "Important"))
        {
            $ParchesPendientes_CritSec++
        }
        else
        {
            $ParchesPendientes_Opcionales++
        }
    }

    if ($ParchesPendientes_CritSec -gt 0)
    {
        Write-Log "ALERTA" "Existen $ParchesPendientes_CritSec parches críticos o de seguridad pendientes."
        $TotalHallazgos++
    }
}
catch
{
    Write-Log "ADVERTENCIA" "No fue posible consultar parches pendientes."
    $ParchesPendientes_CritSec    = "N/D"
    $ParchesPendientes_Opcionales = "N/D"
}



########################################################################
# SERVICIOS
########################################################################

# Cuenta servicios automáticos detenidos.
$ServiciosAutomaticosDetenidos = (
Get-Service |
Where-Object {
$_.StartType -eq "Automatic" -and
$_.Status -ne "Running"
}
).Count

########################################################################
# EVENTOS
########################################################################

# Cuenta errores críticos registrados en 7 días.
$EventosCriticos = (
Get-WinEvent `
-FilterHashtable @{
LogName = "System"
Level = 2
StartTime = (Get-Date).AddDays(-7)
}
).Count

if($EventosCriticos -gt 1000)
{
    Write-Log "ALERTA" "Existen $EventosCriticos eventos críticos en los últimos 7 días."

    $TotalHallazgos++
}

########################################################################
# SALUD GENERAL
########################################################################

# Inicializa el estado general.
$EstadoGeneral = "OK"

if($Disco_Libre -lt 20)
{
    $EstadoGeneral = "CRITICO-DISCO"
}
elseif($RAM_Uso -gt 90)
{
    $EstadoGeneral = "CRITICO-RAM"
}
elseif($CPU_Uso -gt 85)
{
    $EstadoGeneral = "CRITICO-CPU"
}
elseif(!$DefenderActivo)
{
    $EstadoGeneral = "RIESGO-DEFENDER"
}

elseif(!$Firewall)
{
    $EstadoGeneral = "RIESGO-FIREWALL"
}
elseif($ParchesPendientes_CritSec -is [int] -and $ParchesPendientes_CritSec -gt 0)
{
    $EstadoGeneral = "RIESGO-PARCHES"
}


########################################################################
# HEALTH SCORE
########################################################################

$HealthScore = 100

$HealthScore -= ($TotalHallazgos * 5)

if($EventosCriticos -gt 1000)
{
    $HealthScore -= 10
}

if($ServiciosAutomaticosDetenidos -gt 10)
{
    $HealthScore -= 5
}

if($HealthScore -lt 0)
{
    $HealthScore = 0
}

########################################################################
# SEMAFORO
########################################################################

if($HealthScore -ge 90)
{
    $Semaforo = "VERDE"
}
elseif($HealthScore -ge 70)
{
    $Semaforo = "AMARILLO"
}
else
{
    $Semaforo = "ROJO"
}


$FinScript = Get-Date

$Duracion = [Math]::Round(
    ($FinScript - $InicioScript).TotalSeconds,
    2
)


########################################################################
# CONSOLIDACIÓN
########################################################################

# Consolida los resultados del Health Check.
$Reporte = [PSCustomObject]@{
Fecha = $Fecha
Equipo = $Equipo
Usuario = $Usuario
SistemaOperativo = $SistemaOperativo
VersionSO = $VersionSO
Dominio = $Dominio
PartOfDomain = $PartOfDomain

Fabricante = $Fabricante
Modelo = $Modelo
NumeroSerie = $NumeroSerie

BIOSFabricante = $BIOSFabricante
BIOSVersion = $BIOSVersion

CPU = $CPU_Nombre
CPU_Nucleos = $CPU_Nucleos
CPU_Logicos = $CPU_Logicos
CPU_Uso_Porcentaje = $CPU_Uso

RAM_Total_GB = $RAM_Total
RAM_Disponible_GB = $RAM_Disponible
RAM_Uso_Porcentaje = $RAM_Uso

Disco_Total_GB = $Disco_Total
Disco_Libre_GB = $Disco_Libre
Disco_Uso_Porcentaje = $Disco_Uso

DiskHealthStatus = $DiskHealth
DiskOperationalStatus = $DiskOperationalStatus

IP = $IP
Gateway = $Gateway
DNS = $DNS
Internet = $Internet

Defender = $DefenderActivo
ProteccionTiempoReal = $RealtimeProtection
FirmasAV = $FirmaAV

Firewall = $Firewall

TPMPresent = $TpmPresent
TPMReady = $TpmReady

SecureBoot = $SecureBoot

SMBv1 = $SMBv1
BitLocker = $BitLockerEstado

UltimoHotFix = if($HotFix){$HotFix.HotFixID}else{"No Disponible"}
FechaHotFix = if($HotFix){$HotFix.InstalledOn.ToString("yyyy-MM-dd")}else{"No Disponible"}
ParchesPendientes_CritSec = $ParchesPendientes_CritSec
ParchesPendientes_Opcionales = $ParchesPendientes_Opcionales

ServiciosAutomaticosDetenidos = $ServiciosAutomaticosDetenidos


EventosCriticos7Dias = $EventosCriticos

UptimeDias = $Uptime.Days
UptimeHoras = $Uptime.Hours

EstadoGeneral = $EstadoGeneral

HealthScore = $HealthScore

Semaforo = $Semaforo

TiempoEjecucionSegundos = $Duracion

TotalHallazgos = $TotalHallazgos


}

########################################################################
# EXPORTACIÓN
########################################################################

try
{
    # Reporte actual (se reemplaza en cada ejecución)
    $Reporte |
    Export-Csv `
        -Path $RutaReporte `
        -NoTypeInformation `
        -Encoding UTF8

    # Histórico (se conserva cada ejecución)
    $RutaHistorico =
    "C:\Semana9-Automatizacion\Reportes\HealthCheck_History.csv"

    # Si el histórico no existe, crea el archivo con encabezados
    if (!(Test-Path $RutaHistorico))
    {
        $Reporte |
        Export-Csv `
            -Path $RutaHistorico `
            -NoTypeInformation `
            -Encoding UTF8
    }
    else
    {
        $Reporte |
        Export-Csv `
            -Path $RutaHistorico `
            -NoTypeInformation `
            -Append `
            -Encoding UTF8
    }

    Write-Log "INFO" "Reporte exportado correctamente."
}
catch
{
    Write-Log "ERROR" "Falló la exportación del reporte."

    $TotalHallazgos++
}

########################################################################
# RESULTADO
########################################################################

# Muestra los resultados en pantalla.
$Reporte | Format-List

Write-Host ""
Write-Host "Health Check completado correctamente."
Write-Host "Archivo generado:"
Write-Host $RutaReporte
Write-Log "INFO" "Hallazgos detectados: $TotalHallazgos"
Write-Log "INFO" "Health Score: $HealthScore"
Write-Log "INFO" "Semáforo: $Semaforo"
Write-Log "INFO" "Duración: $Duracion segundos"
Write-Log "INFO" "Estado general: $EstadoGeneral"
Write-Log "INFO" "Fin de la ejecución."