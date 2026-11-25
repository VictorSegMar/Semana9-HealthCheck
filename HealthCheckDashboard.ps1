########################################################################
# DASHBOARD HEALTH CHECK EJECUTIVO
# Autor: Victor Manuel Segundo Marquez
#
# Objetivo:
# Leer los archivos generados por HealthCheck.ps1 (Report, History, Log)
# y construir un Dashboard Ejecutivo dinámico, mediante ImportExcel + EPPlus.
########################################################################

Import-Module ImportExcel

########################################################################
# 1. RUTAS Y VALIDACIÓN DE ARCHIVOS
########################################################################

$RutaBase       = "C:\Semana9-Automatizacion\Reportes"
$RutaReporte    = Join-Path $RutaBase "HealthCheck_Report.csv"
$RutaHistorico  = Join-Path $RutaBase "HealthCheck_History.csv"
$RutaLog        = Join-Path $RutaBase "HealthCheck_Log.txt"
$RutaDashboard  = Join-Path $RutaBase "HealthCheck_Dashboard.xlsx"

foreach ($Archivo in @($RutaReporte, $RutaHistorico, $RutaLog)) {
    if (-not (Test-Path $Archivo)) {
        Write-Host "[ERROR] No se encontró el archivo requerido: $Archivo" -ForegroundColor Red
        return
    }
}

########################################################################
# 2. CARGA Y PROCESAMIENTO DE DATOS
########################################################################

$Reporte   = Import-Csv $RutaReporte -Encoding UTF8
$Historico = Import-Csv $RutaHistorico -Encoding UTF8
$LogsRaw   = Get-Content $RutaLog -Encoding UTF8

$Actual = $Reporte | Select-Object -First 1

# Parseo estructurado del Log para auditoría

$LogEstructurado = foreach ($Linea in $LogsRaw) {
    if ($Linea -match '^(\d{4}-\d{2}-\d{2}\s+\d{2}:\d{2}:\d{2})\s+\[([^\]]+)\]\s+(.*)$') {
        [PSCustomObject]@{
            Fecha_Hora = $Matches[1]
            Nivel      = $Matches[2]
            Mensaje    = $Matches[3]
        }
    }
}

########################################################################
# 3. EVALUACIÓN DE CUMPLIMIENTO DE SEGURIDAD (8 CONTROLES)
########################################################################

$ListaControles = @(
    [PSCustomObject]@{ Control = "Microsoft Defender";     Categoria = "Antimalware";       Cumple = ($Actual.Defender -eq "True") }
    [PSCustomObject]@{ Control = "Protección Tiempo Real"; Categoria = "Antimalware";       Cumple = ($Actual.ProteccionTiempoReal -eq "True") }
    [PSCustomObject]@{ Control = "Firewall Windows";       Categoria = "Red";               Cumple = ($Actual.Firewall -eq "True") }
    [PSCustomObject]@{ Control = "TPM Presente";           Categoria = "Hardware Security"; Cumple = ($Actual.TPMPresent -eq "True") }
    [PSCustomObject]@{ Control = "TPM Listo (Ready)";      Categoria = "Hardware Security"; Cumple = ($Actual.TPMReady -eq "True") }
    [PSCustomObject]@{ Control = "Secure Boot";            Categoria = "Arranque Seguro";   Cumple = ($Actual.SecureBoot -eq "True") }
    [PSCustomObject]@{ Control = "SMBv1 Deshabilitado";    Categoria = "Hardening";         Cumple = ($Actual.SMBv1 -eq "Disabled") }
    [PSCustomObject]@{ Control = "BitLocker Activo (C:)";  Categoria = "Cifrado";           Cumple = ($Actual.BitLocker -eq "On") }
)

$TablaSeguridad = $ListaControles | ForEach-Object {
    [PSCustomObject]@{
        Control   = $_.Control
        Categoria = $_.Categoria
        Estado    = if ($_.Cumple) { "CUMPLE" } else { "NO CUMPLE" }
        Puntaje   = if ($_.Cumple) { 100 } else { 0 }
    }
}

$TotalControles      = $ListaControles.Count
$ControlesOK         = ($ListaControles | Where-Object { $_.Cumple }).Count
$PorcentajeSeguridad = [Math]::Round((($ControlesOK / $TotalControles) * 100), 1)

########################################################################
# 4. EVALUACIÓN DINÁMICA DE RIESGOS Y PLAN DE ACCIÓN
########################################################################

$Riesgos        = @()
$Recomendaciones = @()

if ($Actual.Defender -ne "True") {
    $Riesgos += [PSCustomObject]@{ ID = "RSK-01"; Componente = "Seguridad"; Hallazgo = "Microsoft Defender deshabilitado"; Severidad = "CRITICA"; Impacto = "Exposición total ante malware y ransomware" }
    $Recomendaciones += [PSCustomObject]@{ ID = "REC-01"; Prioridad = "ALTA"; Accion = "Habilitar servicio de Microsoft Defender Antivirus vía GPO o PowerShell"; Responsable = "Seguridad TI" }
}

if ($Actual.ProteccionTiempoReal -ne "True") {
    $Riesgos += [PSCustomObject]@{ ID = "RSK-02"; Componente = "Seguridad"; Hallazgo = "Protección en tiempo real deshabilitada"; Severidad = "CRITICA"; Impacto = "Ejecución de archivos maliciosos sin bloqueo preventivo" }
    $Recomendaciones += [PSCustomObject]@{ ID = "REC-02"; Prioridad = "ALTA"; Accion = "Activar RealTimeProtectionEnabled y verificar políticas de exclusión"; Responsable = "Seguridad TI" }
}

if ($Actual.SecureBoot -ne "True") {
    $Riesgos += [PSCustomObject]@{ ID = "RSK-03"; Componente = "Firmware"; Hallazgo = "Secure Boot deshabilitado"; Severidad = "MEDIA"; Impacto = "Riesgo de rootkits o bootkits en el arranque del SO" }
    $Recomendaciones += [PSCustomObject]@{ ID = "REC-03"; Prioridad = "MEDIA"; Accion = "Habilitar Secure Boot en configuración UEFI/BIOS del equipo"; Responsable = "Soporte Técnico" }
}

if ($Actual.SMBv1 -ne "Disabled") {
    $Riesgos += [PSCustomObject]@{ ID = "RSK-04"; Componente = "Red / Hardening"; Hallazgo = "Protocolo SMBv1 habilitado"; Severidad = "CRITICA"; Impacto = "Vulnerabilidad crítica frente a exploits laterales (ej. EternalBlue)" }
    $Recomendaciones += [PSCustomObject]@{ ID = "REC-04"; Prioridad = "ALTA"; Accion = "Deshabilitar SMB1Protocol mediante Disable-WindowsOptionalFeature"; Responsable = "Infraestructura" }
}

if ($Actual.BitLocker -ne "On") {
    $Riesgos += [PSCustomObject]@{ ID = "RSK-05"; Componente = "Cifrado"; Hallazgo = "BitLocker no activo en unidad C:"; Severidad = "ALTA"; Impacto = "Exposición de datos corporativos en caso de robo o extravío" }
    $Recomendaciones += [PSCustomObject]@{ ID = "REC-05"; Prioridad = "ALTA"; Accion = "Activar cifrado BitLocker y respaldar clave de recuperación en AD/Entra ID"; Responsable = "Seguridad TI" }
}

if ([double]$Actual.Disco_Uso_Porcentaje -gt 80) {
    $Riesgos += [PSCustomObject]@{ ID = "RSK-06"; Componente = "Almacenamiento"; Hallazgo = "Uso de disco elevado ($($Actual.Disco_Uso_Porcentaje)%)"; Severidad = "MEDIA"; Impacto = "Degradación de rendimiento y fallo en instalación de parches" }
    $Recomendaciones += [PSCustomObject]@{ ID = "REC-06"; Prioridad = "MEDIA"; Accion = "Ejecutar limpieza de archivos temporales, caché de Windows Update y perfiles antiguos"; Responsable = "Soporte Técnico" }
}

if ([int]$Actual.EventosCriticos7Dias -gt 50) {
    $Riesgos += [PSCustomObject]@{ ID = "RSK-07"; Componente = "Estabilidad SO"; Hallazgo = "Alto volumen de eventos críticos en 7 días ($($Actual.EventosCriticos7Dias) eventos)"; Severidad = "ADVERTENCIA"; Impacto = "Posibles errores recurrentes de controladores, servicios o aplicaciones" }
    $Recomendaciones += [PSCustomObject]@{ ID = "REC-07"; Prioridad = "MEDIA"; Accion = "Auditar Visor de Eventos (Log System Level 2) para depurar origen de errores repetitivos"; Responsable = "Operaciones TI" }
}

if ([int]$Actual.ServiciosAutomaticosDetenidos -gt 10) {
    $Riesgos += [PSCustomObject]@{ ID = "RSK-08"; Componente = "Servicios"; Hallazgo = "$($Actual.ServiciosAutomaticosDetenidos) servicios automáticos detenidos"; Severidad = "BAJA"; Impacto = "Funcionalidades secundarias o agentes de gestión sin ejecución" }
    $Recomendaciones += [PSCustomObject]@{ ID = "REC-08"; Prioridad = "BAJA"; Accion = "Validar si los servicios detenidos corresponden a tareas diferidas (Delayed Start) o fallos"; Responsable = "Operaciones TI" }
}

if ($Riesgos.Count -eq 0) {
    $Riesgos += [PSCustomObject]@{ ID = "RSK-00"; Componente = "General"; Hallazgo = "Sin riesgos críticos detectados"; Severidad = "OK"; Impacto = "Operación dentro de parámetros normales" }
    $Recomendaciones += [PSCustomObject]@{ ID = "REC-00"; Prioridad = "INFORMATIVA"; Accion = "Mantener monitoreo preventivo programado"; Responsable = "Operaciones TI" }
}

########################################################################
# 5. EXPORTACIÓN BASE CON IMPORTEXCEL
########################################################################

if (Test-Path $RutaDashboard) {
    Remove-Item $RutaDashboard -Force
}

$Excel = Export-Excel -Path $RutaDashboard -WorksheetName "01_ResumenEjecutivo" -PassThru

if ($Reporte) {
    $Excel = $Reporte | Export-Excel -ExcelPackage $Excel -WorksheetName "02_UltimaEjecucion" -StartRow 3 -TableName "TblUltimaEjecucion" -TableStyle Light1 -AutoSize -PassThru
} else {
    $Excel.Workbook.Worksheets.Add("02_UltimaEjecucion") | Out-Null
}

if ($Historico) {
    $Excel = $Historico | Export-Excel -ExcelPackage $Excel -WorksheetName "03_Historico" -StartRow 3 -TableName "TblHistorico" -TableStyle Light1 -AutoSize -PassThru
} else {
    $Excel.Workbook.Worksheets.Add("03_Historico") | Out-Null
}

$Excel = $TablaSeguridad | Export-Excel -ExcelPackage $Excel -WorksheetName "04_Seguridad"       -StartRow 3 -TableName "TblSeguridad"       -TableStyle Light1 -AutoSize -PassThru
$Excel = $Riesgos        | Export-Excel -ExcelPackage $Excel -WorksheetName "05_RiesgosYAcciones" -StartRow 4 -TableName "TblRiesgos"         -TableStyle Light1 -AutoSize -PassThru

$FilaInicioRec = 6 + $Riesgos.Count
$Excel = $Recomendaciones | Export-Excel -ExcelPackage $Excel -WorksheetName "05_RiesgosYAcciones" -StartRow $FilaInicioRec -TableName "TblRecomendaciones" -TableStyle Light1 -AutoSize -PassThru

if ($LogEstructurado) {
    $Excel = $LogEstructurado | Export-Excel -ExcelPackage $Excel -WorksheetName "06_LogEjecucion" -StartRow 3 -TableName "TblLogEjecucion" -TableStyle Light1 -AutoSize -PassThru
} else {
    $Excel.Workbook.Worksheets.Add("06_LogEjecucion") | Out-Null
}


$wsAux = $Excel.Workbook.Worksheets.Add("99_DatosGraficos")


########################################################################
#6. PALETA DE COLORES CORPORATIVA Y FUNCIONES AUXILIARES DE ESTILO
########################################################################

$ColorNavyDark   = [System.Drawing.Color]::FromArgb(15, 37, 55)     # #0F2537
$ColorNavyMed    = [System.Drawing.Color]::FromArgb(30, 78, 121)    # #1E4E79
$ColorSlateBG    = [System.Drawing.Color]::FromArgb(244, 246, 249)  # #F4F6F9
$ColorWhite      = [System.Drawing.Color]::White
$ColorBorder     = [System.Drawing.Color]::FromArgb(217, 225, 232)  # #D9E1E8
$ColorTextDark   = [System.Drawing.Color]::FromArgb(31, 41, 55)     # #1F2937
$ColorTextMuted  = [System.Drawing.Color]::FromArgb(107, 114, 128)  # #6B7280
$ColorGreenBG    = [System.Drawing.Color]::FromArgb(16, 124, 65)    # #107C41
$ColorGreenLight = [System.Drawing.Color]::FromArgb(230, 244, 234)  # #E6F4EA
$ColorAmberBG    = [System.Drawing.Color]::FromArgb(245, 158, 11)   # #F59E0B
$ColorAmberLight = [System.Drawing.Color]::FromArgb(254, 243, 199)  # #FEF3C7
$ColorRedBG      = [System.Drawing.Color]::FromArgb(200, 30, 30)    # #C81E1E
$ColorRedLight   = [System.Drawing.Color]::FromArgb(253, 232, 232)  # #FDE8E8
function Set-BannerTitulo {
param($Worksheet, [string]$Rango, [string]$Texto)
$Worksheet.Cells[$Rango].Merge = $true
$celda = $Worksheet.Cells[$Rango.Split(":")[0]]
$celda.Value = $Texto
$celda.Style.Font.Name = "Calibri"
$celda.Style.Font.Size = 14
$celda.Style.Font.Bold = $true
$celda.Style.Font.Color.SetColor($ColorWhite)
$celda.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
$celda.Style.Fill.BackgroundColor.SetColor($ColorNavyDark)
$celda.Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Left
$celda.Style.VerticalAlignment = [OfficeOpenXml.Style.ExcelVerticalAlignment]::Center
$celda.Style.Indent = 1
}
function Set-BordeFino {
param($Worksheet, [string]$Rango)
$Worksheet.Cells[$Rango].Style.Border.Top.Style    = "Thin"
$Worksheet.Cells[$Rango].Style.Border.Bottom.Style = "Thin"
$Worksheet.Cells[$Rango].Style.Border.Left.Style   = "Thin"
$Worksheet.Cells[$Rango].Style.Border.Right.Style  = "Thin"
$Worksheet.Cells[$Rango].Style.Border.Top.Color.SetColor($ColorBorder)
$Worksheet.Cells[$Rango].Style.Border.Bottom.Color.SetColor($ColorBorder)
$Worksheet.Cells[$Rango].Style.Border.Left.Color.SetColor($ColorBorder)
$Worksheet.Cells[$Rango].Style.Border.Right.Color.SetColor($ColorBorder)
}
########################################################################

#7. POBLAR HOJA AUXILIAR (99_DatosGraficos) PARA GRÁFICOS
########################################################################

$cpuUso   = [double]$Actual.CPU_Uso_Porcentaje
$ramUso   = [double]$Actual.RAM_Uso_Porcentaje
$discoUso = [double]$Actual.Disco_Uso_Porcentaje

#Tablas individuales para Donas

$wsAux.Cells["E1"].Value = "CPU";   $wsAux.Cells["F1"].Value = "Porcentaje"
$wsAux.Cells["E2"].Value = "En Uso"; $wsAux.Cells["F2"].Value = $cpuUso
$wsAux.Cells["E3"].Value = "Libre";  $wsAux.Cells["F3"].Value = [Math]::Round((100 - $cpuUso), 2)
$wsAux.Cells["H1"].Value = "RAM";   $wsAux.Cells["I1"].Value = "Porcentaje"
$wsAux.Cells["H2"].Value = "En Uso"; $wsAux.Cells["I2"].Value = $ramUso
$wsAux.Cells["H3"].Value = "Libre";  $wsAux.Cells["I3"].Value = [Math]::Round((100 - $ramUso), 2)
$wsAux.Cells["K1"].Value = "Disco"; $wsAux.Cells["L1"].Value = "Porcentaje"
$wsAux.Cells["K2"].Value = "En Uso"; $wsAux.Cells["L2"].Value = $discoUso
$wsAux.Cells["K3"].Value = "Libre";  $wsAux.Cells["L3"].Value = [Math]::Round((100 - $discoUso), 2)

#Tabla de Tendencia Histórica (A7:E...)

$wsAux.Cells["A7"].Value = "Ejecución"
$wsAux.Cells["B7"].Value = "CPU (%)"
$wsAux.Cells["C7"].Value = "RAM (%)"
$wsAux.Cells["D7"].Value = "Disco (%)"
$wsAux.Cells["E7"].Value = "Health Score"
$filaHist = 8
foreach ($h in $Historico) {
$fechaStr = [string]$h.Fecha
$etiqueta = if ($fechaStr.Length -ge 16) { $fechaStr.Substring(0, 16) } else { $fechaStr }
$wsAux.Cells["A$filaHist"].Value = $etiqueta
$wsAux.Cells["B$filaHist"].Value = [double]$h.CPU_Uso_Porcentaje
$wsAux.Cells["C$filaHist"].Value = [double]$h.RAM_Uso_Porcentaje
$wsAux.Cells["D$filaHist"].Value = [double]$h.Disco_Uso_Porcentaje
$wsAux.Cells["E$filaHist"].Value = [double]$h.HealthScore
$filaHist++
}
$finHist = $filaHist - 1

# Tabla Radar de Seguridad (G7:H15)

$wsAux.Cells["G7"].Value = "Control Seguridad"
$wsAux.Cells["H7"].Value = "Cumplimiento (%)"
$filaSec = 8
foreach ($s in $TablaSeguridad) {
$wsAux.Cells["G$filaSec"].Value = $s.Control
$wsAux.Cells["H$filaSec"].Value = [double]$s.Puntaje
$filaSec++
}
$finSec = $filaSec - 1



########################################################################
# 8. ESTILIZAR PESTANAS DE DETALLE (02 a 06 y 99_DatosGraficos)
########################################################################

function Format-TablaEjecutiva {
    param(
        $Worksheet,
        [int]$FilaEncabezado,
        [int]$FilaFin,
        [int]$ColInicio = 1,
        [int]$ColFin,
        [int[]]$ColsCentradas = @(),
        [switch]$AjustarBanner
    )

    $Worksheet.View.ShowGridLines = $false
    if ($AjustarBanner) {
        $Worksheet.Row(1).Height = 30
        $Worksheet.Row(2).Height = 10
    }

    if ($FilaFin -lt $FilaEncabezado) { return }

    # Estilo de Encabezados (1 sola linea)
    $Worksheet.Row($FilaEncabezado).Height = 25
    for ($c = $ColInicio; $c -le $ColFin; $c++) {
        $cellHdr = $Worksheet.Cells[$FilaEncabezado, $c]
        $cellHdr.Style.Font.Name = "Calibri"
        $cellHdr.Style.Font.Size = 10
        $cellHdr.Style.Font.Bold = $true
        $cellHdr.Style.Font.Color.SetColor($ColorWhite)
        $cellHdr.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
        $cellHdr.Style.Fill.BackgroundColor.SetColor($ColorNavyMed)
        $cellHdr.Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Center
        $cellHdr.Style.VerticalAlignment   = [OfficeOpenXml.Style.ExcelVerticalAlignment]::Center
        $cellHdr.Style.WrapText = $false
    }

    # Estilo de Filas de Datos (1 sola linea limpia)
    for ($r = ($FilaEncabezado + 1); $r -le $FilaFin; $r++) {
        $Worksheet.Row($r).Height = 22
        $bgFila = if (($r - $FilaEncabezado) % 2 -eq 1) { $ColorWhite } else { $ColorSlateBG }

        for ($c = $ColInicio; $c -le $ColFin; $c++) {
            $cell = $Worksheet.Cells[$r, $c]
            $cell.Style.Font.Name = "Calibri"
            $cell.Style.Font.Size = 10
            $cell.Style.Font.Color.SetColor($ColorTextDark)
            $cell.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
            $cell.Style.Fill.BackgroundColor.SetColor($bgFila)
            $cell.Style.VerticalAlignment = [OfficeOpenXml.Style.ExcelVerticalAlignment]::Center
            $cell.Style.WrapText = $false

            if ($ColsCentradas -contains $c) {
                $cell.Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Center
            } else {
                $cell.Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Left
                $cell.Style.Indent = 1
            }
        }
    }

    # Bordes finos y ancho holgado para que todo quepa en 1 linea
    $rangoTabla = $Worksheet.Cells[$FilaEncabezado, $ColInicio, $FilaFin, $ColFin].Address
    Set-BordeFino -Worksheet $Worksheet -Rango $rangoTabla

    for ($c = $ColInicio; $c -le $ColFin; $c++) {
        $Worksheet.Column($c).AutoFit()
        $anchoActual = $Worksheet.Column($c).Width
        if ($anchoActual -lt 15) {
            $Worksheet.Column($c).Width = 15
        } else {
            $Worksheet.Column($c).Width = $anchoActual + 5
        }
    }
}


# Banners de Titulo (02 a 06)
Set-BannerTitulo -Worksheet $Excel.Workbook.Worksheets["02_UltimaEjecucion"] -Rango "A1:AV1" -Texto "REPORTE TECNICO DETALLADO - ULTIMA EJECUCION ($($Actual.Equipo))"
Set-BannerTitulo -Worksheet $Excel.Workbook.Worksheets["03_Historico"]       -Rango "A1:AV1" -Texto "HISTORICO DE EJECUCIONES HEALTH CHECK ($($Actual.Equipo))"
Set-BannerTitulo -Worksheet $Excel.Workbook.Worksheets["04_Seguridad"]       -Rango "A1:D1"  -Texto "MATRIZ DE CUMPLIMIENTO DE SEGURIDAD Y HARDENING"
Set-BannerTitulo -Worksheet $Excel.Workbook.Worksheets["05_RiesgosYAcciones"] -Rango "A1:E1"  -Texto "GESTION DE RIESGOS DETECTADOS Y PLAN DE ACCION EJECUTIVO"
Set-BannerTitulo -Worksheet $Excel.Workbook.Worksheets["06_LogEjecucion"]    -Rango "A1:C1"  -Texto "BITACORA DE AUDITORIA Y EJECUCION (LOG)"

# 02_UltimaEjecucion
$wsUlt = $Excel.Workbook.Worksheets["02_UltimaEjecucion"]
if ($wsUlt.Dimension) {
    Format-TablaEjecutiva -Worksheet $wsUlt -FilaEncabezado 3 -FilaFin $wsUlt.Dimension.End.Row -ColInicio 1 -ColFin $wsUlt.Dimension.End.Column -ColsCentradas @(1,2,5,7,14,15,16,17,18,19,20,21,22,23,24,25,28,29,30,32,33,34,35,36,37,40,41,42,43,45,46,47,48) -AjustarBanner
}

# 03_Historico
$wsHist = $Excel.Workbook.Worksheets["03_Historico"]
if ($wsHist.Dimension) {
    Format-TablaEjecutiva -Worksheet $wsHist -FilaEncabezado 3 -FilaFin $wsHist.Dimension.End.Row -ColInicio 1 -ColFin $wsHist.Dimension.End.Column -ColsCentradas @(1,2,5,7,14,15,16,17,18,19,20,21,22,23,24,25,28,29,30,32,33,34,35,36,37,40,41,42,43,45,46,47,48) -AjustarBanner
}

# 04_Seguridad
$wsSec = $Excel.Workbook.Worksheets["04_Seguridad"]
$finSecTabla = 3 + $TablaSeguridad.Count
Format-TablaEjecutiva -Worksheet $wsSec -FilaEncabezado 3 -FilaFin $finSecTabla -ColInicio 1 -ColFin 4 -ColsCentradas @(3, 4) -AjustarBanner

for ($r = 4; $r -le $finSecTabla; $r++) {
    $valEstado = $wsSec.Cells["C$r"].Value
    $wsSec.Cells["C$r"].Style.Font.Bold = $true
    if ($valEstado -eq "CUMPLE") {
        $wsSec.Cells["C$r"].Style.Fill.BackgroundColor.SetColor($ColorGreenLight)
        $wsSec.Cells["C$r"].Style.Font.Color.SetColor($ColorGreenBG)
    } else {
        $wsSec.Cells["C$r"].Style.Fill.BackgroundColor.SetColor($ColorRedLight)
        $wsSec.Cells["C$r"].Style.Font.Color.SetColor($ColorRedBG)
    }
}

# 05_RiesgosYAcciones
$wsRisk = $Excel.Workbook.Worksheets["05_RiesgosYAcciones"]
$finRiesgos = 4 + $Riesgos.Count
$finRec     = $FilaInicioRec + $Recomendaciones.Count

Format-TablaEjecutiva -Worksheet $wsRisk -FilaEncabezado 4 -FilaFin $finRiesgos -ColInicio 1 -ColFin 5 -ColsCentradas @(1, 4) -AjustarBanner
Format-TablaEjecutiva -Worksheet $wsRisk -FilaEncabezado $FilaInicioRec -FilaFin $finRec -ColInicio 1 -ColFin 4 -ColsCentradas @(1, 2, 4)

$wsRisk.Row(3).Height = 24
$wsRisk.Cells["A3"].Value = "1. RIESGOS Y HALLAZGOS IDENTIFICADOS"
$wsRisk.Cells["A3"].Style.Font.Size = 11
$wsRisk.Cells["A3"].Style.Font.Bold = $true
$wsRisk.Cells["A3"].Style.Font.Color.SetColor($ColorNavyDark)

$wsRisk.Row($FilaInicioRec - 1).Height = 24
$wsRisk.Cells["A$($FilaInicioRec - 1)"].Value = "2. PLAN DE REMEDIACION RECOMENDADO"
$wsRisk.Cells["A$($FilaInicioRec - 1)"].Style.Font.Size = 11
$wsRisk.Cells["A$($FilaInicioRec - 1)"].Style.Font.Bold = $true
$wsRisk.Cells["A$($FilaInicioRec - 1)"].Style.Font.Color.SetColor($ColorNavyDark)

for ($r = 5; $r -le $finRiesgos; $r++) {
    $wsRisk.Row($r).Height = 28
    $sev = [string]$wsRisk.Cells["D$r"].Value
    $wsRisk.Cells["D$r"].Style.Font.Bold = $true
    if ($sev -match "CRITICA|ALTA") {
        $wsRisk.Cells["D$r"].Style.Fill.BackgroundColor.SetColor($ColorRedLight)
        $wsRisk.Cells["D$r"].Style.Font.Color.SetColor($ColorRedBG)
    } elseif ($sev -match "MEDIA|ADVERTENCIA") {
        $wsRisk.Cells["D$r"].Style.Fill.BackgroundColor.SetColor($ColorAmberLight)
        $wsRisk.Cells["D$r"].Style.Font.Color.SetColor([System.Drawing.Color]::FromArgb(180, 83, 9))
    } else {
        $wsRisk.Cells["D$r"].Style.Fill.BackgroundColor.SetColor($ColorGreenLight)
        $wsRisk.Cells["D$r"].Style.Font.Color.SetColor($ColorGreenBG)
    }
}

for ($r = ($FilaInicioRec + 1); $r -le $finRec; $r++) {
    $wsRisk.Row($r).Height = 28
    $prio = [string]$wsRisk.Cells["B$r"].Value
    $wsRisk.Cells["B$r"].Style.Font.Bold = $true
    if ($prio -eq "ALTA") {
        $wsRisk.Cells["B$r"].Style.Fill.BackgroundColor.SetColor($ColorRedLight)
        $wsRisk.Cells["B$r"].Style.Font.Color.SetColor($ColorRedBG)
    } elseif ($prio -eq "MEDIA") {
        $wsRisk.Cells["B$r"].Style.Fill.BackgroundColor.SetColor($ColorAmberLight)
        $wsRisk.Cells["B$r"].Style.Font.Color.SetColor([System.Drawing.Color]::FromArgb(180, 83, 9))
    } else {
        $wsRisk.Cells["B$r"].Style.Fill.BackgroundColor.SetColor($ColorGreenLight)
        $wsRisk.Cells["B$r"].Style.Font.Color.SetColor($ColorGreenBG)
    }
}

# 06_LogEjecucion
$wsLog = $Excel.Workbook.Worksheets["06_LogEjecucion"]
if ($wsLog.Dimension) {
    $finLog = $wsLog.Dimension.End.Row
    Format-TablaEjecutiva -Worksheet $wsLog -FilaEncabezado 3 -FilaFin $finLog -ColInicio 1 -ColFin 3 -ColsCentradas @(1, 2) -AjustarBanner
    for ($r = 4; $r -le $finLog; $r++) {
        $niv = [string]$wsLog.Cells["B$r"].Value
        $wsLog.Cells["B$r"].Style.Font.Bold = $true
        if ($niv -match "ERROR|CRIT") {
            $wsLog.Cells["B$r"].Style.Fill.BackgroundColor.SetColor($ColorRedLight)
            $wsLog.Cells["B$r"].Style.Font.Color.SetColor($ColorRedBG)
        } elseif ($niv -match "ALERTA|WARN") {
            $wsLog.Cells["B$r"].Style.Fill.BackgroundColor.SetColor($ColorAmberLight)
            $wsLog.Cells["B$r"].Style.Font.Color.SetColor([System.Drawing.Color]::FromArgb(180, 83, 9))
        } else {
            $wsLog.Cells["B$r"].Style.Font.Color.SetColor($ColorNavyMed)
        }
    }
}

# 99_DatosGraficos (Tablas auxiliares limpias y alineadas)
Format-TablaEjecutiva -Worksheet $wsAux -FilaEncabezado 1 -FilaFin 3 -ColInicio 5  -ColFin 6  -ColsCentradas @(6)
Format-TablaEjecutiva -Worksheet $wsAux -FilaEncabezado 1 -FilaFin 3 -ColInicio 8  -ColFin 9  -ColsCentradas @(9)
Format-TablaEjecutiva -Worksheet $wsAux -FilaEncabezado 1 -FilaFin 3 -ColInicio 11 -ColFin 12 -ColsCentradas @(12)
Format-TablaEjecutiva -Worksheet $wsAux -FilaEncabezado 7 -FilaFin $finHist -ColInicio 1 -ColFin 5 -ColsCentradas @(1,2,3,4,5)
Format-TablaEjecutiva -Worksheet $wsAux -FilaEncabezado 7 -FilaFin $finSec  -ColInicio 7 -ColFin 8 -ColsCentradas @(8)

$wsAux.Column(1).Width  = 20
$wsAux.Column(7).Width  = 26
$wsAux.Column(8).Width  = 20




########################################################################
#9. CONSTRUCCIÓN DE PORTADA EJECUTIVA (01_ResumenEjecutivo)
########################################################################


$wsExec = $Excel.Workbook.Worksheets["01_ResumenEjecutivo"]
$wsExec.View.ShowGridLines = $false

#Cuadrícula limpia (Columna A y R como márgenes, B..Q como área de trabajo)

$wsExec.Column(1).Width = 3
for ($c = 2; $c -le 17; $c++) {
$wsExec.Column($c).Width = 11.5
}
$wsExec.Column(18).Width = 3
$wsExec.Row(1).Height  = 10
$wsExec.Row(2).Height  = 34
$wsExec.Row(3).Height  = 22
$wsExec.Row(4).Height  = 10
$wsExec.Row(5).Height  = 20
$wsExec.Row(6).Height  = 24
$wsExec.Row(7).Height  = 10
$wsExec.Row(8).Height  = 22
$wsExec.Row(9).Height  = 34
$wsExec.Row(10).Height = 18
$wsExec.Row(11).Height = 12

#Banner Principal

$wsExec.Cells["B2:Q2"].Merge = $true
$wsExec.Cells["B2"].Value = "REPORTE EJECUTIVO DE SALUD OPERATIVA E INFRAESTRUCTURA"
$wsExec.Cells["B2"].Style.Font.Size = 18
$wsExec.Cells["B2"].Style.Font.Bold = $true
$wsExec.Cells["B2"].Style.Font.Color.SetColor($ColorWhite)
$wsExec.Cells["B2"].Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
$wsExec.Cells["B2"].Style.Fill.BackgroundColor.SetColor($ColorNavyDark)
$wsExec.Cells["B2"].Style.VerticalAlignment = [OfficeOpenXml.Style.ExcelVerticalAlignment]::Center
$wsExec.Cells["B2"].Style.Indent = 1
$wsExec.Cells["B3:Q3"].Merge = $true
$wsExec.Cells["B3"].Value = "Auditoría Automatizada  |  Fecha de Corte: $($Actual.Fecha)  |  Dominio: $($Actual.Dominio)"
$wsExec.Cells["B3"].Style.Font.Size = 10
$wsExec.Cells["B3"].Style.Font.Italic = $true
$wsExec.Cells["B3"].Style.Font.Color.SetColor([System.Drawing.Color]::FromArgb(226, 232, 240))
$wsExec.Cells["B3"].Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
$wsExec.Cells["B3"].Style.Fill.BackgroundColor.SetColor($ColorNavyMed)
$wsExec.Cells["B3"].Style.VerticalAlignment = [OfficeOpenXml.Style.ExcelVerticalAlignment]::Center
$wsExec.Cells["B3"].Style.Indent = 1

#Barra de Inventario del Activo (B5:Q6)

$MetaBloques = @(
@{ RngLbl = "B5:E5"; Lbl = "EQUIPO / HOSTNAME"; RngVal = "B6:E6"; Val = "$($Actual.Equipo) ($($Actual.Fabricante))" }
@{ RngLbl = "F5:I5"; Lbl = "MODELO & SERIE";    RngVal = "F6:I6"; Val = "$($Actual.Modelo) | S/N: $($Actual.NumeroSerie)" }
@{ RngLbl = "J5:M5"; Lbl = "SISTEMA OPERATIVO"; RngVal = "J6:M6"; Val = "$($Actual.SistemaOperativo) (Build $($Actual.VersionSO))" }
@{ RngLbl = "N5:Q5"; Lbl = "PROCESADOR & RED";  RngVal = "N6:Q6"; Val = "$($Actual.CPU_Nucleos) Cores / $($Actual.RAM_Total_GB) GB RAM | IP: $($Actual.IP)" }
)
foreach ($m in $MetaBloques) {
$wsExec.Cells[$m.RngLbl].Merge = $true
$cTop = $wsExec.Cells[$m.RngLbl.Split(":")[0]]
$cTop.Value = $m.Lbl
$cTop.Style.Font.Size = 8
$cTop.Style.Font.Bold = $true
$cTop.Style.Font.Color.SetColor($ColorTextMuted)
$cTop.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
$cTop.Style.Fill.BackgroundColor.SetColor($ColorSlateBG)
$cTop.Style.Indent = 1
$wsExec.Cells[$m.RngVal].Merge = $true
$cBot = $wsExec.Cells[$m.RngVal.Split(":")[0]]
$cBot.Value = $m.Val
$cBot.Style.Font.Size = 10
$cBot.Style.Font.Bold = $true
$cBot.Style.Font.Color.SetColor($ColorTextDark)
$cBot.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
$cBot.Style.Fill.BackgroundColor.SetColor($ColorSlateBG)
$cBot.Style.Indent = 1
}
Set-BordeFino -Worksheet $wsExec -Rango "B5:Q6"

#Tarjetas KPI Ejecutivas (Filas 8 a 10)

$scoreVal = [int]$Actual.HealthScore
$scoreBG  = if ($scoreVal -ge 85) { $ColorGreenBG } elseif ($scoreVal -ge 70) { $ColorAmberBG } else { $ColorRedBG }
$scoreFG  = if ($scoreVal -ge 70 -and $scoreVal -lt 85) { $ColorTextDark } else { $ColorWhite }
$semVal   = ([string]$Actual.Semaforo).ToUpper()
$semBG    = switch ($semVal) { "VERDE" { $ColorGreenBG } "AMARILLO" { $ColorAmberBG } default { $ColorRedBG } }
$semFG    = if ($semVal -eq "AMARILLO") { $ColorTextDark } else { $ColorWhite }
$secBG    = if ($PorcentajeSeguridad -ge 90) { $ColorGreenBG } elseif ($PorcentajeSeguridad -ge 75) { $ColorAmberBG } else { $ColorRedBG }
$secFG    = if ($PorcentajeSeguridad -ge 75 -and $PorcentajeSeguridad -lt 90) { $ColorTextDark } else { $ColorWhite }
$evVal    = [int]$Actual.EventosCriticos7Dias
$evBodyBG = if ($evVal -gt 100) { $ColorRedLight } else { $ColorSlateBG }
$evBodyFG = if ($evVal -gt 100) { $ColorRedBG } else { $ColorNavyDark }
$TarjetasKPI = @(
@{ C1="B"; C2="C"; Titulo="HEALTH SCORE";         Valor="$scoreVal / 100";                         Sub="Índice Global";                     HdrBG=$scoreBG;      HdrFG=$scoreFG;    BodyBG=$scoreBG;      BodyFG=$scoreFG }
@{ C1="D"; C2="E"; Titulo="SEMÁFORO";             Valor=$semVal;                                   Sub="Estado: $($Actual.EstadoGeneral)"; HdrBG=$semBG;        HdrFG=$semFG;      BodyBG=$semBG;        BodyFG=$semFG }
@{ C1="F"; C2="H"; Titulo="SEGURIDAD";            Valor="$PorcentajeSeguridad%";                   Sub="$ControlesOK/$TotalControles Controles OK"; HdrBG=$secBG; HdrFG=$secFG;      BodyBG=$secBG;        BodyFG=$secFG }
@{ C1="I"; C2="K"; Titulo="EVENTOS CRÍTICOS (7D)";Valor=("{0:N0}" -f $evVal);                      Sub="Visor System Level 2";            HdrBG=$ColorNavyDark;HdrFG=$ColorWhite; BodyBG=$evBodyBG;     BodyFG=$evBodyFG }
@{ C1="L"; C2="N"; Titulo="SERVICIOS DETENIDOS";  Valor="$($Actual.ServiciosAutomaticosDetenidos)";Sub="Inicio Automático";                 HdrBG=$ColorNavyDark;HdrFG=$ColorWhite; BodyBG=$ColorSlateBG; BodyFG=$ColorNavyDark }
@{ C1="O"; C2="Q"; Titulo="UPTIME / HALLAZGOS";   Valor="$($Actual.UptimeDias)d $($Actual.UptimeHoras)h"; Sub="Hallazgos Activos: $($Actual.TotalHallazgos)"; HdrBG=$ColorNavyDark;HdrFG=$ColorWhite; BodyBG=$ColorSlateBG; BodyFG=$ColorNavyDark }
)
foreach ($k in $TarjetasKPI) {
$rH = "$($k.C1)8:$($k.C2)8"
$rV = "$($k.C1)9:$($k.C2)9"
$rS = "$($k.C1)10:$($k.C2)10"
$wsExec.Cells[$rH].Merge = $true
$cH = $wsExec.Cells["$($k.C1)8"]
$cH.Value = $k.Titulo
$cH.Style.Font.Size = 9; $cH.Style.Font.Bold = $true; $cH.Style.Font.Color.SetColor($k.HdrFG)
$cH.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
$cH.Style.Fill.BackgroundColor.SetColor($k.HdrBG)
$cH.Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Center
$cH.Style.VerticalAlignment   = [OfficeOpenXml.Style.ExcelVerticalAlignment]::Center
$wsExec.Cells[$rV].Merge = $true
$cV = $wsExec.Cells["$($k.C1)9"]
$cV.Value = $k.Valor
$cV.Style.Font.Size = 18; $cV.Style.Font.Bold = $true; $cV.Style.Font.Color.SetColor($k.BodyFG)
$cV.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
$cV.Style.Fill.BackgroundColor.SetColor($k.BodyBG)
$cV.Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Center
$cV.Style.VerticalAlignment   = [OfficeOpenXml.Style.ExcelVerticalAlignment]::Center
$wsExec.Cells[$rS].Merge = $true
$cS = $wsExec.Cells["$($k.C1)10"]
$cS.Value = $k.Sub
$cS.Style.Font.Size = 8; $cS.Style.Font.Italic = $true; $cS.Style.Font.Color.SetColor($k.BodyFG)
$cS.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
$cS.Style.Fill.BackgroundColor.SetColor($k.BodyBG)
$cS.Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Center
$cS.Style.VerticalAlignment   = [OfficeOpenXml.Style.ExcelVerticalAlignment]::Center
Set-BordeFino -Worksheet $wsExec -Rango "$($k.C1)8:$($k.C2)10"
}


########################################################################
# 10. GRÁFICOS EJECUTIVOS EN PORTADA (DONAS + RADAR + LÍNEAS)
########################################################################

# Sección 1: Donas de Recursos
Set-BannerTitulo -Worksheet $wsExec -Rango "B12:Q12" -Texto "1. UTILIZACIÓN ACTUAL DE CAPACIDAD Y RECURSOS DE HARDWARE"
$wsExec.Row(12).Height = 24
for ($r = 13; $r -le 24; $r++) { $wsExec.Row($r).Height = 18 }

function Add-DonaEjecutiva {
    param([string]$Nombre, [string]$Titulo, [string]$RangoValores, [string]$RangoCategorias, [int]$FilaAncla, [int]$ColAncla)
    $chart = $wsExec.Drawings.AddChart($Nombre, [OfficeOpenXml.Drawing.Chart.eChartType]::Doughnut)
    $chart.Title.Text = $Titulo
    $chart.Title.Font.Size = 11
    $chart.Title.Font.Bold = $true
    $chart.SetPosition($FilaAncla, 0, $ColAncla, 0)
    $chart.SetSize(310, 205)
    $chart.Legend.Position = [OfficeOpenXml.Drawing.Chart.eLegendPosition]::Bottom
    $serie = $chart.Series.Add($wsAux.Cells[$RangoValores], $wsAux.Cells[$RangoCategorias])
    $serie.Header = $Titulo
}

Add-DonaEjecutiva -Nombre "DonaCPU"   -Titulo "CPU (${cpuUso}%)"           -RangoValores "F2:F3" -RangoCategorias "E2:E3" -FilaAncla 12 -ColAncla 1
Add-DonaEjecutiva -Nombre "DonaRAM"   -Titulo "Memoria RAM (${ramUso}%)"   -RangoValores "I2:I3" -RangoCategorias "H2:H3" -FilaAncla 12 -ColAncla 6
Add-DonaEjecutiva -Nombre "DonaDisco" -Titulo "Disco C: (${discoUso}%)"    -RangoValores "L2:L3" -RangoCategorias "K2:K3" -FilaAncla 12 -ColAncla 12

# Sección 2: Radar de Seguridad + Líneas Históricas
Set-BannerTitulo -Worksheet $wsExec -Rango "B26:I26" -Texto "2. POSTURA DE SEGURIDAD Y HARDENING (8 CONTROLES)"
Set-BannerTitulo -Worksheet $wsExec -Rango "J26:Q26" -Texto "3. EVOLUCIÓN HISTÓRICA DE RENDIMIENTO Y HEALTH SCORE"
$wsExec.Row(26).Height = 24
for ($r = 27; $r -le 40; $r++) { $wsExec.Row($r).Height = 18 }

# Gráfico Radar (Seguridad)
$radar = $wsExec.Drawings.AddChart("RadarSeguridad", [OfficeOpenXml.Drawing.Chart.eChartType]::RadarFilled)
$radar.Title.Text = "Cobertura de Controles de Seguridad (%)"
$radar.Title.Font.Size = 11
$radar.Title.Font.Bold = $true
$radar.SetPosition(26, 0, 1, 0)
$radar.SetSize(480, 240)
$radar.Legend.Remove()
$sRadar = $radar.Series.Add($wsAux.Cells["H8:H$finSec"], $wsAux.Cells["G8:G$finSec"])
$sRadar.Header = "Cumplimiento (%)"

# Gráfico de Líneas (Histórico)
if ($Historico -and @($Historico).Count -gt 0) {
    $lineas = $wsExec.Drawings.AddChart("LineaHistorico", [OfficeOpenXml.Drawing.Chart.eChartType]::LineMarkers)
    $lineas.Title.Text = "Tendencia de Consumo (%) y Health Score"
    $lineas.Title.Font.Size = 11
    $lineas.Title.Font.Bold = $true
    $lineas.YAxis.MinValue = 0
    $lineas.YAxis.MaxValue = 100
    $lineas.SetPosition(26, 0, 9, 0)
    $lineas.SetSize(480, 240)
    $lineas.Legend.Position = [OfficeOpenXml.Drawing.Chart.eLegendPosition]::Bottom

    foreach ($colLetra in @("B", "C", "D", "E")) {
        $sLine = $lineas.Series.Add($wsAux.Cells["${colLetra}8:${colLetra}${finHist}"], $wsAux.Cells["A8:A${finHist}"])
        $sLine.Header = $wsAux.Cells["${colLetra}7"].Text
    }
}



########################################################################
#11. TABLA DINÁMICA DE RIESGOS Y PLAN DE ACCIÓN EN PORTADA
########################################################################

Set-BannerTitulo -Worksheet $wsExec -Rango "B42:Q42" -Texto "4. HALLAZGOS OPERATIVOS, RIESGOS DETECTADOS Y PLAN DE ACCIÓN"
$wsExec.Row(42).Height = 24
$wsExec.Row(43).Height = 22
$HeadersTabla = @(
@{ Rng = "B43:C43"; Txt = "SEVERIDAD" }
@{ Rng = "D43:E43"; Txt = "ÁREA" }
@{ Rng = "F43:K43"; Txt = "HALLAZGO DETECTADO" }
@{ Rng = "L43:Q43"; Txt = "ACCIÓN RECOMENDADA PARA DIRECCIÓN / TI" }
)
foreach ($ht in $HeadersTabla) {
$wsExec.Cells[$ht.Rng].Merge = $true
$c = $wsExec.Cells[$ht.Rng.Split(":")[0]]
$c.Value = $ht.Txt
$c.Style.Font.Size = 9
$c.Style.Font.Bold = $true
$c.Style.Font.Color.SetColor($ColorWhite)
$c.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
$c.Style.Fill.BackgroundColor.SetColor($ColorNavyMed)
$c.Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Center
$c.Style.VerticalAlignment   = [OfficeOpenXml.Style.ExcelVerticalAlignment]::Center
}
Set-BordeFino -Worksheet $wsExec -Rango "B43:Q43"
for ($i = 0; $i -lt $Riesgos.Count; $i++) {
$filaActual = 44 + $i
$wsExec.Row($filaActual).Height = 26
$rItem   = $Riesgos[$i]
$recItem = if ($i -lt $Recomendaciones.Count) { $Recomendaciones[$i].Accion } else { "Monitoreo preventivo estándar" }
$wsExec.Cells["B${filaActual}:C${filaActual}"].Merge = $true
$wsExec.Cells["D${filaActual}:E${filaActual}"].Merge = $true
$wsExec.Cells["F${filaActual}:K${filaActual}"].Merge = $true
$wsExec.Cells["L${filaActual}:Q${filaActual}"].Merge = $true
$sevTxt = $rItem.Severidad.ToUpper()
if ($sevTxt -match "CRITICA|ALTA") {
$fillSev = $ColorRedLight; $fontSev = $ColorRedBG
} elseif ($sevTxt -match "MEDIA|ADVERTENCIA") {
$fillSev = $ColorAmberLight; $fontSev = [System.Drawing.Color]::FromArgb(180, 83, 9)
} else {
$fillSev = $ColorGreenLight; $fontSev = $ColorGreenBG
}
$fillFila = if ($i % 2 -eq 0) { $ColorWhite } else { $ColorSlateBG }
$cSev = $wsExec.Cells["B$filaActual"]
$cSev.Value = $sevTxt
$cSev.Style.Font.Size = 9; $cSev.Style.Font.Bold = $true; $cSev.Style.Font.Color.SetColor($fontSev)
$cSev.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
$cSev.Style.Fill.BackgroundColor.SetColor($fillSev)
$cSev.Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Center
$cSev.Style.VerticalAlignment   = [OfficeOpenXml.Style.ExcelVerticalAlignment]::Center
$cArea = $wsExec.Cells["D$filaActual"]
$cArea.Value = $rItem.Componente
$cArea.Style.Font.Size = 9; $cArea.Style.Font.Bold = $true; $cArea.Style.Font.Color.SetColor($ColorTextDark)
$cArea.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
$cArea.Style.Fill.BackgroundColor.SetColor($fillFila)
$cArea.Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Center
$cArea.Style.VerticalAlignment   = [OfficeOpenXml.Style.ExcelVerticalAlignment]::Center
$cHal = $wsExec.Cells["F$filaActual"]
$cHal.Value = $rItem.Hallazgo
$cHal.Style.Font.Size = 9; $cHal.Style.Font.Color.SetColor($ColorTextDark)
$cHal.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
$cHal.Style.Fill.BackgroundColor.SetColor($fillFila)
$cHal.Style.VerticalAlignment = [OfficeOpenXml.Style.ExcelVerticalAlignment]::Center
$cHal.Style.Indent = 1; $cHal.Style.WrapText = $true
$cAcc = $wsExec.Cells["L$filaActual"]
$cAcc.Value = $recItem
$cAcc.Style.Font.Size = 9; $cAcc.Style.Font.Color.SetColor($ColorTextDark)
$cAcc.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
$cAcc.Style.Fill.BackgroundColor.SetColor($fillFila)
$cAcc.Style.VerticalAlignment = [OfficeOpenXml.Style.ExcelVerticalAlignment]::Center
$cAcc.Style.Indent = 1; $cAcc.Style.WrapText = $true
Set-BordeFino -Worksheet $wsExec -Rango "B${filaActual}:Q${filaActual}"
}


########################################################################
# 12. CONFIGURACION DE IMPRESION / EXPORTACION A PDF (1 PAGINA HORIZONTAL)
########################################################################
$wsExec.PrinterSettings.Orientation = [OfficeOpenXml.eOrientation]::Landscape
$wsExec.PrinterSettings.PaperSize   = [OfficeOpenXml.ePaperSize]::Letter
$wsExec.PrinterSettings.FitToPage   = $true
$wsExec.PrinterSettings.FitToWidth  = 1
$wsExec.PrinterSettings.FitToHeight = 1
########################################################################
#13. GUARDAR Y CERRAR LIBRO
########################################################################

Close-ExcelPackage $Excel
Write-Host ""
Write-Host "Dashboard Ejecutivo generado correctamente:" -ForegroundColor Green
Write-Host $RutaDashboard -ForegroundColor Cyan