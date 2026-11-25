# 🛡️ Health Check Enterprise & Dashboard Ejecutivo

**Autor:** Victor Manuel Segundo Marquez  
**Tecnologías:** PowerShell | ImportExcel (EPPlus) | Windows Update Agent (COM) | Task Scheduler | CIM/WMI

---

## 📋 Descripción General
Solución integral de automatización en PowerShell diseñada para auditar la salud operativa, postura de seguridad y estado de parcheo de equipos Windows, consolidando los hallazgos en un **Dashboard Ejecutivo en Excel** con ejecución diaria programada y mecanismos de respaldo/rollback.

---

## 📂 Estructura del Repositorio

| Archivo | Función Principal |
| :--- | :--- |
| **`HealthCheck.ps1`** | Recolecta métricas de hardware, red, seguridad (8 controles), servicios, eventos críticos y estado de Windows Update (último KB aplicado y parches pendientes). Exporta a CSV (`Report` e `History`) y registra bitácora (`Log`). |
| **`HealthCheckDashboard.ps1`** | Procesa los datos recolectados, evalúa 11 reglas dinámicas de riesgo y construye un libro de Excel de 6 pestañas con tarjetas KPI, gráficos ejecutivos y formato listo para PDF. |
| **`Run-HealthCheckSuite.ps1`** | Script orquestador que ejecuta secuencialmente el diagnóstico (`HealthCheck.ps1`) y la generación del reporte gerencial (`HealthCheckDashboard.ps1`). |
| **`Configurar-TareaProgramada.ps1`** | Instalador parametrizable que registra o actualiza la ejecución automática diaria en el Programador de Tareas de Windows. |
| **`Rollback.ps1`** | Script de contingencia que crea un respaldo comprimido (`.zip`) de los reportes actuales, limpia el directorio de salida y registra la auditoría de restauración. |

---

## 🚀 Capacidades Principales

### 1. Diagnóstico Integral (`HealthCheck.ps1`)
* **Inventario y Rendimiento:** Hardware, BIOS, CPU, Memoria RAM, almacenamiento lógico (`C:`), salud física del disco (SSD/HDD) y tiempo de actividad (*Uptime*).
* **Seguridad y Hardening (8 Controles):** Microsoft Defender, Protección en Tiempo Real, Firewall de Windows (todos los perfiles), TPM (Presente y Listo), Secure Boot, estado de SMBv1 y cifrado BitLocker.
* **Gestión de Parches (Windows Update):**
  * Identificación del **último parche aplicado** (ID `KB` y fecha de instalación).
  * Conteo automatizado de **actualizaciones pendientes**, clasificadas en **Críticas / Seguridad** y **Opcionales**.
* **Cálculo de Salud:** Generación automática de *Health Score* (0–100) y Semáforo operativo (`VERDE`, `AMARILLO`, `ROJO`).

### 2. Dashboard Ejecutivo (`HealthCheckDashboard.ps1`)
* **Portada Gerencial (`01_ResumenEjecutivo`):** Tarjetas KPI dinámicas, barra ejecutiva de Windows Update, gráficos de dona (CPU/RAM/Disco), gráfico radar de seguridad, tendencia histórica y tabla priorizada de riesgos y remediación.
* **Trazabilidad Técnica:** Pestañas dedicadas para la última ejecución (`02`), histórico (`03`), matriz de cumplimiento (`04`), plan de acción (`05`) y bitácora auditada (`06`).

---

## ⚙️ Requisitos Previos
1. **Windows 10 / 11** o **Windows Server** con PowerShell 5.1 o superior.
2. Ejecutar PowerShell con permisos de **Administrador** (requerido para consultar BitLocker, Secure Boot, eventos del sistema y registrar tareas programadas).
3. Módulo **`ImportExcel`** instalado:
   ```powershell
   Install-Module -Name ImportExcel -Scope CurrentUser -Force
   ```

---

## ▶️ Guía de Ejecución

### Ejecución Manual
1. Abrir **PowerShell como Administrador** y navegar al directorio de scripts:
   ```powershell
   cd C:\Semana9-Automatizacion\Scripts\OFICIALES
   ```
2. Ejecutar la suite completa (diagnóstico + dashboard):
   ```powershell
   .\Run-HealthCheckSuite.ps1
   ```
3. Los archivos resultantes se almacenarán en `C:\Semana9-Automatizacion\Reportes\`.

### Ejecución Automática Programada
1. Editar la variable `$HoraEjecucion` dentro de **`Configurar-TareaProgramada.ps1`** (por defecto `"15:20"`).
2. Ejecutar como Administrador para registrar o actualizar la tarea en Windows:
   ```powershell
   .\Configurar-TareaProgramada.ps1
   ```

---

## 🔄 Procedimiento de Rollback (Reversión)

* **Reversión de entorno (Respaldo ZIP y limpieza de reportes):**
  ```powershell
  .\Rollback.ps1
  ```
* **Reversión de código (Restaurar un script corrupto desde Git):**
  ```powershell
  git log --oneline
  git checkout <hash-version-estable> -- HealthCheck.ps1
  ```
