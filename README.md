# 🛡️ Health Check Enterprise & Dashboard Ejecutivo

**Autor:** Victor Manuel Segundo Marquez  
**Tecnologías:** PowerShell | ImportExcel (EPPlus) | Windows Update Agent (COM) | Task Scheduler | CIM/WMI

---

## 📋 Descripción General
Solución integral de automatización en PowerShell diseñada para auditar la salud operativa, postura de seguridad y estado de parcheo de equipos Windows, consolidando los hallazgos en un **Dashboard Ejecutivo en Excel** con ejecución diaria programada.

---

## 📂 Estructura del Repositorio

| Archivo | Función Principal |
| :--- | :--- |
| **`HealthCheck.ps1`** | Recolecta métricas de hardware, red, seguridad (8 controles), servicios, eventos críticos y estado de Windows Update (último KB aplicado y parches pendientes). Exporta a CSV (`Report` e `History`) y registra bitácora (`Log`). |
| **`HealthCheckDashboard.ps1`** | Procesa los datos recolectados, evalúa 11 reglas dinámicas de riesgo y construye un libro de Excel de 6 pestañas con tarjetas KPI, gráficos ejecutivos y formato listo para PDF. |
| **`Run-HealthCheckSuite.ps1`** | Script orquestador que ejecuta secuencialmente el diagnóstico (`HealthCheck.ps1`) y la generación del reporte gerencial (`HealthCheckDashboard.ps1`). |
| **`Configurar-TareaProgramada.ps1`** | Instalador parametrizable que registra o actualiza la ejecución automática diaria en el Programador de Tareas de Windows. |
| **`Aplicar_Nuevo_Horario.bat`** | Lanzador de un solo clic con elevación automática de privilegios para aplicar cambios de horario sin abrir manualmente la consola. |
| **`Rollback.ps1`** | Script de contingencia para limpiar archivos generados y restaurar el directorio al estado inicial. |

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
2. Ejecutar con permisos de **Administrador** (requerido para BitLocker, Secure Boot, eventos del sistema y registro de tareas programadas).
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

### Ejecución Automática Programada (Sin consola)
1. Abrir **`Configurar-TareaProgramada.ps1`** y editar la variable `$HoraEjecucion` (formato 24h, ej. `"15:50"`).
2. Hacer doble clic sobre **`Aplicar_Nuevo_Horario.bat`** y aceptar el control de cuentas de usuario (**Sí**) para aplicar el nuevo horario automáticamente.

---

## 🔄 Procedimiento de Rollback y Recuperación

### 1. Limpieza de Archivos Generados (`.\Rollback.ps1`)
```powershell
.\Rollback.ps1
```

### 2. Rollback de Código Fuente (Git)
* **Descartar cambios locales no confirmados:**
  ```powershell
  git checkout -- HealthCheck.ps1
  ```
* **Restaurar un script desde una versión estable anterior:**
  ```powershell
  git log --oneline
  git checkout <hash-version-estable> -- HealthCheck.ps1
  ```
