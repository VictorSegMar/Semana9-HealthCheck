# 🛡️ Health Check Enterprise & Dashboard Ejecutivo

**Autor:** Victor Manuel Segundo Marquez  
**Tecnologías:** PowerShell | ImportExcel (EPPlus) | Windows Update Agent (COM) | CIM/WMI

---

## 📋 Descripción General
Solución integral de automatización en PowerShell diseñada para auditar la salud operativa, postura de seguridad y estado de parcheo de equipos Windows, consolidando los hallazgos en un **Dashboard Ejecutivo en Excel** listo para presentación gerencial o auditoría TI.

---

## 📂 Estructura del Repositorio

| Archivo | Función Principal |
| :--- | :--- |
| **`HealthCheck.ps1`** | Recolecta métricas de hardware, red, seguridad, servicios, eventos críticos y estado de Windows Update. Exporta resultados a CSV (`Report` e `History`) y genera una bitácora (`Log`). |
| **`HealthCheckDashboard.ps1`** | Procesa los datos recolectados, evalúa controles de seguridad y riesgos dinámicos, y construye un libro de Excel de 6 pestañas con gráficos ejecutivos y formato listo para PDF. |
| **`Rollback.ps1`** | Script de contingencia para limpiar reportes generados y restaurar el entorno operativo a su estado inicial. |

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
* **Portada Gerencial (`01_ResumenEjecutivo`):** Tarjetas KPI dinámicas, barra de estado de Windows Update, gráficos de dona (CPU/RAM/Disco), gráfico radar de seguridad, tendencia histórica y tabla priorizada de riesgos y remediación.
* **Trazabilidad Técnica:** Pestañas dedicadas para la última ejecución (`02`), histórico (`03`), matriz de cumplimiento (`04`), plan de acción (`05`) y bitácora auditada (`06`).

---

## ⚙️ Requisitos Previos
1. **Windows 10 / 11** o **Windows Server** con PowerShell 5.1 o superior.
2. Ejecutar PowerShell con permisos de **Administrador** (requerido para consultar BitLocker, Secure Boot y eventos del sistema).
3. Módulo **`ImportExcel`** instalado:
   ```powershell
   Install-Module -Name ImportExcel -Scope CurrentUser -Force
   ```

---

## ▶️ Guía de Ejecución

1. Abrir **PowerShell como Administrador** y navegar al directorio de scripts:
   ```powershell
   cd C:\Semana9-Automatizacion\Scripts\OFICIALES
   ```
2. Ejecutar el diagnóstico del sistema:
   ```powershell
   .\HealthCheck.ps1
   ```
3. Generar el Dashboard Ejecutivo en Excel:
   ```powershell
   .\HealthCheckDashboard.ps1
   ```
4. Los archivos resultantes se almacenarán automáticamente en `C:\Semana9-Automatizacion\Reportes\`.

---

## 🔄 Procedimiento de Rollback (Reversión)

* **Reversión de entorno (Archivos generados):**
  ```powershell
  .\Rollback.ps1
  ```
* **Reversión de código (Control de versiones con Git):**
  ```powershell
  git log --oneline
  git revert HEAD
  ```
