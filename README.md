# 🛡️ Health Check Enterprise – Automatización Operativa e Infraestructura

| Práctica | Instructor | Autor | Plataforma | Estado |
| :--- | :--- | :--- | :--- | :--- |
| **Semana 9 – Automatización Operativa** | Reyes Hernández René Iván | Víctor Manuel Segundo Márquez | Windows 10 / 11 / Server (PowerShell 5.1+) | Producción |

Solución integral desarrollada en **PowerShell** para la auditoría automatizada de inventario, diagnóstico de salud (*Health Check*), generación de tableros ejecutivos en Excel, gestión de reversión (*Rollback*) con respaldo preventivo y ejecución desatendida en entornos Windows.

---

## 🎯 Resumen Ejecutivo y Capacidades

El sistema transforma la revisión manual de servidores y estaciones de trabajo en un flujo automatizado, reproducible y auditado:

* **Diagnóstico Integral (`HealthCheck.ps1`):** Evalúa hardware, sistema operativo, conectividad de red, postura de seguridad (Windows Defender, Firewall, BitLocker, Secure Boot) y parches pendientes de Windows Update.
* **Semaforización y Health Score:** Calcula un puntaje de salud global (`0–100`) y clasifica cada métrica en `OK`, `ADVERTENCIA` o `CRÍTICO`.
* **Reportería Gerencial (`HealthCheckDashboard.ps1`):** Genera de forma nativa un libro de Excel (`.xlsx`) de 6 pestañas con KPIs, gráficas de utilización y matriz de hallazgos sin requerir Office instalado.
* **Resiliencia y Rollback (`Rollback.ps1`):** Protege los datos mediante respaldos automáticos comprimidos (`.zip`) con retención de los últimos 10 cortes, permitiendo limpieza total, reversión parcial del último registro o restauración completa.
* **Ejecución Desatendida (`Configurar-TareaProgramada.ps1`):** Integra toda la suite al Programador de Tareas de Windows con privilegios elevados.

---

## 📂 Arquitectura del Proyecto

```text
C:\Semana9-Automatizacion\
├── Scripts\OFICIALES\
│   ├── HealthCheck.ps1                # Motor principal de diagnóstico, validaciones y auditoría
│   ├── HealthCheckDashboard.ps1       # Generador de Dashboard Ejecutivo en Excel (.xlsx)
│   ├── Rollback.ps1                   # Herramienta de respaldo preventivo y reversión operativa
│   ├── Run-HealthCheckSuite.ps1       # Script orquestador (ejecución secuencial)
│   ├── Configurar-TareaProgramada.ps1 # Instalador parametrizable en Windows Task Scheduler
│   └── README.md                      # Documentación técnica del repositorio
├── Reportes\
│   ├── HealthCheck_Report.csv         # Último diagnóstico ejecutado
│   ├── HealthCheck_History.csv        # Histórico acumulado de ejecuciones
│   ├── HealthCheck_Log.txt            # Bitácora cronológica de eventos y alertas
│   └── HealthCheck_Dashboard.xlsx     # Reporte gerencial con KPIs y gráficas
└── Backups_Rollback\                  # Respaldos comprimidos (.zip) generados por Rollback.ps1
```

---

## 🛠️ Catálogo de Scripts

| Script | Propósito | Entradas | Salidas |
| :--- | :--- | :--- | :--- |
| **`HealthCheck.ps1`** | Recopila métricas de sistema, seguridad y red con manejo de errores (`try/catch`). | WMI/CIM, Registro, Red, Windows Update | `HealthCheck_Report.csv`, `HealthCheck_History.csv`, `HealthCheck_Log.txt` |
| **`HealthCheckDashboard.ps1`** | Construye el tablero visual gerencial con tablas estilizadas y gráficas. | `HealthCheck_Report.csv`, `HealthCheck_History.csv` | `HealthCheck_Dashboard.xlsx` |
| **`Run-HealthCheckSuite.ps1`** | Orquesta la ejecución secuencial del diagnóstico y la actualización del Excel. | Invoca a ambos scripts principales | Suite completa actualizada |
| **`Rollback.ps1`** | Respalda reportes en `.zip` y revierte cambios operativos de forma controlada. | Archivos en `Reportes\` o `Backups_Rollback\` | `Backup_PreRollback_*.zip` y registro en log |
| **`Configurar-TareaProgramada.ps1`** | Crea/actualiza la tarea diaria en el Programador de Tareas de Windows. | Variable `$HoraEjecucion` (`"HH:mm"`) | Tarea `HealthCheck_Enterprise_Diario` |

---

## ⚙️ Requisitos Previos e Instalación

1. **Sistema Operativo:** Windows 10, Windows 11 o Windows Server.
2. **Permisos:** Consola de **PowerShell 5.1+** ejecutada como **Administrador**.
3. **Dependencia de Excel:** Instalar el módulo `ImportExcel` (solo la primera vez):

```powershell
Install-Module ImportExcel -Scope CurrentUser -Force
```

---

## 🚀 Guía de Ejecución

### Ejecución Completa (Recomendada)
Ejecuta el diagnóstico completo y regenera el tablero ejecutivo en un solo paso:
```powershell
cd C:\Semana9-Automatizacion\Scripts\OFICIALES
.\Run-HealthCheckSuite.ps1
```

### Ejecución Individual por Módulo
```powershell
.\HealthCheck.ps1             # Solo auditoría y generación de CSV/Log
.\HealthCheckDashboard.ps1    # Solo regeneración del libro Excel
```

### Automatización Diaria (Task Scheduler)
Para programar la ejecución diaria desatendida (editable en `$HoraEjecucion`):
```powershell
.\Configurar-TareaProgramada.ps1
```

---

## 🔄 Procedimiento de Rollback y Recuperación

El sistema cuenta con dos niveles de reversión para garantizar continuidad operativa:

### 1. Rollback Operativo de Datos (`Rollback.ps1`)
Antes de modificar o borrar cualquier archivo, el script verifica que `HealthCheck_Dashboard.xlsx` no esté bloqueado por Excel y crea un respaldo automático en `Backups_Rollback\` (manteniendo los **10 respaldos más recientes**):

```powershell
.\Rollback.ps1
```

| Opción | Acción | Caso de Uso |
| :---: | :--- | :--- |
| **1** | **Limpieza Total** | Respalda en `.zip` y vacía `Reportes\` para iniciar un ciclo limpio. |
| **2** | **Rollback Parcial** | Elimina únicamente el último registro añadido a `HealthCheck_History.csv`. |
| **3** | **Restauración desde Backup** | Lista los `.zip` disponibles y restaura íntegramente el corte elegido. |
| **4** | **Cancelar / Salir** | Aborta la operación sin realizar cambios. |

### 2. Rollback de Código Fuente (Git)
Para revertir modificaciones en los scripts a una versión estable previa:
```powershell
git log --oneline
git revert HEAD
```
