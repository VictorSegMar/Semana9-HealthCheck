\\$readme = @'

\# 🛡️ Health Check Enterprise – Automatización Operativa e Infraestructura



Solución integral desarrollada en \*\*PowerShell\*\* para la auditoría automatizada de inventario, diagnóstico de salud (\*Health Check\*), generación de tableros ejecutivos en Excel, gestión de reversión (\*Rollback\*) y ejecución desatendida en entornos Windows.



> \*\*Práctica Semana 9 – Automatización Operativa\*\*  

> \*\*Instructor:\*\* Reyes Hernández René Iván  

> \*\*Autor:\*\* Víctor Manuel Segundo Márquez  



\---



\## 📋 Arquitectura del Proyecto



```text

C:\\Semana9-Automatizacion\\

├── Scripts\\OFICIALES\\

│   ├── HealthCheck.ps1                # Motor principal de diagnóstico, validaciones y auditoría

│   ├── HealthCheckDashboard.ps1       # Generador de Dashboard Ejecutivo en Excel (.xlsx)

│   ├── Rollback.ps1                   # Herramienta de respaldo preventivo y reversión operativa

│   ├── Run-HealthCheckSuite.ps1       # Script orquestador (ejecución secuencial)

│   ├── Configurar-TareaProgramada.ps1 # Instalador parametrizable en Windows Task Scheduler

│   └── README.md                      # Documentación técnica del repositorio

├── Reportes\\

│   ├── HealthCheck\_Report.csv         # Último diagnóstico ejecutado

│   ├── HealthCheck\_History.csv        # Histórico acumulado de ejecuciones

│   ├── HealthCheck\_Log.txt            # Bitácora cronológica de eventos y alertas

│   └── HealthCheck\_Dashboard.xlsx     # Reporte gerencial con KPIs y gráficas

└── Backups\_Rollback\\                  # Respaldos comprimidos (.zip) generados por Rollback.ps1

```



\---



\## ⚙️ Requisitos Previos



\* \*\*Sistema Operativo:\*\* Windows 10 / Windows 11 / Windows Server.

\* \*\*PowerShell:\*\* Versión 5.1 o superior (ejecutado con privilegios de \*\*Administrador\*\*).

\* \*\*Módulo de Excel:\*\* `ImportExcel` (para la generación nativa de `.xlsx` sin requerir Microsoft Excel instalado).



```powershell

Install-Module ImportExcel -Scope CurrentUser -Force

```



\---



\## 🚀 Descripción de Componentes y Uso



\### 1. Diagnóstico y Monitoreo (`HealthCheck.ps1`)

Recopila métricas de hardware, sistema operativo, red, seguridad (Windows Defender, Firewall, BitLocker, Secure Boot) y parches pendientes de Windows Update. Incorpora manejo de excepciones (`try/catch`), cálculo de \*Health Score\* (0–100) y semaforización.

```powershell

.\\HealthCheck.ps1

```



\### 2. Tablero Gerencial (`HealthCheckDashboard.ps1`)

Procesa los archivos `.csv` y `.txt` generados para construir un libro de Excel de 6 pestañas con KPIs, gráficas de utilización de recursos y matriz de hallazgos.

```powershell

.\\HealthCheckDashboard.ps1

```



\### 3. Orquestación Completa (`Run-HealthCheckSuite.ps1`)

Ejecuta de forma secuencial el diagnóstico (`HealthCheck.ps1`) y actualiza inmediatamente el tablero visual (`HealthCheckDashboard.ps1`).

```powershell

.\\Run-HealthCheckSuite.ps1

```



\### 4. Programación de Tarea Diaria (`Configurar-TareaProgramada.ps1`)

Registra la suite en el \*\*Programador de Tareas de Windows\*\* (`HealthCheck\_Enterprise\_Diario`) para su ejecución diaria desatendida con privilegios máximos. Permite modificar la variable `$HoraEjecucion` (formato 24h `"HH:mm"`) según las necesidades operativas.

```powershell

.\\Configurar-TareaProgramada.ps1

```



\---



\## 🔄 Procedimiento de Rollback y Recuperación



El proyecto cuenta con dos niveles de reversión:



\### A. Rollback Operativo de Datos (`Rollback.ps1`)

Verifica que los archivos no estén bloqueados por otros procesos, genera un respaldo comprimido `.zip` (conservando los últimos 10 backups) y permite elegir entre:

1\. \*\*Limpieza Total:\*\* Elimina los reportes actuales tras respaldarlos.

2\. \*\*Rollback Parcial:\*\* Elimina únicamente el último registro añadido a `HealthCheck\_History.csv`.

3\. \*\*Restauración desde Backup:\*\* Recupera el estado anterior desde un archivo `.zip`.

```powershell

.\\Rollback.ps1

```



\### B. Rollback de Código Fuente (Git)

Para revertir cambios en los scripts a una versión estable anterior:

```powershell

git log --oneline

git revert HEAD

```

'@



$readme | Out-File -FilePath "C:\\Semana9-Automatizacion\\Scripts\\OFICIALES\\README.md" -Encoding utf8



