@echo off
powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "Start-Process powershell.exe -ArgumentList '-NoProfile -ExecutionPolicy Bypass -NoExit -File \"C:\Semana9-Automatizacion\Scripts\OFICIALES\Configurar-TareaProgramada.ps1\"' -Verb RunAs"
