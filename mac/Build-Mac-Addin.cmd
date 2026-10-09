@echo off
rem Run this on WINDOWS with Excel installed (one time, or after changing the mac code).
rem Builds mac\Sankey.xlam from the macOS code files, with the Sankey ribbon inside.
rem Close all Excel windows first. Then commit mac\Sankey.xlam so the macOS installer can use it.
powershell -NoProfile -ExecutionPolicy Bypass -Command "$env:SANKEY_DIR='%~dp0'; $env:SANKEY_TOOLS='%~dp0..\windows'; & ([scriptblock]::Create([IO.File]::ReadAllText('%~dp0..\windows\Install-Sankey.ps1'))) -AddinPath '%~dp0Sankey.xlam' -NoRegister"
if exist "%~dp0Sankey.xlam.bak" del "%~dp0Sankey.xlam.bak"
