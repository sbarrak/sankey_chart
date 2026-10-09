@echo off
rem Runs the installer as text so Windows script-signing / "downloaded file" blocks do not apply.
powershell -NoProfile -ExecutionPolicy Bypass -Command "$env:SANKEY_DIR='%~dp0windows'; $env:SANKEY_TOOLS='%~dp0windows'; & ([scriptblock]::Create([IO.File]::ReadAllText('%~dp0windows\Install-Sankey.ps1')))"
