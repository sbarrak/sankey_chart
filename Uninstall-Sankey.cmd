@echo off
rem Runs the uninstaller as text so Windows script-signing / "downloaded file" blocks do not apply.
powershell -NoProfile -ExecutionPolicy Bypass -Command "& ([scriptblock]::Create([IO.File]::ReadAllText('%~dp0Uninstall-Sankey.ps1')))"
