@echo off
setlocal

set "SCRIPT_DIR=%~dp0"
set "PS_SCRIPT=%SCRIPT_DIR%backup_project.ps1"

if not exist "%PS_SCRIPT%" (
  echo ERROR: Cannot find backup_project.ps1 in:
  echo   %SCRIPT_DIR%
  echo.
  pause
  exit /b 1
)

:menu
cls
echo ==================================
echo Research Project Backup Tool
echo ==================================
echo Project root:
echo   %SCRIPT_DIR%
echo.
echo 1. Incremental sync backup
echo 2. Dated archive backup
echo 3. Pull newer files from cloud backup
echo 4. Reconfigure workspace roots
echo 5. Exit
echo.
set /p CHOICE=Choose an option [1-5]: 

if "%CHOICE%"=="1" goto incremental
if "%CHOICE%"=="2" goto archive
if "%CHOICE%"=="3" goto pull
if "%CHOICE%"=="4" goto configure
if "%CHOICE%"=="5" goto end

echo.
echo Invalid choice.
pause
goto menu

:incremental
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%PS_SCRIPT%" -Mode Incremental
goto done

:archive
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%PS_SCRIPT%" -Mode Archive
goto done

:pull
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%PS_SCRIPT%" -Mode Pull
goto done

:configure
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%PS_SCRIPT%" -ConfigureOnly
goto done

:done
echo.
pause
goto menu

:end
endlocal
exit /b 0
