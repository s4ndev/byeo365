@echo off
setlocal DisableDelayedExpansion


::=========================================================================================================================================

::  Set environment variables

setlocal EnableExtensions
setlocal DisableDelayedExpansion

set "PathExt=.COM;.EXE;.BAT;.CMD;.VBS;.VBE;.JS;.JSE;.WSF;.WSH;.MSC"

set "SysPath=%SystemRoot%\System32"
set "Path=%SystemRoot%\System32;%SystemRoot%;%SystemRoot%\System32\Wbem;%SystemRoot%\System32\WindowsPowerShell\v1.0\"
if exist "%SystemRoot%\Sysnative\reg.exe" (
set "SysPath=%SystemRoot%\Sysnative"
set "Path=%SystemRoot%\Sysnative;%SystemRoot%;%SystemRoot%\Sysnative\Wbem;%SystemRoot%\Sysnative\WindowsPowerShell\v1.0\;%Path%"
)

set "ComSpec=%SysPath%\cmd.exe"
set "PSModulePath=%ProgramFiles%\WindowsPowerShell\Modules;%SysPath%\WindowsPowerShell\v1.0\Modules"

::========================================================================================================================================


cls

echo.
echo ========================================================================
echo ------------------- BYE0365 - Office/M365 Purge Tool -------------------
echo ========================================================================
echo.
echo.
echo           1: Quick (Clears credentials, restore defaults)
echo           2: Full (Quick + reinstall O365)
echo.
echo.
set Choice=
set /p Choice=""
echo.
echo Please input which choice to perform (1/2).

if '%Choice%'=='1' goto quick
if '%Choice%'=='2' goto full


:quick
curl -O https://raw.githubusercontent.com/s4ndev/byeo365/refs/heads/main/assets/scripts.zip
powershell.exe -nologo -noprofile -command "& { Add-Type -A 'System.IO.Compression.FileSystem'; [IO.Compression.ZipFile]::ExtractToDirectory('scripts.zip', 'scripts'); }"

del "%localappdata%/Microsoft/OneAuth"
del "%localappdata%/Microsoft/IdentityCache"

cd scripts
cscript OLicenseCleanup.vbs
call "WPJCleanUp/WPJCleanup.cmd"

"C:\Program Files\Microsoft Office 15\ClientX64\OfficeClickToRun.exe" scenario=Repair platform=x64 culture=en-us RepairType=QuickRepair DisplayLevel=True

:full
pause
cls


