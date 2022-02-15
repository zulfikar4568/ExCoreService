echo off
cls
SETLOCAL EnableDelayedExpansion
for /F "tokens=1,2 delims=#" %%a in ('"prompt #$H#$E# & echo on & for %%b in (1) do rem"') do (
  set "DEL=%%a"
)

:ForOrder
@SET ExCoreServiceOrderLocation=ExCoreServiceOrder\bin\Release
:ForOrderBOM
@SET ExCoreServiceOrderBOMLocation=ExCoreServiceOrderBOM\bin\Release
:ForProductMaster
@SET ExCoreServiceProductMasterLocation=ExCoreServiceProductMaster\bin\Release

call :ColorText 0a "                                            Powered by M.Zulfikar Isnaen"
echo(
call :ColorText 0e "                                                Services installer"
echo(
call :ColorText 0b "                                                 Copyright(c) 2022"
echo(
call :ColorText 2F "======================================================================================================================="
echo(
call :ColorText 0a "Searching Folder of ExCoreServiceOrder=" && <nul set /p=%~dp0%ExCoreServiceOrderLocation%
echo(
IF EXIST %~dp0%ExCoreServiceOrderLocation%\NUL (
  call :ColorText 0e "                                             Folder found" && <nul set /p=":)"
  echo(
  cd %~dp0%ExCoreServiceOrderLocation%
  ExCoreServiceOrder.exe install start
  call :ColorText 0a "                                            Install Finished" && <nul set /p=":)"
  echo(
) else (
  call :ColorText 0C "                                             Folder not found" && <nul set /p=":("
  echo(
)
call :ColorText 2F "======================================================================================================================="
echo(
call :ColorText 0a "Searching Folder of ExCoreServiceOrderBOM=" && <nul set /p=%~dp0%ExCoreServiceOrderBOMLocation%
echo(
IF EXIST %~dp0%ExCoreServiceOrderBOMLocation%\NUL (
  call :ColorText 0e "                                             Folder found" && <nul set /p=":)"
  echo(
  cd %~dp0%ExCoreServiceOrderBOMLocation%
  ExCoreServiceOrderBOM.exe install start
  call :ColorText 0a "                                            Install Finished" && <nul set /p=":)"
  echo(
) else (
  call :ColorText 0C "                                             Folder not found" && <nul set /p=":("
  echo(
)
call :ColorText 2F "======================================================================================================================="
echo(
call :ColorText 0a "Searching Folder of ExCoreServiceProductMaster=" && <nul set /p=%~dp0%ExCoreServiceProductMasterLocation%
echo(
IF EXIST %~dp0%ExCoreServiceProductMasterLocation%\NUL (
  call :ColorText 0e "                                             Folder found" && <nul set /p=":)"
  echo(
  cd %~dp0%ExCoreServiceProductMasterLocation%
  ExCoreServiceProductMaster.exe install start
  call :ColorText 0a "                                            Install Finished" && <nul set /p=":)"
  echo(
) else (
  call :ColorText 0C "                                             Folder not found" && <nul set /p=":("
  echo(
)
@PAUSE

goto :eof
:ColorText
echo off
<nul set /p ".=%DEL%" > "%~2"
findstr /v /a:%1 /R "^$" "%~2" nul
del "%~2" > nul 2>&1
goto :eof