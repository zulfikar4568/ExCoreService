echo off
cls
SETLOCAL EnableDelayedExpansion
for /F "tokens=1,2 delims=#" %%a in ('"prompt #$H#$E# & echo on & for %%b in (1) do rem"') do (
  set "DEL=%%a"
)

:ForOrder
REM Add Custom Folder in here if you want customize for example: Deployment\bin\Release
@SET ExCoreServiceOrderLocation=

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
  ExCoreServiceOrder.exe uninstall
  call :ColorText 0a "                                            Uninstall Finished" && <nul set /p=":)"
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