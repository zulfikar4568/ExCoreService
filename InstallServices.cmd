echo off
cls
SETLOCAL EnableDelayedExpansion
for /F "tokens=1,2 delims=#" %%a in ('"prompt #$H#$E# & echo on & for %%b in (1) do rem"') do (
  set "DEL=%%a"
)

call :ColorText 0a "                                            Powered by M.Zulfikar Isnaen"
echo(
call :ColorText 0e "                                                Services installer"
echo(
call :ColorText 0b "                                                 Copyright(c) 2022"
echo(
call :ColorText 2F "======================================================================================================================="
echo(
:AskForOrder
@SET /P ExCoreServiceOrderLocation="Where's the path Installer File of ExCoreServiceOrder: "
:AskForOrderBOM
@SET /P ExCoreServiceOrderBOMLocation="Where's the path Installer File of ExCoreServiceOrderBOM: "
:AskForProductMaster
@SET /P ExCoreServiceProductMasterLocation="Where's the path Installer File of ExCoreServiceProductMaster: "

call :ColorText 2F "======================================================================================================================="
echo(
call :ColorText 0a "                                   Searching Folder of ExCoreServiceOrder" && <nul set /p="..."
echo(
IF EXIST %ExCoreServiceOrderLocation% (
  call :ColorText 0e "                                             Folder found" && <nul set /p=":)"
  echo(
  cd %ExCoreServiceOrderLocation%
  ExCoreServiceOrder.exe install start
  call :ColorText 0a "                                            Install Finished" && <nul set /p=":)"
  echo(
) else (
  call :ColorText 0C "                                             Folder not found" && <nul set /p=":("
  echo(
)
call :ColorText 2F "======================================================================================================================="
echo(
call :ColorText 0a "                                   Searching Folder of ExCoreServiceOrderBOM" && <nul set /p="..."
echo(
IF EXIST %ExCoreServiceOrderBOMLocation% (
  call :ColorText 0e "                                             Folder found" && <nul set /p=":)"
  echo(
  cd %ExCoreServiceOrderBOMLocation%
  ExCoreServiceOrderBOM.exe install start
  call :ColorText 0a "                                            Install Finished" && <nul set /p=":)"
  echo(
) else (
  call :ColorText 0C "                                             Folder not found" && <nul set /p=":("
  echo(
)
call :ColorText 2F "======================================================================================================================="
echo(
call :ColorText 0a "                                   Searching Folder of ExCoreServiceProductMaster" && <nul set /p="..."
echo(
IF EXIST %ExCoreServiceProductMasterLocation% (
  call :ColorText 0e "                                             Folder found" && <nul set /p=":)"
  echo(
  cd %ExCoreServiceProductMasterLocation%
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