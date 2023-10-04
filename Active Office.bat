@echo off
if not "%1"=="am_admin" (powershell start -verb runas '%0' am_admin & exit /b)
title Active Office
color a
if exist "%ProgramFiles(x86)%\Microsoft Office\Office16" goto 1
if exist "%ProgramFiles%\Microsoft Office\Office16" goto 1
:run
cls
echo Chouse Office Language:
echo 1) Arabic
echo 2) English
set /p office1=num:
if %office1%==1 goto Arabic
if %office1%==2 goto English
goto run
:Arabic
echo Do you want to download offline version?
echo 1) Yes
echo 2) No
set /p office2=num:
if %office2%==1 goto Office-Ar-Offline
if %office2%==2 goto Office-Ar
goto Arabic
:English
echo Do you want to download offline version?
echo 1) Yes
echo 2) No
set /p office3=num:
if %office3%==1 goto Office-En-Offline
if %office3%==2 goto Office-En
goto English
:Office-Ar-Offline
powershell -command "Start-BitsTransfer -Source 'https://github.com/a7ecc/Office/raw/main/AR-Offline.exe' -Destination '%temp%\OfficeSetup.exe'"
goto RunOffice
:Office-Ar
powershell -command "Start-BitsTransfer -Source 'https://github.com/a7ecc/Office/raw/main/AR.exe' -Destination '%temp%\OfficeSetup.exe'"
goto RunOffice
:Office-EN
powershell -command "Start-BitsTransfer -Source 'https://github.com/a7ecc/Office/raw/main/EN.exe' -Destination '%temp%\OfficeSetup.exe'"
goto RunOffice
:Office-EN-Offline
powershell -command "Start-BitsTransfer -Source 'https://github.com/a7ecc/Office/raw/main/EN-Offline.exe' -Destination '%temp%\OfficeSetup.exe'"
goto RunOffice
:RunOffice
cd "%temp%"
start OfficeSetup.exe
exit
:1
cd "%ProgramFiles(x86)%\Microsoft Office\Office16"
cd "%ProgramFiles%\Microsoft Office\Office16"
for /f %%x in ('dir /b ..\root\Licenses16\ProPlus2021VL_KMS*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x"
cscript ospp.vbs /setprt:1688
cscript ospp.vbs /unpkey:6F7TH >nul
cscript ospp.vbs /inpkey:FXYTK-NJJ8C-GB6DW-3DYQT-6F7TH
cscript ospp.vbs /sethst:s8.uk.to
cscript ospp.vbs /act
if exist "%temp%\OfficeSetup.exe" del "%temp%\OfficeSetup.exe"
