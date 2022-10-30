@echo off
if not "%1"=="am_admin" (powershell start -verb runas '%0' am_admin & exit /b)
color a
title Active Office
if exist "%ProgramFiles(x86)%\Microsoft Office\Office16" goto 1
if exist "%ProgramFiles%\Microsoft Office\Office16" goto 1
color c
timeout 1
start https://github.com/a7ecc/Office
msg * /time:5 You must download Office first
exit
:1
cd /d %ProgramFiles(x86)%\Microsoft Office\Office16
cd /d %ProgramFiles%\Microsoft Office\Office16
for /f %%x in ('dir /b ..\root\Licenses16\ProPlus2021VL_KMS*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x"
cscript ospp.vbs /setprt:1688
cscript ospp.vbs /unpkey:6F7TH >nul
cscript ospp.vbs /inpkey:FXYTK-NJJ8C-GB6DW-3DYQT-6F7TH
cscript ospp.vbs /sethst:s8.uk.to
cscript ospp.vbs /act
msg * /time:5 office has been activated