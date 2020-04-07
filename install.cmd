@echo off
pushd %~dp0
cls
::FreeSoftwareServers.com
::https://www.freesoftwareservers.com/display/FREES/Distributing+Macro+via+Add-In+-+Customized+Ribbon+-+Via+Batch

set XLSTART=%AppData%\Microsoft\Outlook
set OFFICEUI=%userprofile%\AppData\Local\Microsoft\Office

set USERDIR=%USERPROFILE:C:\Users\=%

call BatchSubstitute.bat "USERNAME" %USERDIR% Excel_officeUI > Excel.OfficeUI

For %%Z In ("%XLSTART%") Do If "%%~aZ" GEq "-" (GoTo XLSTARTFILE) ELSE (GOTO NEXT)

:XLSTARTFILE
del "%XLSTART%" /s /f /q

IF EXIST "%XLSTART%" GOTO FOUNDXL
md "%XLSTART%"
:FOUNDXL

IF EXIST "%OFFICEUI%" GOTO FOUNDUI
md "%OFFICEUI%"
:FOUNDUI

copy "%~dp0*.OTM" "%XLSTART%"
copy "%~dp0*.OfficeUI" "%OFFICEUI%"

del "%~dp0olkExplorer.OfficeUI" /s /f /q

::pause