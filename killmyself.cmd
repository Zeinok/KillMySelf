@echo off
::KMS Script by Zeinok
::publish with GPLv3
::Load script config.
if not exist killmyself.ini echo The config file killmyself.ini is not found. Creating new template... & call :genConfig & exit /b
for /f "tokens=1,* delims==" %%a in (killmyself.ini) do set %%a=%%b
setlocal EnableDelayedExpansion
cd /d %windir%/system32/
set err=0
fsutil.exe dirty query %systemdrive% > nul
if %ERRORLEVEL%==0 (
    echo Administrator permission ok.
    ) else (
    echo.|set/p"=Error, administrator permission required for this script."
    pause > nul
    exit /b 1
)
echo.
echo Welcome to KMS Installation Script.
:selectAction
if defined AUTO_INSTALL set selection=%AUTO_INSTALL% & goto :gotoSelection
echo.
echo Please select your action.
echo.
if defined KMS_SERVER echo. 1.Windows only
if defined OFFICE_KMS_SERVER echo. 2.Office only
if defined OFFICE_KMS_SERVER echo. 3.Windows + Office
echo.
set/p "selection=Your choice(1,2,3):"
:gotoSelection
set /a selection=!selection!
if !selection! EQU 0 echo Invalid choice.
if !selection! EQU 1 call :WinKMS
if !selection! EQU 2 call :OfficeKMS
if !selection! EQU 3 call :WinKMS & call :OfficeKMS
if !selection! GTR 3 echo Invalid choice.
echo.
if !nopause! EQU 0 pause
goto :EOF


:WinKMS
if NOT defined KMS_SERVER call :notDefined KMS_SERVER & goto :EOF
echo.
echo Checking Windows Version...
for /f "tokens=1,2*" %%i in ('reg query "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion" /v ProductName') do (
    if "%%i"=="ProductName" set "OS=%%~k"
)
if defined OS (
	echo Target product is "%OS%".
	set OS=%OS: =_%
) else (
    echo Failed to fetch product name.
    set err=3
    goto :end
)
echo.
if %nopause%==0 (
    echo.|set/p"=Press any key to start KMS Installation..."
    pause > nul
    echo.
)
if defined OS (
    call :%OS% 2>nul || echo Sorry, no KMS key available for this Windows product. & set err=2 || goto :end
	call :kms 
)
:end
goto :EOF
:kms
echo Setting up KMS Server...
echo.
net stop sppsvc && net start sppsvc
if defined KMS_PORT (
    cscript.exe //nologo slmgr.vbs /skms %KMS_SERVER%:%KMS_PORT% && echo KMS Server set.
    ) else (
    cscript.exe //nologo slmgr.vbs /skms %KMS_SERVER% && echo KMS Server set.
    )
echo Activating Windows...
cscript.exe //nologo slmgr.vbs /ato && echo Done activating Windows.
goto :EOF
:OfficeKMS
if NOT defined OFFICE_KMS_SERVER call :notDefined OFFICE_KMS_SERVER & goto :EOF
echo.
echo Activating Office...
echo.
for /d %%d in ("%ProgramW6432%\Microsoft Office"\Office* "%ProgramFiles(x86)%\Microsoft Office"\Office*) do (
    if exist "%%d\ospp.vbs" (
        cscript.exe //nologo "%%d\ospp.vbs" /remhst
        cscript.exe //nologo "%%d\ospp.vbs" /sethst:%OFFICE_KMS_SERVER%
        if defined OFFICE_KMS_PORT cscript.exe //nologo "%%d\ospp.vbs" /setprt:%OFFICE_KMS_PORT%
        cscript.exe //nologo "%%d\ospp.vbs" /act && echo Done activating Office. || echo Failed to activate Office, is your Office volume license?
    )
)
goto :EOF
:genConfig
echo.;KillMySelf config file.>> killmyself.ini
echo.KMS_SERVER= >> killmyself.ini
echo.KMS_PORT= >> killmyself.ini
echo.OFFICE_KMS_SERVER= >> killmyself.ini
echo.OFFICE_KMS_PORT= >> killmyself.ini
echo.;You can leave port blank if server's KMS port is default. >> killmyself.ini
echo.nopause=0 >> killmyself.ini
echo.;if not 0, exit script immediately after KMS installation. >> killmyself.ini
echo.>> killmyself.ini
echo.AUTO_INSTALL= >> killmyself.ini
echo.;if defined, ignore menu, and install with the selection. >> killmyself.ini
echo.;1=Windows only,2=Office only,3=Windows and Office >> killmyself.ini
echo.>> killmyself.ini
echo Done creating template config. Opening notepad...
start "" notepad.exe killmyself.ini
ping -n 3 127.0.0.1 > nul
goto :EOF
:notDefined
echo %~1 is not defined. Will not begin the KMS configuration.
ping -n 3 127.0.0.1
goto :EOF
::define key
:Windows_10_Pro
cscript.exe //nologo slmgr.vbs /ipk W269N-WFGWX-YVC9B-4J6C9-T83GX
goto :EOF
:Windows_10_Enterprise
cscript.exe //nologo slmgr.vbs /ipk NPPR9-FWDCX-D2C8J-H872K-2YT43
goto :EOF
:Windows_10_Education
cscript.exe //nologo slmgr.vbs /ipk NW6C2-QMPVW-D7KKK-3GKT6-VCFB2
goto :EOF
:Windows_Server_2016_Datacenter
cscript.exe //nologo slmgr.vbs /ipk CB7KF-BWN84-R7R2Y-793K2-8XDDG
goto :EOF
:Windows_Server_2016_Standard
cscript.exe //nologo slmgr.vbs /ipk WC2BQ-8NRM3-FDDYY-2BFGV-KHKQY
goto :EOF
:Windows_Server_2016_Essentials
cscript.exe //nologo slmgr.vbs /ipk JCKRF-N37P4-C2D82-9YXRT-4M63B
goto :EOF
:Windows_8.1_Pro
cscript.exe //nologo slmgr.vbs /ipk GCRJD-8NW9H-F2CDX-CCM8D-9D6T9
goto :EOF
:Windows_8.1_Enterprise
cscript.exe //nologo slmgr.vbs /ipk MHF9N-XY6XB-WVXMC-BTDCT-MKKG7
goto :EOF
:Windows_Server_2012_R2_Server_Standard
cscript.exe //nologo slmgr.vbs /ipk D2N9P-3P6X9-2R39C-7RTCD-MDVJX
goto :EOF
:Windows_Server_2012_R2_Datacenter
cscript.exe //nologo slmgr.vbs /ipk W3GGN-FT8W3-Y4M27-J84CP-Q3VJ9
goto :EOF
:Windows_Server_2012_R2_Essentials
cscript.exe //nologo slmgr.vbs /ipk KNC87-3J2TX-XB4WP-VCPJV-M4FWM
goto :EOF
:Windows_7_Enterprise
cscript.exe //nologo slmgr.vbs /ipk 33PXH-7Y6KF-2VJC9-XBBR8-HVTHH
goto :EOF
:Windows_7_Professional
cscript.exe //nologo slmgr.vbs /ipk FJ82H-XT6CR-J8D7P-XQJJ2-GPDD4
goto :EOF
:Windows_Server_2008_R2_Web
cscript.exe //nologo slmgr.vbs /ipk KNC87-3J2TX-XB4WP-VCPJV-M4FWM
goto :EOF
:Windows_Server_2008_R2_Standard
cscript.exe //nologo slmgr.vbs /ipk YC6KT-GKW9T-YTKYR-T4X34-R7VHC
goto :EOF
:Windows_Server_2008_R2_Enterprise
cscript.exe //nologo slmgr.vbs /ipk 489J6-VHDMP-X63PK-3K798-CPX3Y
goto :EOF
:Windows_Server_2008_R2_Datacenter
cscript.exe //nologo slmgr.vbs /ipk 74YFP-3QFB3-KQT8W-PMXWJ-7M648
goto :EOF
::end of key definition