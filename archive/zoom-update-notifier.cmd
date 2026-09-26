@ECHO OFF
SETLOCAL
COLOR 17
SET "URL=https://zoom.us/client/latest/ZoomInstallerFull.msi?archType=x64"
SET "UAS=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/129.0.0.0 Safari/537.36"
SET "ZOOM_EXE_PATH=C:\Program Files\Zoom\bin\zoom.exe"

IF NOT EXIST "%ZOOM_EXE_PATH%" (
  ECHO zoom.exe does not exist at "%ZOOM_EXE_PATH%"
  EXIT /B 1
)

FOR /F "tokens=3 delims=/" %%a in ('2^>^&1 CURL.EXE -A "%UAS%" -I -k -L -v "%URL%" ^| FINDSTR /C:"HEAD /prod/"') DO (
    SET "LATEST_VERSION=%%a"
)

SET "ZOOM_EXE_PATH=%ZOOM_EXE_PATH:\=\\%"
REM FOR /F "usebackq delims=^= tokens=2" %%a IN (`WMIC DATAFILE WHERE "NAME='%ZOOM_EXE_PATH%'" GET VERSION /VALUE ^| FINDSTR Version`) DO (
FOR /F "usebackq tokens=*" %%a IN (`powershell -c ^"Get-CimInstance -ClassName CIM_DataFile -Filter \"Name='%ZOOM_EXE_PATH%'\" ^| Select-Object -ExpandProperty Version^"`) DO (
    SET "INSTALLED_VERSION=%%a"
)

ECHO Installed version is %INSTALLED_VERSION%
ECHO Latest version is %LATEST_VERSION%

IF NOT "%LATEST_VERSION%"=="%INSTALLED_VERSION%" (
    ECHO Update required!
    PAUSE
    EXIT /B 2
)
