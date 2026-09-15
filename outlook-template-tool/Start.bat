@echo off
setlocal
cd /d "%~dp0"
if exist "runtime\bin\java.exe" (
  "runtime\bin\java.exe" -Dsun.net.http.retryPost=false -jar "OutlookTemplateTool.jar"
) else (
  java -version >nul 2>&1
  if errorlevel 1 (
    echo Java was not found. Install a supported Java 8+ runtime, or ask IT for help.
    echo See the included user guide.
    pause
    exit /b 1
  )
  java -Dsun.net.http.retryPost=false -jar "OutlookTemplateTool.jar"
)
if errorlevel 1 pause
endlocal
