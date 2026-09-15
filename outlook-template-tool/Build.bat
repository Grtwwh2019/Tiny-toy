@echo off
setlocal
cd /d "%~dp0"
if exist build rmdir /s /q build
mkdir build
javac -encoding UTF-8 -source 8 -target 8 -d build src\mailtool\*.java
if errorlevel 1 goto fail
if exist OutlookTemplateTool.jar del /q OutlookTemplateTool.jar
jar cfe OutlookTemplateTool.jar mailtool.App -C build mailtool
if errorlevel 1 goto fail
echo Build complete: OutlookTemplateTool.jar
exit /b 0
:fail
echo Build failed. A Java JDK 8 or newer is required.
pause
exit /b 1
