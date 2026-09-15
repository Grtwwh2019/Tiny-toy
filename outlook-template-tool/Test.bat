@echo off
setlocal
cd /d "%~dp0"
if exist test-build rmdir /s /q test-build
mkdir test-build
javac -encoding UTF-8 -source 8 -target 8 -d test-build src\mailtool\*.java test\mailtool\Tests.java
if errorlevel 1 goto fail
java -Djava.awt.headless=true -cp test-build mailtool.Tests
if errorlevel 1 goto fail
pause
exit /b 0
:fail
echo Test failed. See output above.
pause
exit /b 1
