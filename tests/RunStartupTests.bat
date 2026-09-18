@echo off
setlocal
cd /d "%~dp0.."
call "C:\Program Files (x86)\Embarcadero\Studio\18.0\bin\rsvars.bat"
if errorlevel 1 exit /b 1
if not exist "build\startup-audit" mkdir "build\startup-audit"
"C:\Windows\Microsoft.NET\Framework\v4.0.30319\MSBuild.exe" SIAI.dproj /t:Build /p:Config=Release /p:Platform=Win32 /p:MainSource=tests\StartupSmoke.dpr /p:DCC_ForceExecute=true /p:DCC_ExeOutput=.\build\startup-audit\tests /p:DCC_DcuOutput=.\build\startup-audit\tests\dcu /v:minimal /nologo > "build\startup-audit\tests-build.log" 2>&1
if errorlevel 1 (
  type "build\startup-audit\tests-build.log"
  exit /b 1
)
"build\startup-audit\tests\StartupSmoke.exe" > "build\startup-audit\tests-run.log" 2>&1
set "SIAI_TEST_RESULT=%ERRORLEVEL%"
type "build\startup-audit\tests-run.log"
exit /b %SIAI_TEST_RESULT%
