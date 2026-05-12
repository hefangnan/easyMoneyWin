@echo off
setlocal
set "ROOT=%~dp0"
set "SRC=%ROOT%easy_money_input.cpp"
set "OUT=%ROOT%easy_money_input.dll"
set "LOCAL_GPP=%ROOT%.tools\w64devkit\bin\g++.exe"

if exist "%LOCAL_GPP%" (
  set "PATH=%ROOT%.tools\w64devkit\bin;%PATH%"
  "%LOCAL_GPP%" -O3 -DNDEBUG -shared -static-libgcc -static-libstdc++ -o "%OUT%" "%SRC%" -luser32
  exit /b %errorlevel%
)

where cl.exe >nul 2>nul
if %errorlevel%==0 (
  cl.exe /nologo /O2 /EHsc /LD "%SRC%" /Fe:"%OUT%" /link user32.lib
  exit /b %errorlevel%
)

set "VSWHERE=%ProgramFiles(x86)%\Microsoft Visual Studio\Installer\vswhere.exe"
if exist "%VSWHERE%" (
  for /f "usebackq delims=" %%I in (`"%VSWHERE%" -latest -products * -requires Microsoft.VisualStudio.Component.VC.Tools.x86.x64 -property installationPath`) do set "VSINSTALL=%%I"
  if defined VSINSTALL (
    if exist "%VSINSTALL%\VC\Auxiliary\Build\vcvars64.bat" (
      call "%VSINSTALL%\VC\Auxiliary\Build\vcvars64.bat" >nul
      cl.exe /nologo /O2 /EHsc /LD "%SRC%" /Fe:"%OUT%" /link user32.lib
      exit /b %errorlevel%
    )
  )
)

where g++.exe >nul 2>nul
if %errorlevel%==0 (
  g++.exe -O3 -DNDEBUG -shared -static-libgcc -static-libstdc++ -o "%OUT%" "%SRC%" -luser32
  exit /b %errorlevel%
)

where clang++.exe >nul 2>nul
if %errorlevel%==0 (
  clang++.exe -O3 -DNDEBUG -shared -o "%OUT%" "%SRC%" -luser32
  exit /b %errorlevel%
)

echo Error: C++ compiler not found.
echo Install Visual Studio Build Tools and run from Developer PowerShell, or install MinGW-w64 g++.
exit /b 1
