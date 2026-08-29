@echo off
setlocal

echo =====================================================================
echo   Building Excel Viewer Pro x64 Assembly Acceleration Suite
echo =====================================================================

cd /d "%~dp0"
set FASM=C:\asm\hdd\FASM.EXE

if not exist "%FASM%" (
    echo Error: FASM.EXE not found at %FASM%
    pause
    exit /b 1
)

echo.
echo [1/2] Compiling xlsx_math64.asm -^> xlsx_math64.dll ...
"%FASM%" xlsx_math64.asm xlsx_math64.dll
if %ERRORLEVEL% neq 0 (
    echo Compilation failed for xlsx_math64.asm!
    pause
    exit /b %ERRORLEVEL%
)

taskkill /F /IM xlsx_gui64.exe >nul 2>&1

echo.
echo [2/2] Compiling xlsx_gui64.asm -^> xlsx_gui64.exe ...
"%FASM%" xlsx_gui64.asm xlsx_gui64.exe
if %ERRORLEVEL% neq 0 (
    echo Compilation failed for xlsx_gui64.asm!
    pause
    exit /b %ERRORLEVEL%
)

echo.
echo =====================================================================
echo   Build Successful! Running Python Integration & Benchmark Tests...
echo =====================================================================
echo.

python test_asm.py

echo.
echo =====================================================================
echo   Finished.
echo   - xlsx_math64.dll : 64-bit SIMD Acceleration Library
echo   - xlsx_gui64.exe  : Standalone Native Windows x64 GUI Math Suite
echo =====================================================================
start xlsx_gui64.exe
pause
