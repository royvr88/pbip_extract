@echo off
:: pushd maps a UNC path (\\wsl.localhost\...) to a temporary drive letter
pushd "%~dp0"

:: Locate Python — try the Windows launcher first, then plain python
where py >nul 2>&1
if %errorlevel%==0 (
    set PYTHON=py
) else (
    where python >nul 2>&1
    if %errorlevel%==0 (
        set PYTHON=python
    ) else (
        echo ERROR: Python not found.
        echo Install Python from https://www.python.org ^(tick "Add to PATH"^).
        popd & pause & exit /b 1
    )
)

echo Using: %PYTHON%
echo.

echo Installing PyInstaller...
%PYTHON% -m pip install pyinstaller --quiet
if %errorlevel% neq 0 (
    echo ERROR: pip install failed.
    popd & pause & exit /b 1
)

echo.
echo Building PBIP Extract.exe...
%PYTHON% -m PyInstaller --onefile --windowed --name "PBIP Extract" pbip_gui.py

echo.
if exist "dist\PBIP Extract.exe" (
    echo Build successful^^!  Find the .exe in the dist\ folder.
) else (
    echo Build may have failed. Check the output above for errors.
)

popd
pause
