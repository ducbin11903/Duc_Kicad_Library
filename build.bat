@echo off
REM ===========================================================
REM  KiCad Library Manager - one-click build
REM  Put this file next to kicad_lib_manager.py and logo.png
REM ===========================================================
setlocal

echo.
echo [1/4] Installing build requirements...
python -m pip install --quiet --upgrade --no-warn-script-location pyinstaller pillow openpyxl
if errorlevel 1 goto :error

echo.
echo [2/4] Creating icon from logo.png...
if exist logo.png (
    python make_icon.py logo.png
) else (
    echo     logo.png not found - the exe will use the default icon
)

echo.
echo [3/4] Building the executable...
REM  Called through "python -m" so it works even when the Scripts
REM  folder is not on PATH.
if exist app.ico (
    python -m PyInstaller --onefile --windowed ^
        --name "KiCad_Lib_Manager" ^
        --icon "app.ico" ^
        --add-data "app.ico;." ^
        --add-data "app_logo.png;." ^
        --version-file "version_info.txt" ^
        --clean --noconfirm ^
        kicad_lib_manager.py
) else (
    python -m PyInstaller --onefile --windowed ^
        --name "KiCad_Lib_Manager" ^
        --version-file "version_info.txt" ^
        --clean --noconfirm ^
        kicad_lib_manager.py
)
if errorlevel 1 goto :error

echo.
echo [4/4] Cleaning up...
rmdir /s /q build 2>nul
del /q "KiCad_Lib_Manager.spec" 2>nul

echo.
echo ===========================================================
echo  DONE
echo  Your program:  dist\KiCad_Lib_Manager.exe
echo.
echo  Copy the exe into your KiCad library folder, or run it
echo  anywhere and use CHANGE to point it at the folder.
echo ===========================================================
echo.
pause
exit /b 0

:error
echo.
echo Build failed. Read the messages above.
pause
exit /b 1
