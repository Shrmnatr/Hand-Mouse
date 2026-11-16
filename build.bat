@echo off
setlocal
cd /d "%~dp0"
title Build Hand Mouse

rem Temporary build files go OUTSIDE OneDrive, so syncing can't lock them
set "WORK=%LOCALAPPDATA%\HandMouseBuild"

echo [1/5] Checking files...
set "MISSING="
for %%F in (hand_mouse_pro.py hand_mouse_app.py hand_landmarker.task "READ ME FIRST.txt" "Gesture Cheat Sheet.txt" "Hand Mouse Far Mode.bat" venv\Scripts\python.exe) do (
    if not exist %%F (
        echo    MISSING: %%~F
        set "MISSING=1"
    )
)
if defined MISSING (
    echo.
    echo    Run hand_mouse_pro.py once first so hand_landmarker.task
    echo    and hand_mouse_settings.json are created.
    goto fail
)
set "PY=venv\Scripts\python.exe"

echo [2/5] Installing PyInstaller...
"%PY%" -m pip install --upgrade pyinstaller pyinstaller-hooks-contrib || goto fail

echo [3/5] Clearing old build files...
if exist "%WORK%" rmdir /s /q "%WORK%"
if exist "%WORK%" (
    echo    Could not remove old build files in:
    echo      %WORK%
    echo    Close any running HandMouse.exe and try again.
    goto fail
)

echo [4/5] Building HandMouse.exe (several minutes, WARNING lines are normal)...
set "ICON="
if exist icon.ico set "ICON=--icon icon.ico"
"%PY%" -m PyInstaller --noconfirm --clean --onefile --name HandMouse %ICON% --collect-all mediapipe --hidden-import tkinter --add-data "hand_landmarker.task;." --workpath "%WORK%\build" --distpath "%WORK%\dist" hand_mouse_app.py || goto fail

echo [5/5] Packaging HandMouse.zip...
if exist release rmdir /s /q release
if exist release (
    echo    Could not replace the old release folder.
    echo    Close HandMouse.exe if it's running from there, then try again.
    goto fail
)
mkdir "release\HandMouse"
copy /y "%WORK%\dist\HandMouse.exe" "release\HandMouse\" >nul || goto fail
copy /y "READ ME FIRST.txt" "release\HandMouse\" >nul || goto fail
copy /y "Gesture Cheat Sheet.txt" "release\HandMouse\" >nul || goto fail
copy /y "Hand Mouse Far Mode.bat" "release\HandMouse\" >nul || goto fail
if exist "hand_mouse_settings.json" copy /y "hand_mouse_settings.json" "release\HandMouse\" >nul
powershell -NoProfile -Command "Compress-Archive -Path 'release\HandMouse' -DestinationPath 'release\HandMouse.zip' -Force" || goto fail

echo.
echo ==================================================
echo   Done!
echo   Test it:   release\HandMouse\HandMouse.exe
echo   Send it:   release\HandMouse.zip
echo ==================================================
explorer release
pause
exit /b 0

:fail
echo.
echo ==================================================
echo   The build stopped. Read the lines above, or
echo   scroll up to the first line containing "ERROR".
echo ==================================================
pause
exit /b 1
