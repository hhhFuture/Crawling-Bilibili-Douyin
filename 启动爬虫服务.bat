@echo off
setlocal EnableExtensions
title Bilibili / Douyin Crawler Service

rem Always start from the folder that contains this script.
cd /d "%~dp0"

set "CONDA_ROOT=D:\Miniconda3"
set "ENV_NAME=py310"
set "ENV_DIR=%CONDA_ROOT%\envs\%ENV_NAME%"
set "PYTHON_EXE=%ENV_DIR%\python.exe"
set "APP_FILE=%~dp0bilibili_crawler_app.py"

if not exist "%PYTHON_EXE%" goto :missing_python
if not exist "%APP_FILE%" goto :missing_app

rem Use the py310 environment directly and add its DLL/tool folders to PATH.
set "PATH=%ENV_DIR%;%ENV_DIR%\Library\mingw-w64\bin;%ENV_DIR%\Library\usr\bin;%ENV_DIR%\Library\bin;%ENV_DIR%\Scripts;%ENV_DIR%\bin;%CONDA_ROOT%;%CONDA_ROOT%\Library\mingw-w64\bin;%CONDA_ROOT%\Library\usr\bin;%CONDA_ROOT%\Library\bin;%CONDA_ROOT%\Scripts;%CONDA_ROOT%\bin;%PATH%"
set "CONDA_PREFIX=%ENV_DIR%"
set "CONDA_DEFAULT_ENV=%ENV_NAME%"
set "CONDA_SHLVL=1"

echo ============================================================
echo   Bilibili / Douyin Crawler Service
echo ------------------------------------------------------------
echo   Environment : %ENV_NAME%
echo   Working dir : %CD%
echo   App         : bilibili_crawler_app.py
echo ------------------------------------------------------------
echo   The service is starting. Keep this window open.
echo   Press Ctrl+C to stop the service.
echo ============================================================
echo.

"%PYTHON_EXE%" -m streamlit run "%APP_FILE%"
set "EXIT_CODE=%ERRORLEVEL%"

echo.
if "%EXIT_CODE%"=="0" (
    echo Service stopped.
) else (
    echo Service stopped. Exit code: %EXIT_CODE%
)
echo Press any key to close this window...
pause >nul
exit /b %EXIT_CODE%

:missing_python
echo.
echo [ERROR] Python was not found at:
echo         %PYTHON_EXE%
echo Please check the conda installation and the py310 environment.
echo.
pause
exit /b 1

:missing_app
echo.
echo [ERROR] Application file was not found at:
echo         %APP_FILE%
echo Please keep this script in the same folder as bilibili_crawler_app.py.
echo.
pause
exit /b 1
