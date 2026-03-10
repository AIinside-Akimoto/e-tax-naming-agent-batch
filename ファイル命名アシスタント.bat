@echo off
setlocal

REM =========================================================
REM e-tax-naming-agent-batch launcher
REM
REM Usage:
REM   ファイル命名アシスタント.bat "InputFolder" "OutputFolder" "ApiBaseUrl" "ApiKey" [Timeout] [CopyNonRenamedFiles] [LogFilePath]
REM
REM Example:
REM   ファイル命名アシスタント.bat "C:\work\input" "C:\work\output" "https://example.contoso.com" "YOUR_API_KEY"
REM   ファイル命名アシスタント.bat "C:\work\input" "C:\work\output" "https://example.contoso.com" "YOUR_API_KEY" 900 true "C:\work\logs\rename.log"
REM =========================================================

if "%~1"=="" goto :usage
if "%~2"=="" goto :usage
if "%~3"=="" goto :usage
if "%~4"=="" goto :usage

set "INPUT_FOLDER=%~1"
set "OUTPUT_FOLDER=%~2"
set "API_BASE_URL=%~3"
set "API_KEY=%~4"
set "TIMEOUT=%~5"
set "COPY_NON_RENAMED=%~6"
set "LOG_FILE_PATH=%~7"

if "%TIMEOUT%"=="" set "TIMEOUT=600"

set "SCRIPT_DIR=%~dp0"
set "PS1=%SCRIPT_DIR%Invoke-NamingAgentBatch.ps1"

if not exist "%PS1%" (
  echo [ERROR] PowerShell script not found: "%PS1%"
  exit /b 1
)

set "COPY_ARG="
if /I "%COPY_NON_RENAMED%"=="true" set "COPY_ARG=-CopyNonRenamedFiles"

set "LOG_ARG="
if not "%LOG_FILE_PATH%"=="" set "LOG_ARG=-LogFilePath \"%LOG_FILE_PATH%\""

powershell -NoProfile -ExecutionPolicy Bypass -File "%PS1%" ^
  -InputFolder "%INPUT_FOLDER%" ^
  -OutputFolder "%OUTPUT_FOLDER%" ^
  -ApiBaseUrl "%API_BASE_URL%" ^
  -ApiKey "%API_KEY%" ^
  -Timeout %TIMEOUT% ^
  %COPY_ARG% ^
  %LOG_ARG% ^
  -Verbose

set "EXIT_CODE=%ERRORLEVEL%"
if not "%EXIT_CODE%"=="0" (
  echo [ERROR] Failed. ExitCode=%EXIT_CODE%
  exit /b %EXIT_CODE%
)

echo [INFO] Completed successfully.
exit /b 0

:usage
echo Usage:
echo   ファイル命名アシスタント.bat "InputFolder" "OutputFolder" "ApiBaseUrl" "ApiKey" [Timeout] [CopyNonRenamedFiles] [LogFilePath]
echo.
echo Example:
echo   ファイル命名アシスタント.bat "C:\work\input" "C:\work\output" "https://example.contoso.com" "YOUR_API_KEY"
echo   ファイル命名アシスタント.bat "C:\work\input" "C:\work\output" "https://example.contoso.com" "YOUR_API_KEY" 900 true "C:\work\logs\rename.log"
exit /b 1
