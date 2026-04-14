@echo off
setlocal EnableExtensions EnableDelayedExpansion

rem ============================================
rem ファイル命名アシスタント Leapnet API 実行BAT
rem ============================================

rem ===== 文字コード設定（必要なら有効化）=====
rem chcp 65001 >nul

rem ===== スクリプトパス =====
set "SCRIPT_DIR=%~dp0"
set "PS1=%SCRIPT_DIR%Invoke-NamingAgentBatch.ps1"

REM =========================================================
REM ファイル命名アシスタント 実行BAT
REM
REM 引数:
REM   1  入力フォルダー
REM   2  出力フォルダー
REM   3  APIベースURL
REM   4  APIキー
REM   5  タイムアウト秒（省略可）
REM   6  元ファイルコピー(true/false)（省略可）
REM   7  ログファイルパス（省略可）
REM   8  Parallelism 並列数（省略可）
REM   9  OrganizeSourceFilesAfterCopy(true/false)（省略可）
REM   10 OriginalFolder (省略可)
REM   11 MappingCsvPath (省略可)
REM
REM Parallelism の決定ルール:
REM   - 第8引数が指定されていればその値を使用
REM   - 未指定なら CPUコア数をもとに自動決定
REM   - 自動計算式: NUMBER_OF_PROCESSORS - 1
REM   - ただし最低 2、最大 6 に補正
REM
REM OrganizeSourceFilesAfterCopy:
REM   - Default is true
REM   - true / 1 / yes を指定すると有効
REM   - false / 0 / no disables it
REM   - コピーできた元ファイルは入力フォルダーから削除
REM   - コピーできなかったファイルは「対象外だったフォルダー」へ移動
REM
REM 例:
REM   ファイル命名アシスタント.bat
REM      "C:\work\input"
REM      "C:\work\output"
REM      "https://example.contoso.com"
REM      "YOUR_API_KEY"
REM
REM   ファイル命名アシスタント.bat
REM      "C:\work\input"
REM      "C:\work\output"
REM      "https://example.contoso.com"
REM      "YOUR_API_KEY"
REM      900
REM      true
REM      "C:\work\logs\rename.log"
REM      5
REM      true
REM      "C:\work\Original"
REM      "C:\work\logs\mapping.csv"
REM =========================================================

rem ===== 引数 =====
set "INPUT_FOLDER=%~1"
set "OUTPUT_FOLDER=%~2"
set "API_BASE_URL=%~3"
set "API_KEY=%~4"
set "TIMEOUT=%~5"
set "COPY_ORIGINAL=%~6"
set "LOG_FILE_PATH=%~7"
set "PARALLELISM=%~8"
set "ORGANIZE_SOURCE=%~9"
shift
shift
shift
shift
shift
shift
shift
shift
shift
set "ORIGINAL_FOLDER=%~1"
set "MAPPING_CSV_PATH=%~2"

rem ===== デフォルト値 =====
if "%TIMEOUT%"=="" set "TIMEOUT=600"
if "%COPY_ORIGINAL%"=="" set "COPY_ORIGINAL=false"
if "%ORGANIZE_SOURCE%"=="" set "ORGANIZE_SOURCE=true"

rem ===== Parallelism 自動決定 =====
if "%PARALLELISM%"=="" (
  set "CPU_COUNT=%NUMBER_OF_PROCESSORS%"
  if "!CPU_COUNT!"=="" set "CPU_COUNT=4"

  set /a PARALLELISM=!CPU_COUNT!-1 2>nul
  if errorlevel 1 set "PARALLELISM=3"

  if !PARALLELISM! LSS 2 set "PARALLELISM=2"
  if !PARALLELISM! GTR 6 set "PARALLELISM=6"

  set "PARALLELISM_SOURCE=auto"
) else (
  set /a PARALLELISM=%PARALLELISM% 2>nul
  if errorlevel 1 (
    echo [ERROR] Parallelism は数値で指定してください: "%~8"
    exit /b 1
  )
  if !PARALLELISM! LSS 1 (
    echo [ERROR] Parallelism は 1 以上で指定してください: "!PARALLELISM!"
    exit /b 1
  )
  set "PARALLELISM_SOURCE=manual"
)

rem ===== normalize OrganizeSourceFilesAfterCopy =====
if /I "%ORGANIZE_SOURCE%"=="1"   set "ORGANIZE_SOURCE=true"
if /I "%ORGANIZE_SOURCE%"=="yes" set "ORGANIZE_SOURCE=true"
if /I "%ORGANIZE_SOURCE%"=="0"   set "ORGANIZE_SOURCE=false"
if /I "%ORGANIZE_SOURCE%"=="no"  set "ORGANIZE_SOURCE=false"

if /I not "%ORGANIZE_SOURCE%"=="true" if /I not "%ORGANIZE_SOURCE%"=="false" (
  echo [ERROR] OrganizeSourceFilesAfterCopy は true / false ^(1 / 0 / yes / no^) で指定してください: "%~9"
  exit /b 1
)

set "ORGANIZE_SOURCE_ARG=-OrganizeSourceFilesAfterCopy:$true"
if /I "%ORGANIZE_SOURCE%"=="false" set "ORGANIZE_SOURCE_ARG=-OrganizeSourceFilesAfterCopy:$false"

rem ===== 必須引数チェック =====
if "%INPUT_FOLDER%"=="" goto :usage
if "%OUTPUT_FOLDER%"=="" goto :usage
if "%API_BASE_URL%"=="" goto :usage
if "%API_KEY%"=="" goto :usage

rem ===== PowerShell スクリプト存在確認 =====
if not exist "%PS1%" (
  echo [ERROR] PowerShellスクリプトが見つかりません: "%PS1%"
  exit /b 1
)

rem ===== 入力フォルダー確認 =====
if not exist "%INPUT_FOLDER%" (
  echo [ERROR] 入力フォルダーが存在しません: "%INPUT_FOLDER%"
  exit /b 1
)

rem ===== 出力フォルダー作成 =====
if not exist "%OUTPUT_FOLDER%" (
  mkdir "%OUTPUT_FOLDER%" 2>nul
  if errorlevel 1 (
    echo [ERROR] 出力フォルダー作成に失敗しました: "%OUTPUT_FOLDER%"
    exit /b 1
  )
)

rem ===== CopyOriginal オプション作成 =====
set "COPY_ARG="
if /I "%COPY_ORIGINAL%"=="true" set "COPY_ARG=-CopyNonRenamedFiles"
if /I "%COPY_ORIGINAL%"=="1"    set "COPY_ARG=-CopyNonRenamedFiles"
if /I "%COPY_ORIGINAL%"=="yes"  set "COPY_ARG=-CopyNonRenamedFiles"

rem ===== OriginalFolder / MappingCsvPath オプション作成 =====
set "ORIGINAL_FOLDER_ARG="
if not "%ORIGINAL_FOLDER%"=="" set "ORIGINAL_FOLDER_ARG=-OriginalFolder \"%ORIGINAL_FOLDER%\""

set "MAPPING_CSV_ARG="
if not "%MAPPING_CSV_PATH%"=="" set "MAPPING_CSV_ARG=-MappingCsvPath \"%MAPPING_CSV_PATH%\""

echo [START]
echo [INFO] PowerShellScript = "%PS1%"
echo [INFO] InputFolder      = "%INPUT_FOLDER%"
echo [INFO] OutputFolder     = "%OUTPUT_FOLDER%"
echo [INFO] ApiBaseUrl       = "%API_BASE_URL%"
echo [INFO] TimeoutSeconds   = %TIMEOUT%
echo [INFO] CopyOriginal     = %COPY_ORIGINAL%
echo [INFO] OrganizeSource   = %ORGANIZE_SOURCE%
echo [INFO] CpuCount         = !NUMBER_OF_PROCESSORS!
echo [INFO] Parallelism      = !PARALLELISM!
echo [INFO] ParallelismMode  = !PARALLELISM_SOURCE!
if not "%LOG_FILE_PATH%"=="" echo [INFO] LogFilePath      = "%LOG_FILE_PATH%"
if not "%ORIGINAL_FOLDER%"=="" echo [INFO] OriginalFolder   = "%ORIGINAL_FOLDER%"
if not "%MAPPING_CSV_PATH%"=="" echo [INFO] MappingCsvPath   = "%MAPPING_CSV_PATH%"

rem ===== PowerShell 実行 =====
if not "%LOG_FILE_PATH%"=="" (
  powershell -NoProfile -ExecutionPolicy Bypass -Command "& '%PS1%' -InputFolder '%INPUT_FOLDER%' -OutputFolder '%OUTPUT_FOLDER%' -ApiBaseUrl '%API_BASE_URL%' -ApiKey '%API_KEY%' -Timeout %TIMEOUT% -Parallelism !PARALLELISM! %COPY_ARG% %ORGANIZE_SOURCE_ARG% %ORIGINAL_FOLDER_ARG% %MAPPING_CSV_ARG% -LogFilePath '%LOG_FILE_PATH%' -Verbose"
) else (
  powershell -NoProfile -ExecutionPolicy Bypass -Command "& '%PS1%' -InputFolder '%INPUT_FOLDER%' -OutputFolder '%OUTPUT_FOLDER%' -ApiBaseUrl '%API_BASE_URL%' -ApiKey '%API_KEY%' -Timeout %TIMEOUT% -Parallelism !PARALLELISM! %COPY_ARG% %ORGANIZE_SOURCE_ARG% %ORIGINAL_FOLDER_ARG% %MAPPING_CSV_ARG% -Verbose"
)

set "EXITCODE=%ERRORLEVEL%"

if "%EXITCODE%"=="0" (
  echo [OK] 正常終了
) else (
  echo [ERROR] 処理失敗 ExitCode=%EXITCODE%
)

exit /b %EXITCODE%

:usage
echo.
echo Usage:
echo   %~nx0 "InputFolder" "OutputFolder" "ApiBaseUrl" "ApiKey" [Timeout] [CopyOriginal] [LogFilePath] [Parallelism] [OrganizeSourceFilesAfterCopy] [OriginalFolder] [MappingCsvPath]
echo.
echo Parallelism:
echo   - 指定あり : その値を使用
echo   - 指定なし : CPUコア数から自動決定（最低2、最大6）
echo.
echo OrganizeSourceFilesAfterCopy:
echo   - default is true
echo   - true / false ^(1 / 0 / yes / no^)
echo.
echo Example:
echo   %~nx0 "C:\work\input" "C:\work\output" "https://api.example.com" "api-key"
echo   %~nx0 "C:\work\input" "C:\work\output" "https://api.example.com" "api-key" 600 false "C:\work\logs\rename.log" 5 true "C:\work\Original" "C:\work\logs\mapping.csv"
echo.
exit /b 2