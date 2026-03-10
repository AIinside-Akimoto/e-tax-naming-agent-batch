@echo off
setlocal EnableExtensions

rem ============================================
rem ファイル命名アシスタント（Leapnet API呼び出し用）
rem ============================================

rem ===== 文字コード設定（必要なら有効化）=====
rem chcp 65001 >nul

rem ===== スクリプト配置ディレクトリ =====
set "SCRIPT_DIR=%~dp0"
set "PS1=%SCRIPT_DIR%Invoke-NamingAgentBatch.ps1"

REM =========================================================
REM ファイル命名アシスタント 起動用バッチ
REM
REM 使い方:
REM   ファイル命名アシスタント.bat
REM      "入力フォルダー"
REM      "出力フォルダー"
REM      "APIベースURL"
REM      "APIキー"
REM      [タイムアウト秒]
REM      [未リネームファイルコピー(true/false)]
REM      [ログファイルパス]
REM
REM 実行例:
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
REM =========================================================

rem ===== 引数定義 =====
rem 1: 入力フォルダー
rem 2: 出力フォルダー
rem 3: APIベースURL
rem 4: APIキー
rem 5: タイムアウト秒
rem 6: 未リネームファイルコピー(true/false)
rem 7: ログファイルパス（任意）

set "INPUT_FOLDER=%~1"
set "OUTPUT_FOLDER=%~2"
set "API_BASE_URL=%~3"
set "API_KEY=%~4"
set "TIMEOUT=%~5"
set "COPY_ORIGINAL=%~6"
set "LOG_FILE_PATH=%~7"

rem ===== デフォルト値 =====
if "%TIMEOUT%"=="" set "TIMEOUT=600"
if "%COPY_ORIGINAL%"=="" set "COPY_ORIGINAL=false"

rem ===== 必須引数チェック =====
if "%INPUT_FOLDER%"=="" goto :usage
if "%OUTPUT_FOLDER%"=="" goto :usage
if "%API_BASE_URL%"=="" goto :usage
if "%API_KEY%"=="" goto :usage

rem ===== PowerShellスクリプト存在確認 =====
if not exist "%PS1%" (
  echo [ERROR] PowerShellスクリプトが見つかりません: "%PS1%"
  exit /b 1
)

rem ===== 入力フォルダー存在確認 =====
if not exist "%INPUT_FOLDER%" (
  echo [ERROR] 入力フォルダーが見つかりません: "%INPUT_FOLDER%"
  exit /b 1
)

rem ===== 出力フォルダー作成（存在しない場合）=====
if not exist "%OUTPUT_FOLDER%" (
  mkdir "%OUTPUT_FOLDER%" 2>nul
  if errorlevel 1 (
    echo [ERROR] 出力フォルダーの作成に失敗しました: "%OUTPUT_FOLDER%"
    exit /b 1
  )
)

rem ===== CopyOriginal フラグ作成 =====
set "COPY_ARG="
if /I "%COPY_ORIGINAL%"=="true" set "COPY_ARG=-CopyOriginal"
if /I "%COPY_ORIGINAL%"=="1"    set "COPY_ARG=-CopyOriginal"
if /I "%COPY_ORIGINAL%"=="yes"  set "COPY_ARG=-CopyOriginal"

echo [START]
echo [INFO] PowerShellScript = "%PS1%"
echo [INFO] InputFolder      = "%INPUT_FOLDER%"
echo [INFO] OutputFolder     = "%OUTPUT_FOLDER%"
echo [INFO] ApiBaseUrl       = "%API_BASE_URL%"
echo [INFO] TimeoutSeconds   = %TIMEOUT%
echo [INFO] CopyOriginal     = %COPY_ORIGINAL%
if not "%LOG_FILE_PATH%"=="" echo [INFO] LogFilePath      = "%LOG_FILE_PATH%"

rem ===== PowerShellスクリプト実行 =====
if not "%LOG_FILE_PATH%"=="" (
  powershell -NoProfile -ExecutionPolicy Bypass -File "%PS1%" ^
    -InputFolder "%INPUT_FOLDER%" ^
    -OutputFolder "%OUTPUT_FOLDER%" ^
    -ApiBaseUrl "%API_BASE_URL%" ^
    -ApiKey "%API_KEY%" ^
    -Timeout %TIMEOUT% ^
    %COPY_ARG% ^
    -LogFilePath "%LOG_FILE_PATH%" ^
    -Verbose
) else (
  powershell -NoProfile -ExecutionPolicy Bypass -File "%PS1%" ^
    -InputFolder "%INPUT_FOLDER%" ^
    -OutputFolder "%OUTPUT_FOLDER%" ^
    -ApiBaseUrl "%API_BASE_URL%" ^
    -ApiKey "%API_KEY%" ^
    -Timeout %TIMEOUT% ^
    %COPY_ARG% ^
    -Verbose
)

set "EXITCODE=%ERRORLEVEL%"

if "%EXITCODE%"=="0" (
  echo [OK] 正常に処理が完了しました。
) else (
  echo [ERROR] 処理に失敗しました。ExitCode=%EXITCODE%
)

exit /b %EXITCODE%

:usage
echo.
echo 使い方:
echo   %~nx0 "InputFolder" "OutputFolder" "ApiBaseUrl" "ApiKey" [TimeoutSeconds] [CopyOriginal(true/false)] [LogFilePath]
echo.
echo 実行例:
echo   %~nx0 "C:\work\naming-agent\Input" "C:\work\naming-agent\Output" "https://stg-agent.leapnet.com/xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" "your-api-key" 600 false "C:\work\logs\rename.log"
echo.
exit /b 2