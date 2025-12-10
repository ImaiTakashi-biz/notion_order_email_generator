@echo off
chcp 65001 >nul
echo ========================================
echo 注文書メール自動作成App ビルドスクリプト
echo ========================================
echo.

REM .envファイルの存在確認
if not exist ".env" (
    echo [警告] .envファイルが見つかりません。
    echo ビルド前に.envファイルを作成してください。
    echo.
    pause
    exit /b 1
)

REM email_accounts.jsonの存在確認
if not exist "email_accounts.json" (
    echo [警告] email_accounts.jsonファイルが見つかりません。
    echo ビルド前にemail_accounts.jsonファイルを作成してください。
    echo.
    pause
    exit /b 1
)

REM app_icon.icoの存在確認
if not exist "app_icon.ico" (
    echo [警告] app_icon.icoファイルが見つかりません。
    echo ビルド前にapp_icon.icoファイルを用意してください。
    echo.
    pause
    exit /b 1
)

echo [情報] 必要なファイルの確認が完了しました。
echo.
echo [情報] PyInstallerでビルドを開始します...
echo.

REM PyInstallerでビルド実行
pyinstaller "注文書メール自動作成App.spec"

if %ERRORLEVEL% EQU 0 (
    echo.
    echo ========================================
    echo ビルドが完了しました！
    echo ========================================
    echo.
    echo 実行ファイルの場所: dist\注文書メール自動作成App.exe
    echo.
    echo 配布時は、この.exeファイルのみを配布してください。
    echo .envとemail_accounts.jsonは実行ファイルに含まれています。
    echo.
) else (
    echo.
    echo [エラー] ビルド中にエラーが発生しました。
    echo.
    pause
    exit /b 1
)

pause

