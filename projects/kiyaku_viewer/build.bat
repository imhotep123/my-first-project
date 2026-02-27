@echo off
cd /d %~dp0

echo ===================================
echo  kiyaku_viewer PyInstaller ビルド
echo ===================================
echo.

pyinstaller kiyaku_viewer.spec --clean --noconfirm

if %errorlevel% == 0 (
    echo.
    echo ビルド成功！
    echo.
    echo ★ 実行ファイルの場所:
    echo    %~dp0dist\kiyaku_viewer\kiyaku_viewer.exe
    echo.
    echo ※ build\ フォルダ内の exe は中間ファイルです。
    echo    必ず dist\kiyaku_viewer\ の exe を使用してください。
    echo.
    explorer "%~dp0dist\kiyaku_viewer"
) else (
    echo.
    echo ビルド失敗。エラーを確認してください。
    pause
)
