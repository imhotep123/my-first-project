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
    echo 出力先: dist\kiyaku_viewer\kiyaku_viewer.exe
    explorer dist\kiyaku_viewer
) else (
    echo.
    echo ビルド失敗。エラーを確認してください。
    pause
)
