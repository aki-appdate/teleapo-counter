@echo off
echo === 架電カウンター ビルド ===

pip install pyinstaller gspread google-auth

pyinstaller ^
  --onefile ^
  --windowed ^
  --name "架電カウンター" ^
  --add-data "config.json;." ^
  main.py

echo.
echo ビルド完了！ dist\架電カウンター.exe を配布してください。
echo 同じフォルダに credentials.json と config.json を置いてください。
pause
