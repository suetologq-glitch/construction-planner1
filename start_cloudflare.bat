@echo off
echo ========================================
echo   Запуск строительного планировщика
echo ========================================
echo.

echo 1. Запуск Flask приложения...
start cmd /k "cd /d C:\construction_planner && python app.py"

echo.
echo 2. Ожидание запуска сервера...
timeout /t 5 /nobreak > nul

echo.
echo 3. Запуск Cloudflare Tunnel...
start cmd /k "cd /d C:\construction_planner && cloudflared tunnel --url http://localhost:5000"

echo.
echo ========================================
echo   ГОТОВО!
echo ========================================
echo.
echo Скопируйте ссылку из второго окна (вида: https://xxxx.trycloudflare.com)
echo.
pause