@echo off
echo Запуск Flask приложения...
start cmd /k "python app.py"
timeout /t 3
echo Запуск LocalTunnel...
start cmd /k "lt --port 5000"
echo Готово! Скопируйте ссылку из второго окна
pause