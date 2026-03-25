@echo off
echo Запуск Flask приложения...
start cmd /k "cd C:\construction_planner && python app.py"
timeout /t 3
echo Запуск Pinggy (без пароля)...
start cmd /k "ssh -R 80:localhost:5000 pinggy.io"
echo Готово! Скопируйте ссылку из второго окна
pause