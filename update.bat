@echo off
cd /d "C:\Users\joaoa\OneDrive\Desktop\Dashboard Logistica"
echo Updating dashboard...
git add .
git commit -m "Dashboard update"
git push origin master
echo.
echo Done! Check sacoorlogistics.netlify.app in a few seconds.
pause
