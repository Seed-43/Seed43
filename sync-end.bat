@echo off
cd /d "C:\Users\Fred\AppData\Roaming\pyRevit\Extensions\Seed43.extension"
echo Pushing changes to dev...
git add .
git commit -m "WIP %date% %time%"
git push origin dev
pause