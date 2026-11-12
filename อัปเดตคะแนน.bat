@echo off
cd /d "D:\DATA Drive D\ควบคุมคุณภาพ\Action Plan\2026\Engagement"

echo [1/3] Reading Excel and rebuilding website...
python build_scores.py
if %ERRORLEVEL% == 1 goto error

echo [2/3] Pushing to GitHub...
git add -f scores.html
git commit -m "update scores"
git push origin main
if %ERRORLEVEL% == 1 goto error

echo [3/3] Done! Opening website in 5 seconds...
timeout /t 5 /nobreak >nul
start "" "https://guileless-bombolone-267e3d.netlify.app/scores.html"
echo Press any key to close...
pause >nul
goto end

:error
echo ERROR - something went wrong
pause

:end
