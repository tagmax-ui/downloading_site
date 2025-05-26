@echo off
REM === Deploy TAGmax update to Git and Render ===

REM Step 1: Stage all changes
git add .

REM Step 2: Build commit message with timestamp
for /f "tokens=1-4 delims=/ " %%a in ('date /t') do (
    set day=%%a
    set month=%%b
    set year=%%c
)
for /f "tokens=1-2 delims=: " %%a in ('time /t') do (
    set hour=%%a
    set min=%%b
)
REM Handle AM/PM formats (Windows date/time commands may vary)
set msg=Deploy %year%-%month%-%day%_%hour%-%min%
REM Remove trailing spaces if any
set msg=%msg: =%

REM Step 3: Commit with auto message
git commit -m "%msg%"

REM Step 4: Push to remote
git push

REM Step 5: Reminder for Render
echo.
echo === NOW GO TO RENDER DASHBOARD ===
echo 1. Select "Téléchargement de TAGmax"
echo 2. Click "Deploy from latest deploy" (should be: %msg%)
echo.
pause
