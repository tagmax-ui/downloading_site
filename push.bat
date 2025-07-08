@echo off
REM git.bat: Automatisation du commit et du push Git

REM Se placer dans le répertoire du script
cd /d "%~dp0"

REM Mettre en scène toutes les modifications
git add -A

REM Générer le message de commit basé sur la date et l'heure
for /f "tokens=1-3 delims=/ " %%a in ('date /t') do (
    set today=%%c%%b%%a
)
for /f "tokens=1-2 delims=: " %%a in ('time /t') do (
    set hour=%%a
    set minute=%%b
)

REM Supprimer les espaces éventuels dans l’heure (ex:  9:05 → 09:05)
if "%hour:~1,1%"==":" set hour=0%hour:~0,1%

set commit_msg=Site_Telechargement %today% %hour%:%minute%

REM Afficher le message généré
echo Commit: %commit_msg%

REM Commit
git commit -m "%commit_msg%"

REM Pousser vers origin main
git push origin main

REM Exit sans pause
exit /b 0
