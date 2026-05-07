@echo off
cd /d "%~dp0"

echo.
echo ========================================
echo   AGENTE DIEX — Deploy
echo ========================================
echo.

set /p msg="Descricao da alteracao: "

git add .
git commit -m "%msg%"
git push

echo.
echo ========================================
echo   Deploy concluido! App atualizando...
echo ========================================
echo.
pause
