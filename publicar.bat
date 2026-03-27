@echo off
chcp 65001 >nul
echo.
echo ╔══════════════════════════════════════╗
echo ║   EQS BI - Publicar para Equipe     ║
echo ╚══════════════════════════════════════╝
echo.

cd /d "%~dp0"

echo [1/3] Adicionando data.json...
git add data.json

echo [2/3] Criando commit...
git commit -m "📊 Dados atualizados - %date% %time:~0,5%"

echo [3/3] Enviando para GitHub...
git push

echo.
echo ✅ Publicado com sucesso! A equipe já pode ver os dados atualizados.
echo.
pause
