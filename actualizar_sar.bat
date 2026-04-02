@echo off
echo ==============================
echo ACTUALIZANDO DASHBOARD SAR
echo ==============================

cd /d "C:\Users\14538348-k\Desktop\DASHBOARD (html)\Antiguos\SAR\sar\Gestion Clinica\sar"

echo.
echo [1] Generando JSON...
python generar_json_sar_directivo.py

echo.
echo [2] Agregando cambios...
git add .

echo.
echo [3] Verificando cambios...
git diff --staged --quiet
IF %ERRORLEVEL%==0 (
    echo No hay cambios para subir.
) ELSE (
    echo [4] Haciendo commit...
    git commit -m "update automatico SAR"

    echo [5] Subiendo a GitHub...
    git push
)

echo.
echo ==============================
echo PROCESO FINALIZADO
echo ==============================
pause