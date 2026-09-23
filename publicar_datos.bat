@echo off
REM Copia data\latest.json (generado por actualizar_ruta.bat) al repo
REM PRIVADO mapa-rutas-datos y lo sube. Este es el unico paso que sube
REM datos reales de clientes a algun lado, por eso es privado.
setlocal

set ORIGEN=%~dp0data\latest.json
set DESTINO=C:\MAPA-RUTAS-DATOS

if not exist "%ORIGEN%" (
  echo No existe %ORIGEN%
  echo Corre primero actualizar_ruta.bat con la carpeta del dia.
  pause
  exit /b 1
)

copy /y "%ORIGEN%" "%DESTINO%\latest.json"

cd /d "%DESTINO%"
git add latest.json
git commit -m "Ruta del dia"

REM Cuando arrastras un punto para corregirlo (modo admin), eso se sube
REM directo a GitHub por internet, sin pasar por esta carpeta. Por eso
REM antes de subir lo de hoy hay que traer primero esos cambios (si los
REM hay) para no toparse con un "push rechazado" que se ignore en
REM silencio y deje sin publicar la ruta del dia.
git pull --rebase origin main
if errorlevel 1 (
  echo.
  echo ========================================================
  echo ERROR: no se pudo combinar con los cambios de GitHub.
  echo NO se subio la ruta de hoy. Avisa antes de intentar de nuevo.
  echo ========================================================
  pause
  exit /b 1
)

git push
if errorlevel 1 (
  echo.
  echo ========================================================
  echo ERROR: no se pudo subir el cambio a GitHub. Revisa tu
  echo conexion a internet y vuelve a correr este mismo archivo.
  echo ========================================================
  pause
  exit /b 1
)

echo.
echo Listo. Los choferes ya ven la ruta actualizada en el mapa.
echo.
pause
