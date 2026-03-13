@echo off
echo ====================================
echo  MD to DOCX Converter - Build
echo  by Jair Lima
echo ====================================
echo.

cd /d "%~dp0"

pyinstaller ^
  --onefile ^
  --name md2docx ^
  --console ^
  --clean ^
  --strip ^
  --hidden-import=mistune ^
  --hidden-import=mistune.plugins ^
  --hidden-import=mistune.plugins.table ^
  --hidden-import=mistune.plugins.strikethrough ^
  --hidden-import=mistune.plugins.footnotes ^
  --hidden-import=mistune.plugins.task_lists ^
  --hidden-import=docx ^
  --hidden-import=docx.oxml ^
  --hidden-import=docx.opc ^
  --hidden-import=lxml ^
  --hidden-import=lxml.etree ^
  md2docx.py

if %ERRORLEVEL% EQU 0 (
    echo.
    echo ====================================
    echo  Build concluido com sucesso!
    echo  Executavel: dist\md2docx.exe
    echo ====================================
    echo.
    echo Para instalar globalmente, copie dist\md2docx.exe
    echo para uma pasta que esteja no seu PATH.
    echo Exemplo: C:\Windows\System32  ou  C:\Tools
) else (
    echo.
    echo [ERRO] Build falhou. Verifique os logs acima.
)

pause
