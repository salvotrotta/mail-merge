@echo off
echo ================================================
echo   MAIL MERGE - Compilazione EXE (onedir)
echo ================================================
echo.

:: Verifica Python
where python >nul 2>&1
if errorlevel 1 (
    echo ERRORE: Python non trovato nel PATH.
    pause & exit /b 1
)

:: Installa dipendenze
echo Installo dipendenze...
python -m pip install pyinstaller python-docx openpyxl pillow pikepdf -q

:: Verifica PyInstaller
python -m PyInstaller --version >nul 2>&1
if errorlevel 1 (
    echo ERRORE: PyInstaller non installato correttamente.
    pause & exit /b 1
)

:: Genera icona se mancante
if not exist "icon.ico" (
    echo Genero icona...
    python crea_icona.py
)

:: Rimuovi build precedenti
if exist "dist\MailMerge" rmdir /s /q "dist\MailMerge"
if exist "build" rmdir /s /q "build"

echo.
echo Compilazione in corso...
echo.

python -m PyInstaller ^
    --onedir ^
    --windowed ^
    --name "MailMerge" ^
    --icon "icon.ico" ^
    --hidden-import "openpyxl" ^
    --hidden-import "openpyxl.styles" ^
    --hidden-import "openpyxl.utils" ^
    --hidden-import "openpyxl.workbook" ^
    --hidden-import "docx" ^
    --hidden-import "docx.oxml" ^
    --hidden-import "docx.oxml.ns" ^
    --hidden-import "docx.shared" ^
    --hidden-import "docx.enum.text" ^
    --hidden-import "lxml" ^
    --hidden-import "lxml.etree" ^
    --hidden-import "lxml._elementpath" ^
    --hidden-import "tkinter" ^
    --hidden-import "tkinter.ttk" ^
    --hidden-import "tkinter.filedialog" ^
    --hidden-import "tkinter.messagebox" ^
    --hidden-import "pikepdf" ^
    --hidden-import "pikepdf._core" ^
    --collect-all "docx" ^
    --collect-all "openpyxl" ^
    --collect-all "pikepdf" ^
    --clean ^
    mail_merge_gui.py

if errorlevel 1 (
    echo.
    echo ERRORE durante la compilazione.
    pause & exit /b 1
)

echo.
echo ================================================
echo   Cartella creata in: dist\MailMerge\
echo   File principale:    dist\MailMerge\MailMerge.exe
echo   Ora apri setup.iss con Inno Setup (F9)
echo   per creare l'installer.
echo ================================================
pause