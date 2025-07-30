@echo off
REM Batch file to run the PowerShell script for converting PowerPoint files to PDF.

REM Get the directory of this batch file
set "batchPath=%~dp0"

REM Set the name of the PowerShell script
set "psScript=Convert-PPTXtoPDF.ps1"

REM Check if the PowerShell script exists
if not exist "%batchPath%%psScript%" (
    echo ERROR: PowerShell script '%psScript%' not found in the same directory.
    pause
    exit /b
)

:start
echo.
echo =================================================================
echo   PowerPoint to PDF Batch Converter
echo =================================================================
echo.
echo   This script will convert all .pptx and .ppt files in a
echo   specified folder to PDF format.
echo.
echo   The converted PDF files will be saved in a new 'PDF'
echo   subfolder inside the source directory.
echo.
echo =================================================================
echo.

:getFolder
set "sourceFolder="
set /p "sourceFolder=Please enter the full path to the folder containing your PowerPoint files: "

if not defined sourceFolder (
    echo.
    echo You did not enter a path. Please try again.
    goto getFolder
)

if not exist "%sourceFolder%" (
    echo.
    echo The path you entered does not exist. Please try again.
    goto getFolder
)

echo.
echo Starting conversion process...
echo.

REM Run the PowerShell script
powershell.exe -ExecutionPolicy Bypass -File "%batchPath%%psScript%" -SourceFolder "%sourceFolder%"

echo.
echo =================================================================
echo   Conversion process finished.
echo =================================================================
echo.

:askToContinue
set /p "continueChoice=Do you want to convert more files? (Y/N): "
if /i "%continueChoice%"=="Y" goto start
if /i "%continueChoice%"=="N" goto end
echo Invalid choice. Please enter Y or N.
goto askToContinue

:end
echo.
echo Exiting the converter.
pause
exit /b
