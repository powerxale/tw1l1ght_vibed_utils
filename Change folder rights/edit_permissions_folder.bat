@echo off
title Granting complete control to files and folders...

REM Checks whether a path was provided as an argument
IF "%~1"=="" (
    echo.
    echo ERROR: No folder path specified.
    echo Please drag the desired folder to the "start_with_admin.vbs" icon.
    echo Or, start the script from the command line by providing the path, for example:
    echo "%~nx0" "C:\Path\Directory."
    echo.
    pause
    exit /b 1
)

SET "TARGET_FOLDER=%~1"

echo.
echo Attempt to grant complete control to the folder: "%TARGET_FOLDER%" and all its contents (subfolders and files)...
echo This operation may take some time, depending on the size of the folder.
echo.

REM Grants complete control to the 'Users' group recursively
REM /T: It goes through all subdirectories and files.
REM /C: Continues even in case of errors (useful for locked or inaccessible files).
REM /Q: Suppresses success messages (makes output cleaner).
icacls "%TARGET_FOLDER%" /grant Users:F /T /C /Q

IF %ERRORLEVEL% NEQ 0 (
    echo.
    echo AN ERROR occurred while editing permissions.
    echo Make sure you have administrator rights and that the specified path is correct and exists.
) ELSE (
    echo.
    echo Permissions successfully modified for: "%TARGET_FOLDER%"
)

echo.
echo Operation completed.
pause