@echo off
setlocal enabledelayedexpansion

rem Prompt the user for their username
set "username=%USERNAME%"

rem Set the target path and name for the shortcut
set "targetPath=C:\Users\%username%\OneDrive - Dell Technologies\CM Tool\CM Tool Files\cmtoolapp.exe"
set "shortcutName=CM Tool"

rem Get the desktop folder path
for /f "tokens=2*" %%a in ('reg query "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders" /v Desktop ^| find "Desktop"') do set "desktopPath=%%b"

rem Set the shortcut path
set "shortcutPath=!desktopPath!\%shortcutName%.lnk"

rem Create the shortcut
echo Set WshShell = WScript.CreateObject("WScript.Shell") > CreateShortcut.vbs
echo Set shortcut = WshShell.CreateShortcut("!shortcutPath!") >> CreateShortcut.vbs
echo shortcut.TargetPath = "!targetPath!" >> CreateShortcut.vbs
echo shortcut.Save >> CreateShortcut.vbs
cscript //nologo CreateShortcut.vbs
del CreateShortcut.vbs

echo Shortcut created: %shortcutPath%

cls

rem Add message that auto-closes in 5 seconds
echo.
echo The CMTool was installed and there is a shortcut in your desktop.
echo This window will close in 5 seconds.
timeout /t 5 /nobreak >nul