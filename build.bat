@echo off
cd /d "%~dp0"
taskkill /f /im autokey.exe >nul 2>nul
call "C:\Program Files\Microsoft Visual Studio\2022\Community\VC\Auxiliary\Build\vcvars64.bat" >nul 2>nul
rc /nologo resource.rc
cl /nologo /W4 /O2 /utf-8 main.c macro.c resource.res /Fe:autokey.exe /I. /link user32.lib gdi32.lib shell32.lib setupapi.lib winmm.lib comctl32.lib comdlg32.lib shlwapi.lib /SUBSYSTEM:WINDOWS /ENTRY:wWinMainCRTStartup
