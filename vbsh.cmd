@echo off

chcp 1250 >nul

"%SystemRoot%\system32\cscript.exe" //NoLogo "%~dp0vbsh.vbs"
