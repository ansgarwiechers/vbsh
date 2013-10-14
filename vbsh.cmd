@echo off

title %~n0
chcp 1250 >nul

"%SystemRoot%\system32\cscript.exe" //NoLogo "%~dpn0.vbs"
