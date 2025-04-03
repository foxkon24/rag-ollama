@echo off

@if not "%~0"=="%~dp0.\%~nx0" start /min cmd /c,"%~dp0.\%~nx0" %* & goto :eof

cd C:\laragon\bin\ngrok

cd %~dp0

ngrok http 5010

pause

