@echo off
wmic process where name="pythonw.exe" get CommandLine 2>NUL | find /I /N "automerge.pyw" >NUL
if "%ERRORLEVEL%" NEQ "0" start pythonw automerge.pyw
