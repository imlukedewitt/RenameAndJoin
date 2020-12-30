@echo off
PUSHD %~dp0
powershell -executionpolicy bypass -nologo -file .\RenameAndJoin.ps1
POPD