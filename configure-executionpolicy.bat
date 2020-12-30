rem This script configures PowerShell to allow the execution of local custom scripts
powershell -command "& {Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Force}"
