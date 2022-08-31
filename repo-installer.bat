@echo off
echo This script is for setting up Mika's Windows scripts.

echo Allowing PowerShell scripts
powershell -command "& {Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Force}"
WHERE choco
IF %ERRORLEVEL% EQU 0 (
    echo Chocolatey seems to be already installed.
) ELSE (
    echo Installing Chocolatey
    powershell -command "Set-ExecutionPolicy Bypass -Scope Process -Force; [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor 3072; iex ((New-Object System.Net.WebClient).DownloadString('https://community.chocolatey.org/install.ps1'))"
    REM Load Chocolatey command
    CALL %PROGRAMDATA%\chocolatey\bin\RefreshEnv.cmd
)
WHERE git
IF %ERRORLEVEL% EQU 0 (
    echo Git seems to be already installed.
) ELSE (
    echo Installing Git
    choco upgrade git -y
    REM Load Git command
    CALL %PROGRAMDATA%\chocolatey\bin\RefreshEnv.cmd
)

IF exist %USERPROFILE%\Git (
    echo The Git folder seems to already exist.
) ELSE (
    echo Creating a folder for Git repositories.
    mkdir %USERPROFILE%\Git
)
IF exist %USERPROFILE%\Git\windows-scripts (
    echo The repository seems to already exist.
    echo Performing "git pull" to update the scripts.
    cd %USERPROFILE%\Git\windows-scripts
    git pull
) ELSE (
    echo Cloning the repository
    cd %USERPROFILE%\Git
    git clone "https://github.com/AgenttiX/windows-scripts"
)

echo Configuring the repository directory to be safe.
git config --global --add safe.directory %USERPROFILE%\Git\windows-scripts

echo Opening the scripts folder in File Explorer.
%SYSTEMROOT%\explorer.exe %USERPROFILE%\Git\windows-scripts
echo The setup is ready. You can close this window now.
pause
