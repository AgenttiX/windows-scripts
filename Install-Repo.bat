@echo off
echo This script is for setting up Mika's Windows scripts.

NET SESSION >nul 2>&1
IF %ERRORLEVEL% NEQ 0 (
	echo This script should be run as administrator. Please right-click this script and select "Run as administrator".
	exit 1
)

SET /P GLOBALYESNO=Do you want to install globally for all users? [y/n] (Write "y" and press enter, unless you are developing the scripts yourself.)
IF /I "%GLOBALYESNO%" EQU "y" (
    SET /A GLOBALINSTALL=1
    SET BASEPATH=%SYSTEMDRIVE%
)ELSE (
	IF /I "%GLOBALYESNO%" EQU "n" (
        SET /A GLOBALINSTALL=0
		SET BASEPATH=%USERPROFILE%
	) ELSE (
		echo Invalid selection: %GLOBALYESNO%
		exit /b 1
	)
)

echo Allowing PowerShell scripts
powershell -Command "& {Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Force}"

echo Current path of choco:
WHERE choco
IF %ERRORLEVEL% EQU 0 (
    echo Chocolatey seems to be already installed.
) ELSE (
    echo Installing Chocolatey
    powershell -command "Set-ExecutionPolicy Bypass -Scope Process -Force; [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor 3072; iex ((New-Object System.Net.WebClient).DownloadString('https://community.chocolatey.org/install.ps1'))"
    REM Load Chocolatey command
    CALL %PROGRAMDATA%\chocolatey\bin\RefreshEnv.cmd
)
echo Current path for git:
WHERE git
IF %ERRORLEVEL% EQU 0 (
    echo Git seems to be already installed.
) ELSE (
    echo Installing Git
    choco upgrade git -y
    REM Load Git command
    CALL %PROGRAMDATA%\chocolatey\bin\RefreshEnv.cmd
)

IF exist %BASEPATH%\Git (
    echo The Git folder seems to already exist.
) ELSE (
    echo Creating a folder for Git repositories.
    mkdir %BASEPATH%\Git
)
IF exist %BASEPATH%\Git\windows-scripts (
    echo The repository seems to already exist.
    echo Performing "git pull" to update the scripts.
    cd %BASEPATH%\Git\windows-scripts
    git pull
) ELSE (
    echo Cloning the repository
    cd %BASEPATH%\Git
    git clone "https://github.com/AgenttiX/windows-scripts"
)

IF "%GLOBALINSTALL%"=="0" (
	echo Marking the repository directory to be safe for Git.
	git config --global --add safe.directory %USERPROFILE%\Git\windows-scripts
)

echo Creating shortcuts and scheduled task
powershell -File %BASEPATH%\Git\windows-scripts\Maintenance.ps1 -SetupOnly

echo Opening the scripts folder in File Explorer.
%SYSTEMROOT%\explorer.exe %BASEPATH%\Git\windows-scripts
echo The setup is ready. You can close this window now.
pause
