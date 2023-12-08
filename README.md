# windows-scripts
A collection of Windows scripts I've found useful

## Installation instructions
These installation instructions are for users that have not used command line before.
If you are familiar with Git and PowerShell, you can clone and use the repository as you like instead of following these instructions.

First check whether the scripts have already been installed on your computer.
This can be done by checking whether there is a folder `Git\windows-scripts` in your home directory,
or in the root of your C drive.
The full path to this folder should be something like
`C:\Users\<username>\Git\windows-scripts` or `C:\Git\windows-scripts`.
If this folder exists, move on to the [usage instructions](#usage-instructions).
If it does not exist, please continue these installation instructions.

To ease the setting up of these scripts on your computer, I have created an installation script.
To download it, right-click
[this link](https://raw.githubusercontent.com/AgenttiX/windows-scripts/master/Install-Repo.bat)
and select "Save link as..." to save the file to a directory of your choice.
Then right-click the downloaded file and select "Run as administrator".
This should open a command-line window and setup the scripts and their dependencies for you.

If you get a Windows SmartScreen error saying that the file is blocked,
you have to unblock it by right-clicking the downloaded file and selecting Properties,
and then checking the checkbox named Unblock at the bottom of the window.
Then click OK and run the script again.

Once the setup is complete, you can find the scripts in the directory
`Git\windows-scripts` within your user folder (usually `C:\Users\<username>`) if you chose per-user installation,
or at `C:\Git\windows-scripts` if you chose global installation.

## Usage instructions
Following the installation instructions should have created desktop icons for the installer and maintenance scripts.
You can run these simply by double-clicking them.

To run the other scripts, please follow these instructions.
Most of these scripts are
[PowerShell](https://en.wikipedia.org/wiki/PowerShell)
scripts, which cannot be run by simply clicking them for security reasons.
To run them, first open the `Git\windows-scripts` folder you have downloaded according to the installation instructions above in File Explorer ("resurssienhallinta" or "oma tietokone" in Finnish).
Then press and hold shift, and right-click some empty space in the folder.
In the context menu that opens you should see the option for "Open in Windows Terminal" (Windows 11) or "Open in PowerShell" (Windows 10).
Select it.
This should open a command prompt (a text-based window with a blinking cursor).
If it's been a while since the scripts were downloaded,
or if I've told you that there's a new version available,
write `git pull` and press enter.
This will download the latest updates to the scripts.

Now in the command prompt you should write `.\ ` and the name of the script I have asked you to run.
For example, the reporting script is selected with `.\Report.ps1`, the software installation script with `.\Install-Software.ps1` and the maintenance script with `.\Maintenance.ps1`.
Then press enter to run the script.
Many of these scripts have to be run as an administrator, and therefore they will request those privileges
and then another window that operates with those privileges will open.
If you get any errors, please send me a screenshot.

Once the script has run, it may provide you with additional instructions such as how to send the results.
Please follow those.

## Uninstalling managed software
The installation script uses various package managers such as Chocolatey and Winget to manage the installed software.
Therefore if you uninstall a managed program with its own uninstaller or from the list of installed applications in Windows settings, the program will be reinstalled with the next update.
Therefore the program has to be uninstalled using the package manager.
Please open an elevated PowerShell window as instructed above.
For programs installed with Chocolatey the command is `choco uninstall <name of software>` (without the brackets).
For programs installed with Winget the command is `winget uninstall --id <ID of program>`.
Then press enter.
You can find the names and IDs of the programs from the user interface of the installation script.

## Recreating virtualenvs after a Python upgrade
If Python is installed through Chocolatey or some other package manager, it will be updated automatically.
Major version upgrades will break existing virtualenvs.
To recreate the virtualenv of a Python project, first delete the `venv` folder from the project directory.
Then in PyCharm go to "File -> Settings -> Project -> Python Interpreter" and click on the gear symbol at the top-right.
There select "Show all", select the old virtualenv and click the minus sign.
Now the old virtualenv shold be deleted.
Then click the plus sign and select "Virtualenv Environment -> New environment".
The path should be the same as for the `venv` directory you just deleted.
Then click OK on all the settings windows to close them.
Now we can reinstall the project dependencies.
Select "Terminal" from the bottom of the PyCharm window.
This should open a terminal.
However, it may not yet have the new virtualenv activated.
Therefore write `exit` in the terminal to close it.
Now open it again.
Then write `pip -V` and press enter.
This should show a path to the `venv` we just created.
Now if the project has a `requirements.txt` file we can write `pip install -r requirements.txt` to install the dependencies.
Otherwise we can install the dependencies manually, for example with `pip install matplotlib numpy`.
The project should now be ready to use again.
