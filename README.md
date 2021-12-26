# windows-scripts
A collection of Windows scripts I've found userful

## Installation instructions
These installation instructions are for users that have not used command line before.
If you are familiar with Git and PowerShell, you can clone and use the repository as you like instead of following these instructions.

First check whether the scripts have already been installed on your computer.
This can be done by checking whether there is a folder `Git\windows-scripts` in your home directory.
The full path to this folder should be something like `C:\Users\<username>\Git\windows-scripts`.
If this folder exists, move on to the usage instructions.
If it does not exist, please continue these installation instructions.

To ease the setting up of these scripts on your computer, I have created an installation script.
To download it, right-click
[this link](https://raw.githubusercontent.com/AgenttiX/windows-scripts/master/repo-installer.bat)
and select "Save link as..." to save the file to a directory of your choice.
Then right-click the downloaded file and select "Run as administrator".
This should open a command-line window and setup the scripts and their dependencies for you.
Once the setup is complete, you can find the scripts in the directory `Git\windows-scripts` within your user folder (usually `C:\Users\<username>`).

## Usage instructions
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

Now in the command prompt you should write ".\" and the name of the script I have asked you to run.
For example, the reporting script is selected with `.\report.ps1`.
Then press enter to run the script.
Many of these scripts have to be run as an administrator, and therefore they will request those privileges
and then another window that operates with those privileges will open.
If you get any errors, please send me a screenshot.

Once the script has run, it may provide you with additional instructions such as how to send the results.
Please follow those.
