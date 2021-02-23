@echo off
REM This is a script for fixing Android fastboot on Ryzen-based systems
REM to get rid of the "Press any key to shutdown" error.
REM Found from:
REM https://forum.xda-developers.com/t/help-press-any-key-to-shutdown-in-fastboot.3816021/
REM Original file:
REM https://drive.google.com/file/d/1NqPotx06yRuhPsOdEAdS4JNsk7qWWchF/view

reg add "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\usbflags\18D1D00D0100" /v "osvc" /t REG_BINARY /d "0000" /f
reg add "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\usbflags\18D1D00D0100" /v "SkipContainerIdQuery" /t REG_BINARY /d "01000000" /f
reg add "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\usbflags\18D1D00D0100" /v "SkipBOSDescriptorQuery" /t REG_BINARY /d "01000000" /f

pause
