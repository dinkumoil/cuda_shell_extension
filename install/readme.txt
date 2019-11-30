These scripts are for installing/uninstalling the compiled context menu shell
handler DLL. Versions for 32 bit and 64 bit OS provided.

Installing/uninstalling the context menu shell handler requires administrator
permissions and executing these scripts requires the user to be a member of
the administrators user group.

In case the scripts fail to work for whatever reason, run the following commands
manually from a command prompt:

Installation:   regsvr32 "full-path-to-DLL-file"

Uninstallation: regsvr32 /u "full-path-to-DLL-file"
