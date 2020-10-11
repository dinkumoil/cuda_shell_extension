# Windows Explorer context menu extension for CudaText

Builds for 32 and 64 bit Windows available

Author: Andreas Heim, 2019-2020

Developed with _Lazarus_ IDE and _Free Pascal Compiler_.


# Features

This context menu extension for Windows Explorer installs an entry into Explorer's context menu to be able to open an arbitrary number of files with _CudaText_. It overcomes the 15 files limit that exists for the Explorer context menu entry that has been installed via _cuda_explorer_integration_ plugin which uses a _shell\CudaText\command_ registry entry.

Additionally you are able to open a folder and even a whole drive with _CudaText_. The _Project Manager_ plugin will add the folder/the drive's root folder to the project loaded at startup or to a new project.


# Known issues

When Windows Explorer is compositing the context menu of **a drive or a folder right-clicked in a window's navigation pane** (the folder tree structure at the left side of an Explorer window), for what ever reason the context menu handler's COM interface is called twice. Thus, in the context menu of the items mentioned aforehand there will be two _CudaText_ entries. Since both of these entries work as expected, this is only a minor flaw I won't fix.


# Installation

The DLL file of the Explorer context menu extension can be downloaded [here](https://github.com/dinkumoil/cuda_shell_extension/releases), choose the one that fits the **bitness of your Windows** installation. Copy the downloaded DLL file to the directory of your _cudatext.exe_.

The context menu extension has to be installed via `regsvr32.exe` which requires administrator permissions. The `install` directory of this repository contains batch scripts for its automatic installation and uninstallation. Again, choose the ones that fit the **bitness of your Windows** installation. Copy the downloaded scripts to the directory of your _cudatext.exe_ as well. Then start `install_shellXX.cmd` by double-click.

The scripts try to automate the whole process of installing/uninstalling by requesting administrator permissions. If automatic installation/uninstallation fails, run the following commands from a console window that has been started via `Run as administrator`:

**Installation:** `regsvr32 "<path-to-DLL-file>"`

**Uninstallation:** `regsvr32 /u "<path-to-DLL-file>"`

After uninstalling the context menu extension you won't be able to immediately delete its DLL file. To do so you have to log out from Windows and log in again or, in the brute-force manner, kill all `explorer.exe` processes in Windows Task Manager and restart Explorer via `(menu) File -> New task`.


# History

v1.1 - October 2020
- enhanced: The _CudaText_ context menu entry is also available in the context menu of directories and drives.


v1.0 - December 2019
- Initial version
