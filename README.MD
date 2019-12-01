# Windows Explorer context menu extension for CudaText

Builds for 32 and 64 bit Windows available

Author: Andreas Heim, 2019

Developed with _Lazarus_ IDE and _Free Pascal Compiler_.


# Features

This context menu extension for Windows Explorer installs an entry into Explorer's context menu to be able to open an arbitrary number of files with _CudaText_. It overcomes the 15 files limit that exists for the Explorer context menu entry that has been installed via _cuda_explorer_integration_ plugin which uses a _shell\CudaText\command_ registry entry.


# Installation

The DLL file of the Explorer context menu extension can be downloaded [here](https://github.com/dinkumoil/cuda_shell_extension/releases), choose the one that fits the bitness of your **Windows** installation. Copy the downloaded DLL file to the directory of your _cudatext.exe_.

The context menu extension has to be installed via `regsvr32.exe` which requires administrator permissions. The `install` directory of this repository contains batch scripts for its automatic installation and uninstallation. Again, choose the ones that fit the bitness of your **Windows** installation. Copy the downloaded scripts to the directory of your _cudatext.exe_ as well. Then start `install_shellXX.cmd` by double-click.

The scripts try to automate the whole process of installing/uninstalling by requesting administrator permissions. If automatic installation/uninstallation fails, run the following commands from a console window that has been started via `Run as administrator`:

**Installation:** `regsvr32 "<path-to-DLL-file>"`

**Uninstallation:** `regsvr32 /u "<path-to-DLL-file>"`

After uninstalling the context menu extension you won't be able to immediatly delete its DLL file. To do so you have to kill all `explorer.exe` processes in Windows Taskmanager and restart Explorer via `(menu) File -> New task`.
