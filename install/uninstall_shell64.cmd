:: ****************************************************************************
::
:: Uninstaller for CudaText context menu handler for 64 bit Windows Explorer
::
:: Author:  Andreas Heim, 2019-2020
:: Website: https://github.com/dinkumoil/cuda_shell_extension
::
::
:: In case (de-)installation doesn't work successfully on your machine
:: have a look at the project's website to obtain detailed instructions
:: on how to do it manually.
::
::
:: This program is free software; you can redistribute it and/or modify
:: it under the terms of the Mozilla Public License Version 2.0.
::
:: This program is distributed in the hope that it will be useful,
:: but WITHOUT ANY WARRANTY; without even the implied warranty of
:: MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.
::
:: ****************************************************************************

@echo off & setlocal

::------------------------------------------------------------------------------
:: Basic configuration
::------------------------------------------------------------------------------
::Path for temporary elevation script
set "ElevateScript=%Temp%\ElevateCudaTextShellInstaller.vbs"

::Path for temporary script to retrieve system's ANSI code page
set "GetAnsiCodePageScript=%Temp%\GetAnsiCodePage.vbs"

::Name of context menu handler DLL file
set "ShellHandler=cudatext_shell64.dll"


::------------------------------------------------------------------------------
:: Startup sequence
::------------------------------------------------------------------------------
::Set working directory to script's path
pushd "%~dp0"

::Check for admin permissions and restart elevated if required
call :CheckForAdminPermissions || (
  call :RestartElevated "%~f0"
  goto :Terminate
)


::------------------------------------------------------------------------------
:: Shell context menu handler deinstallation
::------------------------------------------------------------------------------
regsvr32 /s /u "%CD%\%ShellHandler%"

::Clean up
del "%ElevateScript%" 1>NUL 2>&1


::------------------------------------------------------------------------------
:: Script termination
::------------------------------------------------------------------------------
:Terminate
popd
exit /b 0



::==============================================================================
:: Subroutines
::==============================================================================

:CheckForAdminPermissions
  net session 1>NUL 2>&1
  if ERRORLEVEL 1 exit /b 1
exit /b 0


:RestartElevated
  ::Get system's ANSI and OEM code page and set console's code page to ANSI code page.
  ::This is required if this script is stored in a path that contains characters
  ::with different code points in those code pages.
  call :GetAnsiCodePage
  for /f "tokens=2 delims=.:" %%a in ('chcp') do set /a "OEMCP=%%a"
  if "%ACP%" neq "" if "%ACP%" neq "0" chcp %ACP% > NUL

  > "%ElevateScript%" echo.Set objShell = CreateObject("Shell.Application")
  >>"%ElevateScript%" echo.
  >>"%ElevateScript%" echo.strApplication = "cmd.exe"
  >>"%ElevateScript%" echo.strArguments   = "/c ""%~1"""
  >>"%ElevateScript%" echo.
  >>"%ElevateScript%" echo.objShell.ShellExecute strApplication, strArguments, "", "runas", 1

  ::Restore OEM code page
  if "%OEMCP%" neq "" if "%OEMCP%" neq "0" chcp %OEMCP% > NUL

  cscript /nologo "%ElevateScript%"
exit /b 0


:GetAnsiCodePage
  > "%GetAnsiCodePageScript%" echo.Set objWMI = GetObject("winmgmts:root\cimv2:Win32_OperatingSystem=@")
  >>"%GetAnsiCodePageScript%" echo.WScript.Echo objWMI.CodeSet

  set "ACP="
  for /f "delims=" %%a in ('cscript.exe /nologo "%GetAnsiCodePageScript%" 2^>NUL') do set "ACP=%%~a"

  del "%GetAnsiCodePageScript%"
exit /b 0
