@echo off & setlocal

::------------------------------------------------------------------------------
:: Basic configuration
::------------------------------------------------------------------------------
::Path for temporary elevation script
set "ElevateScript=%Temp%\ElevateCudaTextShellInstaller.vbs"

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
del "%ElevateScript%" > NUL 2>&1


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
  > "%ElevateScript%" echo.Set objShell = CreateObject("Shell.Application")
  >>"%ElevateScript%" echo.
  >>"%ElevateScript%" echo.strApplication = "cmd.exe"
  >>"%ElevateScript%" echo.strArguments   = "/c """"%~1"" %~2"""
  >>"%ElevateScript%" echo.
  >>"%ElevateScript%" echo.objShell.ShellExecute strApplication, strArguments, "", "runas", 1

  cscript /nologo "%ElevateScript%"
exit /b 0
