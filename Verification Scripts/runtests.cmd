@REM Copyright (C) Microsoft Corporation. All rights reserved.

setlocal enabledelayedexpansion

@REM Usage: runtests.cmd

@REM Add .NET 2.0 Framework directory to the path to find msbuild.exe.
set PATH=%windir%\Microsoft.NET\Framework\v2.0.50727;%PATH%

@REM Call msbuild.exe to invoke the runtests target, which runs the unit and scenario tests.
call msbuild.exe /t:RunTests /p:Configuration="Debug-Test" "/p:OutputFile=%OUTPUT_FILE%" "%~dp0runtests.targets"

exit /b !ERRORLEVEL!