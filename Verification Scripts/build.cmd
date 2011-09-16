@REM Copyright (C) Microsoft Corporation. All rights reserved.

@if not "%_echo%"=="1" (echo off)
setlocal ENABLEDELAYEDEXPANSION

@REM Usage: build.cmd [Rebuild|Clean] [Debug|Release|FxCop-Release|Debug-Test]

@REM Add .NET 2.0 Framework directory to the path to find msbuild.exe.
set PATH=%windir%\Microsoft.NET\Framework\v3.5;%PATH%

call msbuild.exe "/t:%~1" "/p:Configuration=%~2" "/p:TreatWarningsAsErrors=false" "..\ResourceRefactor.sln"
exit /b "!ERRORLEVEL!"