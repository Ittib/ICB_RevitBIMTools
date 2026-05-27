@echo off
echo =======================================================
echo Building Navisworks Clash Reporter Plugin
echo =======================================================

:: Path to the standalone .NET SDK
set DOTNET_EXE=D:\itti\mtc\dotnet-sdk-10\dotnet.exe

:: Build the project
"%DOTNET_EXE%" build -c Release

if %errorlevel% neq 0 (
    echo.
    echo [ERROR] Build failed! Check the errors above.
    pause
    exit /b %errorlevel%
)

echo.
echo [SUCCESS] Build completed successfully.
echo Copying files to Navisworks Plugins directory...

:: Create target directory if it doesn't exist
set DEST=%APPDATA%\Autodesk\Navisworks Manage 2023\Plugins\NavisworksClashReporter
if not exist "%DEST%" mkdir "%DEST%"

:: Copy the required DLLs
copy /Y "bin\Release\net48\NavisworksClashReporter.dll" "%DEST%"
copy /Y "bin\Release\net48\HtmlAgilityPack.dll" "%DEST%"

echo.
echo [DONE] Files copied successfully! You can now open Navisworks.
pause
