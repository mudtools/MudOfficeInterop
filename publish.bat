@echo off

rem Set project root directory
set "PROJECT_ROOT=%~dp0"
set "OUTPUT_DIR=%PROJECT_ROOT%bin"
set "NUGET_DIR=%PROJECT_ROOT%nuget"
set "VERSION=2.0.6"

echo Preparing to generate OfficeInterop NuGet packages version %VERSION%
echo ================================

rem Clean previous build output
echo Cleaning previous build output...
if exist "%OUTPUT_DIR%" rd /s /q "%OUTPUT_DIR%"
if exist "%NUGET_DIR%" rd /s /q "%NUGET_DIR%"
mkdir "%OUTPUT_DIR%"
mkdir "%NUGET_DIR%"

rem Build projects
echo Building projects...
dotnet build "%PROJECT_ROOT%MudTools.OfficeInterop.Excel\MudTools.OfficeInterop.Excel.csproj" -c Release
dotnet build "%PROJECT_ROOT%MudTools.OfficeInterop.Word\MudTools.OfficeInterop.Word.csproj" -c Release
dotnet build "%PROJECT_ROOT%MudTools.OfficeInterop.PowerPoint\MudTools.OfficeInterop.PowerPoint.csproj" -c Release
dotnet build "%PROJECT_ROOT%MudTools.OfficeInterop.Vbe\MudTools.OfficeInterop.Vbe.csproj" -c Release

rem Check if build succeeded
if %errorlevel% neq 0 (
    echo Build failed!
    pause
    exit /b %errorlevel%
)

rem Generate NuGet packages
echo Generating NuGet packages...
dotnet pack "%PROJECT_ROOT%MudTools.OfficeInterop.Excel\MudTools.OfficeInterop.Excel.csproj" -c Release -o "%NUGET_DIR%"
dotnet pack "%PROJECT_ROOT%MudTools.OfficeInterop.Word\MudTools.OfficeInterop.Word.csproj" -c Release -o "%NUGET_DIR%"
dotnet pack "%PROJECT_ROOT%MudTools.OfficeInterop.PowerPoint\MudTools.OfficeInterop.PowerPoint.csproj" -c Release -o "%NUGET_DIR%"
dotnet pack "%PROJECT_ROOT%MudTools.OfficeInterop.Vbe\MudTools.OfficeInterop.Vbe.csproj" -c Release -o "%NUGET_DIR%"

rem Check if packing succeeded
if %errorlevel% neq 0 (
    echo Packing failed!
    pause
    exit /b %errorlevel%
)

rem List generated NuGet packages
echo Generated NuGet packages:
dir "%NUGET_DIR%"

echo ================================
echo NuGet packages generated successfully!
echo Packages location: %NUGET_DIR%
echo ================================
echo 
echo Example commands to publish to NuGet:
echo dotnet nuget push "%NUGET_DIR%\MudTools.OfficeInterop.Excel.%VERSION%.nupkg" -k YOUR_API_KEY -s https://api.nuget.org/v3/index.json
echo dotnet nuget push "%NUGET_DIR%\MudTools.OfficeInterop.Word.%VERSION%.nupkg" -k YOUR_API_KEY -s https://api.nuget.org/v3/index.json
echo dotnet nuget push "%NUGET_DIR%\MudTools.OfficeInterop.PowerPoint.%VERSION%.nupkg" -k YOUR_API_KEY -s https://api.nuget.org/v3/index.json
echo dotnet nuget push "%NUGET_DIR%\MudTools.OfficeInterop.Vbe.%VERSION%.nupkg" -k YOUR_API_KEY -s https://api.nuget.org/v3/index.json
echo ================================
pause