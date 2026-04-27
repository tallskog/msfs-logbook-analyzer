# MSFS Logbook Analyzer

MSFS Logbook Analyzer is a WPF desktop application for Windows that reads and analyzes Microsoft Flight Simulator logbook data. It provides a simple interface for loading logbook entries and reviewing flight statistics.

## Features

- Windows desktop application built with WPF
- Targets .NET 8 for Windows
- Easy compilation and execution using the .NET CLI or Visual Studio

## Requirements

- Windows 10 or later
- .NET 8 SDK
- Visual Studio 2022/2023 or newer with WPF desktop development workload (optional)

## Build and Run

### Using the .NET CLI

1. Open a terminal in the project folder:
   ```powershell
   cd c:\<installation location>\MsfsLogbookAnalyzer
   ```
2. Restore packages (if needed):
   ```powershell
   dotnet restore
   ```
3. Build the solution:
   ```powershell
   dotnet build
   ```
4. Run the application:
   ```powershell
   dotnet run --project MsfsLogbookAnalyzer.csproj
   ```

### Using Visual Studio

1. Open `MsfsLogbookAnalyzer.sln` in Visual Studio.
2. Ensure the project is targeting `.NET 8.0` and `UseWPF` is enabled.
3. Build the solution with `Build > Build Solution`.
4. Start debugging or run without debugging.

## Installation

This is a simple desktop application without a dedicated repository-provided installer.

- Build locally using the instructions above.
- The executable is published to `bin\Debug\net8.0-windows10.0.19041.0\` after a successful build.

## Creating an installer package

If you want to install the app like a normal Windows application, you can package it as an MSIX installer or create an MSI installer from the published output.

### Option 1: MSIX package (recommended for Windows)

1. Install the `Windows Applicaton Packaging Project` workload in Visual Studio, or use the MSIX Packaging Tool.
2. In Visual Studio, add a new project:
   - `Create a new project` > `Windows Application Packaging Project`
   - Set the Target version to at least `Windows 10, version 19041`
3. Add a reference from the packaging project to `MsfsLogbookAnalyzer`.
4. Configure package metadata in the packaging project manifest:
   - Package display name
   - Publisher
   - Package identity
5. Build the packaging project.
6. The output will include an `.msix` or `.appx` package that you can install on Windows.
7. After installation, the app appears in the Start menu and can be launched like any other installed application.

### Option 2: MSI installer or third-party installer

1. Publish the app for release:
   ```powershell
   dotnet publish -c Release -r win-x64 --self-contained false /p:PublishSingleFile=false
   ```
2. Use an installer tool such as WiX Toolset, Inno Setup, or Advanced Installer.
3. Point the installer to the published output folder (under `bin\Release\net8.0-windows10.0.19041.0\win-x64\publish\`).
4. Configure installation location, shortcuts, and registry entries as needed.

> Note: This repository does not currently include an installer project. The above instructions show how to package the existing WPF app for installation.

## License

This project is licensed under the terms of the `LICENSE` file in this repository.
