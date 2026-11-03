# Builds the Windows release files: dist\VbScreenCleaner.exe (GUI) and dist\VbScreenCleaner.Cli.exe (CLI).
# Each exe is uploaded as its own GitHub Release asset; the names carry no version so README's
# ".../releases/latest/download/<name>" links stay valid.
# Usage: powershell -ExecutionPolicy Bypass -File build\publish.ps1 [-Version 1.0.0]
param([string]$Version = "1.0.0")

$ErrorActionPreference = "Stop"
$root = Split-Path $PSScriptRoot -Parent
$dist = Join-Path $root "dist"

if (Test-Path $dist) { Remove-Item -Recurse -Force $dist }

# Self-contained single-file exes: no .NET installation needed on the target machine.
$common = @("-c", "Release", "-r", "win-x64", "--self-contained", "true",
    "-p:PublishSingleFile=true", "-p:IncludeNativeLibrariesForSelfExtract=true",
    "-p:EnableCompressionInSingleFile=true", "-p:DebugType=none", "-p:Version=$Version", "-o", $dist)

dotnet publish (Join-Path $root "src\VbScreenCleaner\VbScreenCleaner.csproj") @common
if ($LASTEXITCODE -ne 0) { throw "GUI publish failed" }
dotnet publish (Join-Path $root "src\VbScreenCleaner.Cli\VbScreenCleaner.Cli.csproj") @common
if ($LASTEXITCODE -ne 0) { throw "CLI publish failed" }

Get-ChildItem $dist -Filter *.exe | Select-Object Name, @{ n = "MB"; e = { [math]::Round($_.Length / 1MB, 1) } }
