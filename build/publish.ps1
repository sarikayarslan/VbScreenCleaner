# Builds the Windows release package: dist\VbScreenCleaner-win-x64.zip
# The zip name carries no version so README's ".../releases/latest/download/..." link stays valid.
# Usage: powershell -ExecutionPolicy Bypass -File build\publish.ps1 [-Version 1.0.0]
param([string]$Version = "1.0.0")

$ErrorActionPreference = "Stop"
$root = Split-Path $PSScriptRoot -Parent
$out = Join-Path $root "dist\VbScreenCleaner"
$zip = Join-Path $root "dist\VbScreenCleaner-win-x64.zip"

if (Test-Path $out) { Remove-Item -Recurse -Force $out }
if (Test-Path $zip) { Remove-Item -Force $zip }

# Self-contained single-file exes: no .NET installation needed on the target machine.
$common = @("-c", "Release", "-r", "win-x64", "--self-contained", "true",
    "-p:PublishSingleFile=true", "-p:IncludeNativeLibrariesForSelfExtract=true",
    "-p:EnableCompressionInSingleFile=true", "-p:DebugType=none", "-p:Version=$Version", "-o", $out)

dotnet publish (Join-Path $root "src\VbScreenCleaner\VbScreenCleaner.csproj") @common
if ($LASTEXITCODE -ne 0) { throw "GUI publish failed" }
dotnet publish (Join-Path $root "src\VbScreenCleaner.Cli\VbScreenCleaner.Cli.csproj") @common
if ($LASTEXITCODE -ne 0) { throw "CLI publish failed" }

Copy-Item (Join-Path $root "README.md") $out
Copy-Item -Recurse (Join-Path $root "samples\SampleVb5") (Join-Path $out "OrnekProje")

Compress-Archive -Path "$out\*" -DestinationPath $zip
Get-Item $zip | Select-Object Name, @{ n = "MB"; e = { [math]::Round($_.Length / 1MB, 1) } }
