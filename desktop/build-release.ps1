param([string]$OutputDirectory = (Join-Path $PSScriptRoot 'artifacts'))
$ErrorActionPreference = 'Stop'
$msbuild = Get-Command MSBuild.exe -ErrorAction SilentlyContinue
if ($msbuild) { $buildTool = $msbuild.Source } else {
    $vswhere = Join-Path ${env:ProgramFiles(x86)} 'Microsoft Visual Studio/Installer/vswhere.exe'
    $buildTool = & $vswhere -latest -products '*' -requires Microsoft.Component.MSBuild -find 'MSBuild/**/Bin/MSBuild.exe' | Select-Object -First 1
}
if (-not $buildTool) { throw 'Visual Studio 2022 MSBuildが必要です。' }
& $buildTool (Join-Path $PSScriptRoot 'ExcelPdfConverter.sln') /t:Rebuild /p:Configuration=Release /v:minimal /nologo
if ($LASTEXITCODE -ne 0) { throw 'Releaseビルド失敗' }
& $buildTool (Join-Path $PSScriptRoot 'tests/OutputPathsTests.csproj') /t:Rebuild /p:Configuration=Release /v:minimal /nologo
if ($LASTEXITCODE -ne 0) { throw 'テストビルド失敗' }
& (Join-Path $PSScriptRoot 'tests/bin/Release/OutputPathsTests.exe')
if ($LASTEXITCODE -ne 0) { throw 'テスト失敗' }
New-Item -ItemType Directory -Force $OutputDirectory | Out-Null
$binaryRoot = Join-Path $PSScriptRoot 'ExcelPdfConverter/bin/Release'
$packageFiles = @((Join-Path $binaryRoot 'ExcelPdfConverter.exe'),(Join-Path $binaryRoot 'ExcelPdfConverter.Core.dll'),(Join-Path $PSScriptRoot 'README.md'))
$configFile = Join-Path $binaryRoot 'ExcelPdfConverter.exe.config'
if (Test-Path $configFile) { $packageFiles += $configFile }
$archive = Join-Path $OutputDirectory 'ExcelPdfConverter-Windows-Release.zip'
Compress-Archive -LiteralPath $packageFiles -DestinationPath $archive -Force
Write-Output "Release package: $archive"
