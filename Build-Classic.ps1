$ErrorActionPreference = 'Stop'

$projectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $projectRoot

Write-Host '1. Kompilerar TypeScript via SPFx build...' -ForegroundColor Cyan
& gulp build

if ($LASTEXITCODE -ne 0) {
    throw "gulp build misslyckades med exit code $LASTEXITCODE"
}

Write-Host '2. Bygger Classic-bundle...' -ForegroundColor Cyan

$webpackCmd = Join-Path $projectRoot 'node_modules\.bin\webpack.cmd'

if (-not (Test-Path $webpackCmd)) {
    throw "webpack.cmd hittades inte: $webpackCmd"
}

& $webpackCmd --config webpack.classic.js

if ($LASTEXITCODE -ne 0) {
    throw "Classic webpack-build misslyckades med exit code $LASTEXITCODE"
}

$outputFile = Join-Path $projectRoot 'dist\classic\ESignClassic.js'

if (-not (Test-Path $outputFile)) {
    throw "Classic-bundlen skapades inte: $outputFile"
}

Write-Host ''
Write-Host 'Classic-bundle skapad:' -ForegroundColor Green
Write-Host $outputFile
