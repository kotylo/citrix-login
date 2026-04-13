#!/usr/bin/env pwsh
# Remote Auto-Login Launcher
# Installs dependencies (first run only) and launches the login script.

$ErrorActionPreference = "Stop"
$scriptDir = $PSScriptRoot

Push-Location $scriptDir
try {
    if (-not (Test-Path "node_modules")) {
        Write-Host "Installing dependencies (first run)..." -ForegroundColor Cyan
        npm install
        Write-Host ""
    }
    node --env-file=.env login.mjs
} finally {
    Pop-Location
}
