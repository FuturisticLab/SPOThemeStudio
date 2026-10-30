<#
.SYNOPSIS
    Launcher for SharePoint Online Theme Studio & Visualizer
.DESCRIPTION
    Launches the interactive WPF GUI for building, testing, and applying custom
    SharePoint Online themes (Add-SPOTheme / Set-SPOWebTheme).
#>

$studioPath = Join-Path $PSScriptRoot "SPOThemeStudio.ps1"
if (-not (Test-Path $studioPath)) {
    Write-Host "Error: SPOThemeStudio.ps1 not found in $PSScriptRoot" -ForegroundColor Red
    return
}

Write-Host "Launching SharePoint Online Theme Studio..." -ForegroundColor Cyan
& $studioPath
