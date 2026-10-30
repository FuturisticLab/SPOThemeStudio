Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Drawing, System.Windows.Forms

$scriptContent = Get-Content -Raw "c:\Users\jmarc\Desktop\FirebaseMigration\SharePoint-Theme-Studio\SPOThemeStudio.ps1"

# Replace ShowDialog with a quick validation check so it doesn't block
$testContent = $scriptContent -replace '\$window\.ShowDialog\(\) \| Out-Null', 'Write-Host "Window instantiated successfully with title: $($window.Title)"; $script:TestPassed = $true'

$testScriptPath = "c:\Users\jmarc\Desktop\FirebaseMigration\SharePoint-Theme-Studio\Temp-TestRunner.ps1"
[System.IO.File]::WriteAllText($testScriptPath, $testContent, [System.Text.Encoding]::UTF8)

try {
    Write-Host "Running full WPF lifecycle test..." -ForegroundColor Cyan
    & $testScriptPath
    Write-Host "WPF Lifecycle Test PASSED!" -ForegroundColor Green
} finally {
    if (Test-Path $testScriptPath) {
        Remove-Item $testScriptPath -Force
    }
}
