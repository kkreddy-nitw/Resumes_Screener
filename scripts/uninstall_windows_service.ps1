param(
    [string]$Python = "python",
    [string]$VenvPath = ".venv"
)

$ErrorActionPreference = "Stop"

$identity = [Security.Principal.WindowsIdentity]::GetCurrent()
$principal = New-Object Security.Principal.WindowsPrincipal($identity)
$adminRole = [Security.Principal.WindowsBuiltInRole]::Administrator
if (-not $principal.IsInRole($adminRole)) {
    throw "Run this script from an elevated PowerShell window."
}

$projectRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
$venvRoot = if ([System.IO.Path]::IsPathRooted($VenvPath)) {
    $VenvPath
} else {
    Join-Path $projectRoot $VenvPath
}
$venvPython = Join-Path $venvRoot "Scripts\python.exe"
$pythonExe = if (Test-Path $venvPython) {
    $venvPython
} else {
    (Get-Command $Python -ErrorAction Stop).Source
}
$serviceScript = Join-Path $projectRoot "resumes_screener_windows_service.py"

try {
    & $pythonExe $serviceScript stop
} catch {
    Write-Host "Service was not running."
}

& $pythonExe $serviceScript remove
Write-Host "Resumes Screener service removed."
