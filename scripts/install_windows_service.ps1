param(
    [string]$Python = "python",
    [string]$VenvPath = ".venv",
    [string]$HostName = "0.0.0.0",
    [int]$Port = 80
)

$ErrorActionPreference = "Stop"

$identity = [Security.Principal.WindowsIdentity]::GetCurrent()
$principal = New-Object Security.Principal.WindowsPrincipal($identity)
$adminRole = [Security.Principal.WindowsBuiltInRole]::Administrator
if (-not $principal.IsInRole($adminRole)) {
    throw "Run this script from an elevated PowerShell window."
}

$projectRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
Set-Location $projectRoot

$basePython = (Get-Command $Python -ErrorAction Stop).Source
$venvRoot = if ([System.IO.Path]::IsPathRooted($VenvPath)) {
    $VenvPath
} else {
    Join-Path $projectRoot $VenvPath
}

if (-not (Test-Path $venvRoot)) {
    & $basePython -m venv $venvRoot
}

$pythonExe = Join-Path $venvRoot "Scripts\python.exe"

& $pythonExe -m pip install -r (Join-Path $projectRoot "requirements.txt")

$pywinPostInstall = Join-Path $venvRoot "Scripts\pywin32_postinstall.py"
if (Test-Path $pywinPostInstall) {
    & $pythonExe $pywinPostInstall -install
}

$config = @{
    ProjectRoot = $projectRoot
    PythonExe = $pythonExe
    Host = $HostName
    Port = "$Port"
} | ConvertTo-Json

$configPath = Join-Path $projectRoot "windows_service_config.json"
$config | Set-Content -Path $configPath -Encoding UTF8

& $pythonExe (Join-Path $projectRoot "resumes_screener_windows_service.py") install --startup auto
& $pythonExe (Join-Path $projectRoot "resumes_screener_windows_service.py") update --startup auto
& $pythonExe (Join-Path $projectRoot "resumes_screener_windows_service.py") start

Write-Host "Resumes Screener service installed and started on http://$HostName`:$Port"
Write-Host "Open firewall access to TCP port $Port if other systems cannot connect."
