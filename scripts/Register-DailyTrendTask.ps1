[CmdletBinding()]
param(
    [string]$TaskName = "DailySocialTrendReport",
    [string]$RunAt = "08:00",
    [string]$EmailTo = "hammaker.dan@gmail.com",
    [ValidateSet("Outlook","Smtp")]
    [string]$EmailMethod = "Outlook",
    [string]$SmtpUser,
    [string]$SmtpCredentialPath,
    [string]$ProjectRoot = ".."
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$scriptBasePath = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }
$projectFullPath = [System.IO.Path]::GetFullPath((Join-Path $scriptBasePath $ProjectRoot))
$scriptPath = [System.IO.Path]::GetFullPath((Join-Path $projectFullPath "scripts\Invoke-TrendReport.ps1"))

if (-not (Test-Path -LiteralPath $scriptPath)) {
    throw "Collector script not found at $scriptPath"
}

$arguments = "-NoProfile -ExecutionPolicy Bypass -File `"$scriptPath`" -SendEmail -EmailTo `"$EmailTo`" -EmailMethod $EmailMethod"
if ($EmailMethod -eq "Smtp") {
    if (-not $SmtpUser) {
        throw "SmtpUser is required when EmailMethod is Smtp."
    }
    if (-not $SmtpCredentialPath) {
        throw "SmtpCredentialPath is required when EmailMethod is Smtp."
    }

    $arguments += " -SmtpUser `"$SmtpUser`" -SmtpCredentialPath `"$SmtpCredentialPath`""
}

$action = New-ScheduledTaskAction `
    -Execute "powershell.exe" `
    -Argument $arguments

$trigger = New-ScheduledTaskTrigger -Daily -At ([datetime]::ParseExact($RunAt, "HH:mm", $null))
$principal = New-ScheduledTaskPrincipal -UserId $env:USERNAME -LogonType Interactive -RunLevel Limited
$settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -AllowStartIfOnBatteries

Register-ScheduledTask `
    -TaskName $TaskName `
    -Action $action `
    -Trigger $trigger `
    -Principal $principal `
    -Settings $settings `
    -Force | Out-Null

Write-Host "Scheduled task '$TaskName' created. It will run every day at $RunAt."
