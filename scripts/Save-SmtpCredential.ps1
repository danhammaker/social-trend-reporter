[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$Username,
    [Parameter(Mandatory)]
    [string]$Password,
    [string]$OutputPath = "..\secrets\gmail-credential.json"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$scriptBasePath = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }

function Get-AbsolutePath {
    param(
        [Parameter(Mandatory)]
        [string]$Path,
        [Parameter(Mandatory)]
        [string]$BasePath
    )

    if ([System.IO.Path]::IsPathRooted($Path)) {
        return [System.IO.Path]::GetFullPath($Path)
    }

    return [System.IO.Path]::GetFullPath((Join-Path $BasePath $Path))
}

$resolvedPath = Get-AbsolutePath -Path $OutputPath -BasePath $scriptBasePath
$resolvedDir = Split-Path -Parent $resolvedPath
New-Item -ItemType Directory -Path $resolvedDir -Force | Out-Null

$securePassword = ConvertTo-SecureString -String $Password -AsPlainText -Force
$encryptedPassword = $securePassword | ConvertFrom-SecureString

[pscustomobject]@{
    username = $Username
    password = $encryptedPassword
} | ConvertTo-Json | Set-Content -LiteralPath $resolvedPath -Encoding UTF8

Write-Host "SMTP credential saved to $resolvedPath"
