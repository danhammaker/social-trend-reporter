[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$Token,
    [string]$OutputPath = "..\secrets\x-bearer-token.txt"
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

$secureToken = ConvertTo-SecureString -String $Token -AsPlainText -Force
$secureToken | ConvertFrom-SecureString | Set-Content -LiteralPath $resolvedPath -Encoding UTF8

Write-Host "API token saved to $resolvedPath"
