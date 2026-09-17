<#
.SYNOPSIS
  Replace the placeholder host in src/manifest.xml with the real HTTPS host.
.EXAMPLE
  .\scripts\set-host.ps1 -Host addin.firma.dk
  .\scripts\set-host.ps1 -Host addin.firma.dk -Out .\manifest-prod.xml
#>
param(
  [Parameter(Mandatory = $true)] [string] $Host,
  [string] $Out
)
$ErrorActionPreference = "Stop"
$placeholder = "addin.example.com"
$src = Join-Path (Split-Path -Parent $PSScriptRoot) "src\manifest.xml"
if (-not $Out) { $Out = $src }

$xml = Get-Content -Raw -Encoding UTF8 $src
if ($xml -notmatch [regex]::Escape($placeholder)) {
  Write-Warning "placeholder '$placeholder' not found in $src - nothing to do"
}
$xml = $xml.Replace("https://$placeholder", "https://$Host")
[System.IO.File]::WriteAllText($Out, $xml, (New-Object System.Text.UTF8Encoding($false)))
$count = ([regex]::Matches($xml, [regex]::Escape("https://$Host"))).Count
Write-Host "wrote $Out with host https://$Host ($count occurrences)"
