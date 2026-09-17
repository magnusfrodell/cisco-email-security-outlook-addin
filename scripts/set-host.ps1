<#
  Copyright (c) 2026 Cisco and/or its affiliates.

  This software is licensed to you under the terms of the Cisco Sample
  Code License, Version 1.1 (the "License"). You may obtain a copy of the
  License at

                 https://developer.cisco.com/docs/licenses

  All use of the material herein must be in accordance with the terms of
  the License. All rights not expressly granted by the License are
  reserved. Unless required by applicable law or agreed to separately in
  writing, software distributed under the License is distributed on an "AS
  IS" BASIS, WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express
  or implied.
#>
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
