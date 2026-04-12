$pluginDir = "C:\Users\kk\Downloads\XrmToolbox\Plugins"
if (-not (Test-Path $pluginDir)) { New-Item -ItemType Directory -Path $pluginDir | Out-Null }

$source = Join-Path $PSScriptRoot "DocumentXRay\bin\Debug\net48\DocumentXRay.dll"
Copy-Item $source -Destination $pluginDir -Force

Write-Host "Deployed DocumentXRay.dll to $pluginDir"
