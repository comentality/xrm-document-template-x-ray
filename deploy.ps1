$pluginDir = "C:\Users\kk\Downloads\XrmToolbox\Plugins"
if (-not (Test-Path $pluginDir)) { New-Item -ItemType Directory -Path $pluginDir | Out-Null }

$source = Join-Path $PSScriptRoot "DocumentTemplateXRay\bin\Debug\net48\DocumentTemplateXRay.dll"
Copy-Item $source -Destination $pluginDir -Force

Write-Host "Deployed DocumentTemplateXRay.dll to $pluginDir"
