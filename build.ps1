$project = Join-Path $PSScriptRoot "DocumentTemplateXRay\DocumentTemplateXRay.csproj"

dotnet build $project -c Debug
if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }

$pluginDir = "C:\Users\kk\Downloads\XrmToolbox\Plugins"
if (-not (Test-Path $pluginDir)) { New-Item -ItemType Directory -Path $pluginDir | Out-Null }

$source = Join-Path $PSScriptRoot "DocumentTemplateXRay\bin\Debug\net48\DocumentTemplateXRay.dll"
Copy-Item $source -Destination $pluginDir -Force

Write-Host "Built and deployed DocumentTemplateXRay.dll to $pluginDir"
