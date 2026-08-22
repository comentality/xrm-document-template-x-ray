# The tool on a slow link, without XrmToolBox and without an environment.
#
# It asks what the window is doing while somebody waits: which buttons are still live, whether the
# template list and the field pane are describing the same file, whether the tool is still asking
# about a template nobody is looking at any more, and how long the window went without being able
# to draw itself.
#
#   .\slow.ps1                          # every scenario
#   .\slow.ps1 -Scenario cancel         # one
#   .\slow.ps1 -Scenario cancel,big-template
#   .\slow.ps1 -NoBuild                 # reuse the last build
#
# It drives the real control against a fake Dataverse that takes seconds to answer, handed in
# through the connection property XrmToolBox would have set. The templates it serves are the real
# fixtures, as bytes: the tool unzips and scans them with the extractor it ships, and resolving
# their display names costs a round trip per table here the way it does for real. Each scenario
# gets a window and a screenshot per beat of its own under tests\.slow, and the exit code is the
# number of scenarios with findings.
#
# The windows are real windows, but they open past the edge of the desktop and never take the
# keyboard: the shots are the window's own pixels and the gestures are performed on the controls
# rather than typed at them, so the machine stays usable while a suite runs. Set
# DTX_HARNESS_ONSCREEN=1 to watch a scenario play out instead.
#
# One scenario needs a template that is big rather than complicated. It is built once into
# %TEMP%\dtx-slow-harness and reused, which is why the first run of it takes a few seconds longer.

param(
    [string[]]$Scenario,
    [string]$OutputDir,
    [switch]$NoBuild
)

$ErrorActionPreference = "Stop"

$project = Join-Path $PSScriptRoot "DocumentTemplateXRay.SlowHarness\DocumentTemplateXRay.SlowHarness.csproj"
$exe     = Join-Path $PSScriptRoot "DocumentTemplateXRay.SlowHarness\bin\Debug\net48\DocumentTemplateXRay.SlowHarness.exe"
if (-not $OutputDir) { $OutputDir = Join-Path $PSScriptRoot ".slow" }

if (-not $NoBuild) {
    dotnet build $project -c Debug -v q --nologo
    if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }
}

$arguments = @("`"$OutputDir`"") + $Scenario

# Out of process with a timeout: a scenario that deadlocks the message loop - which a dialog
# nobody answers would - should fail the run rather than hang it. The budget is the whole suite's,
# and the suite is about a minute of deliberate waiting.
$p = Start-Process $exe -ArgumentList $arguments -NoNewWindow -PassThru
if (-not $p.WaitForExit(300000)) {
    $p.Kill()
    Write-Host "TIMED OUT" -ForegroundColor Red
    exit 1
}

exit $p.ExitCode
