# Builds the tool and launches a private XrmToolBox instance that contains nothing but it.
#
# The instance lives in .xtb and is created from scratch, so it cannot disturb the
# XrmToolBox you use for real work: its own Plugins folder, its own settings, its own
# connection list. Delete the folder to undo everything this script did.
#
#   .\xtb.ps1              # build, wire up, launch
#   .\xtb.ps1 -Reset       # throw the instance away and rebuild it
#   .\xtb.ps1 -NoLaunch    # set it up without starting XrmToolBox
#
# The connection points at the active organization of the current pac auth profile. Pass
# -Environment <url> to aim somewhere else.
#
# test_doc.docx is left on the clipboard, because the tool takes a local .docx and that is
# the one thing you always have to go and find.
#
# Everything that is XrmToolBox rather than Document Template X-Ray lives in the XtbSandbox
# module (github.com/comentality/xrmtoolbox-sandbox), shared with the other tools.

param(
    [string]$Environment,
    [string]$XrmToolBoxPath,
    [switch]$Reset,
    [switch]$NoLaunch
)

$ErrorActionPreference = "Stop"

if (-not (Get-Module -ListAvailable XtbSandbox)) {
    throw "XtbSandbox is not installed. Run: Install-Module XtbSandbox -Scope CurrentUser`n(see https://github.com/comentality/xrmtoolbox-sandbox)"
}
Import-Module XtbSandbox

Start-XtbSandbox @PSBoundParameters `
    -InstanceRoot (Join-Path $PSScriptRoot ".xtb") `
    -ProjectPath  (Join-Path $PSScriptRoot "DocumentTemplateXRay\DocumentTemplateXRay.csproj") `
    -ToolName     "Document Template X-Ray" `
    -Clipboard    (Join-Path $PSScriptRoot "test_doc.docx")
