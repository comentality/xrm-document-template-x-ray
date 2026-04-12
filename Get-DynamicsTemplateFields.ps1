<#
.SYNOPSIS
    Extracts Dynamics 365 field references from a Word (.docx) document template.

.DESCRIPTION
    Parses a Dynamics 365 Word template (.docx) by treating it as a ZIP archive,
    reads the XML content (document body, headers, footers), and extracts all
    structured document tag (content control) fields that map to Dynamics entities/attributes.

    Outputs field paths in Dynamics notation, e.g.:
        ava_caseenforcementaction.ava_name
        ava_caseenforcementaction/ava_account_ava_caseenforcementaction_licensenumberid/name

.PARAMETER TemplatePath
    The full or relative path to the .docx template file.

.PARAMETER Tree
    When specified, displays unique fields in a sorted folder-tree structure,
    alphabetized at each level. Without this switch, fields are listed in
    document order and duplicates are preserved.

.EXAMPLE
    .\Get-DynamicsTemplateFields.ps1 -TemplatePath ".\MyTemplate.docx"

.EXAMPLE
    .\Get-DynamicsTemplateFields.ps1 -TemplatePath ".\MyTemplate.docx" -Tree
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true, Position = 0)]
    [ValidateScript({
        if (-not (Test-Path $_)) { throw "File not found: $_" }
        if ($_ -notmatch '\.docx$') { throw "File must be a .docx file" }
        $true
    })]
    [string]$TemplatePath,

    [switch]$Tree
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# ── Helpers ──

function Extract-DocxXmlParts {
    param([string]$Path)

    $resolvedPath = (Resolve-Path $Path).Path
    Add-Type -AssemblyName System.IO.Compression.FileSystem

    $zip = [System.IO.Compression.ZipFile]::OpenRead($resolvedPath)
    $parts = @{}

    try {
        foreach ($entry in $zip.Entries) {
            $isRelevant = $entry.FullName -match '^word/(document|header\d*|footer\d*)\.xml$'

            if ($isRelevant) {
                $stream = $entry.Open()
                $reader = New-Object System.IO.StreamReader($stream)
                $parts[$entry.FullName] = $reader.ReadToEnd()
                $reader.Close()
                $stream.Close()
            }
        }
    }
    finally {
        $zip.Dispose()
    }

    return $parts
}

function Convert-XPathToFieldPath {
    <#
    .SYNOPSIS
        Converts a Dynamics template XPath like
          /ns0:DocumentTemplate[1]/ava_caseenforcementaction[1]/ava_account_.../name[1]
        into a clean field path like
          ava_caseenforcementaction/ava_account_.../name
    #>
    param([string]$XPath)

    # Strip namespace prefixes and array indices
    $cleaned = $XPath -replace '\[\d+\]', ''
    $cleaned = $cleaned -replace 'ns\d+:', ''

    # Split into segments, drop empty leading segment and "DocumentTemplate" root
    [string[]]$segments = $cleaned.TrimStart('/') -split '/'
    [System.Collections.Generic.List[string]]$fieldSegments = @()

    foreach ($seg in $segments) {
        if ($seg -eq 'DocumentTemplate') { continue }
        $fieldSegments.Add($seg)
    }

    if ($fieldSegments.Count -eq 0) { return $null }

    # Always use / as separator
    return ($fieldSegments -join '/')
}

function Parse-ContentControls {
    param(
        [string]$XmlContent,
        [string]$PartName
    )

    $fields = [System.Collections.Generic.List[PSObject]]::new()

    # Use regex to extract w:dataBinding xpath values from w:sdtPr blocks.
    # This is more reliable than XmlDocument namespace-qualified attribute access in PS 5.1.
    $sdtPrPattern = '(?s)<w:sdtPr\b[^>]*>(.*?)</w:sdtPr>'
    $sdtPrMatches = [regex]::Matches($XmlContent, $sdtPrPattern)

    foreach ($sdtPrMatch in $sdtPrMatches) {
        $prContent = $sdtPrMatch.Groups[1].Value

        # Extract xpath from w:dataBinding
        $xpathMatch = [regex]::Match($prContent, 'w:xpath="([^"]+)"')
        if (-not $xpathMatch.Success) { continue }
        $xpathValue = $xpathMatch.Groups[1].Value

        # Extract storeItemID
        $storeIdMatch = [regex]::Match($prContent, 'w:storeItemID="([^"]+)"')
        $storeId = if ($storeIdMatch.Success) { $storeIdMatch.Groups[1].Value } else { $null }

        # Extract tag value
        $tagMatch = [regex]::Match($prContent, '<w:tag\s+w:val="([^"]*)"')
        $tagValue = if ($tagMatch.Success) { $tagMatch.Groups[1].Value } else { $null }

        # Extract alias value
        $aliasMatch = [regex]::Match($prContent, '<w:alias\s+w:val="([^"]*)"')
        $aliasValue = if ($aliasMatch.Success) { $aliasMatch.Groups[1].Value } else { $null }

        # Convert xpath to clean field path
        $fieldPath = Convert-XPathToFieldPath -XPath $xpathValue

        $fields.Add([PSCustomObject]@{
            FieldPath = $fieldPath
            Tag       = $tagValue
            Alias     = $aliasValue
            XPath     = $xpathValue
            StoreId   = $storeId
            Location  = $PartName
        })
    }

    return $fields
}

function Show-FieldTree {
    <#
    .SYNOPSIS
        Renders unique field paths as an alphabetized folder tree.
    #>
    param([PSObject[]]$Fields)

    # Build a nested hashtable tree from all field paths
    $root = [ordered]@{}

    foreach ($f in $Fields) {
        $path = $f.FieldPath
        if (-not $path) { continue }

        # Determine segments: always split on /
        [string[]]$segments = $path -split '/'

        $node = $root
        foreach ($seg in $segments) {
            if (-not $node.Contains($seg)) {
                $node[$seg] = [ordered]@{}
            }
            $node = $node[$seg]
        }
    }

    # Recursively render the tree
    function Render-Tree {
        param(
            [System.Collections.Specialized.OrderedDictionary]$Node,
            [string]$Prefix = ''
        )

        [string[]]$keys = @($Node.Keys | Sort-Object)
        $keyCount = $keys.Length
        for ($i = 0; $i -lt $keyCount; $i++) {
            $key = $keys[$i]
            $isLast = ($i -eq ($keyCount - 1))
            $connector = if ($isLast) { [char]0x2514 + [string][char]0x2500 + [char]0x2500 + ' ' } else { [char]0x251C + [string][char]0x2500 + [char]0x2500 + ' ' }
            $childNode = $Node[$key]
            [int]$childCount = @($childNode.Keys).Length

            if ($childCount -gt 0) {
                Write-Host "$Prefix$connector" -NoNewline
                Write-Host $key -ForegroundColor Yellow
            }
            else {
                Write-Host "$Prefix$connector" -NoNewline
                Write-Host $key
            }

            $childPrefix = if ($isLast) { $Prefix + '    ' } else { $Prefix + [char]0x2502 + '   ' }
            Render-Tree -Node $childNode -Prefix $childPrefix
        }
    }

    Render-Tree -Node $root
}

# ── Main ──

Write-Host "`nAnalyzing Dynamics template: $TemplatePath" -ForegroundColor Cyan
Write-Host ("-" * 60)

$xmlParts = Extract-DocxXmlParts -Path $TemplatePath
$allFields = [System.Collections.Generic.List[PSObject]]::new()

foreach ($partName in $xmlParts.Keys) {
    $content = $xmlParts[$partName]
    $sdtFields = Parse-ContentControls -XmlContent $content -PartName $partName
    foreach ($f in $sdtFields) { $allFields.Add($f) }
}

if ($allFields.Count -eq 0) {
    Write-Host "`nNo Dynamics fields found in the template." -ForegroundColor Yellow
    return
}

if ($Tree) {
    # Unique, sorted, folder-tree display
    [array]$uniqueFields = @($allFields | Sort-Object FieldPath -Unique)
    Write-Host "`nFound $($uniqueFields.Length) unique Dynamics field(s):`n" -ForegroundColor Green
    Show-FieldTree -Fields $uniqueFields
    Write-Host ""
}
else {
    Write-Host "`nFound $($allFields.Count) Dynamics field reference(s) (in document order):`n" -ForegroundColor Green
    foreach ($f in $allFields) {
        Write-Host "  $($f.FieldPath)"
    }
    Write-Host ""
}
