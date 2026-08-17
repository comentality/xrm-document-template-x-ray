<#
.SYNOPSIS
    Builds the .docx test fixtures for Document Template X-Ray.

.DESCRIPTION
    A Dynamics 365 Word template is an ordinary .docx whose content controls carry a
    w15:dataBinding pointing into a customXml part, e.g.

        /ns0:DocumentTemplate[1]/account[1]/account_primary_contact[1]/description[1]

    This script writes such packages from scratch, so every fixture is a real Word
    document you can open, not a stub. Each one is built to make one behaviour of the
    tool visible in a screenshot -- see fixtures/README.md for what each proves.

    The field paths use out-of-the-box tables, columns and relationship schema names
    (account, contact, systemuser, task), so display-name resolution works against any
    Dataverse environment without setup.

.PARAMETER OutputDirectory
    Where to write the .docx files. Defaults to this script's folder.

.PARAMETER Name
    Build only the fixtures whose file name matches this wildcard.

.EXAMPLE
    .\New-Fixtures.ps1

.EXAMPLE
    .\New-Fixtures.ps1 -Name "01-*"
#>

[CmdletBinding()]
param(
    [string]$OutputDirectory = $PSScriptRoot,
    [string]$Name = "*"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

# The customXml part Word binds against. One id per document is all we need; Word only
# requires that the sdt's storeItemID matches the ds:datastoreItem it should read from.
$StoreItemId = '{F2958ECF-DBAE-45D0-8444-3C06593CE721}'

# ── Body DSL ──
#
# A fixture body is a list of blocks. Text in a line may embed a field reference as
# {table/relationship/column}; that becomes a bound content control whose visible text is
# the column name -- which is exactly what Word shows you and exactly why two fields from
# different tables are impossible to tell apart on the page.

function Line {
    param([string]$Text, [switch]$Heading, [switch]$Title)
    [PSCustomObject]@{ Kind = 'Line'; Text = $Text; Style = $(if ($Title) { 'Title' } elseif ($Heading) { 'Heading1' } else { $null }) }
}

function Repeat {
    param([string]$Path, [object[]]$Children)
    [PSCustomObject]@{ Kind = 'Repeat'; Path = $Path; Children = $Children }
}

# ── Fixtures ──

$Fixtures = @(
    [PSCustomObject]@{
        File        = '01-duplicate-column-names.docx'
        Title       = 'Account Briefing'
        RootEntity  = 'account'
        TypeCode    = 1
        Description = 'Same column name on four different tables.'
        Body        = @(
            Line 'Account Briefing' -Title
            Line ''
            Line '{account/name}'
            Line '{account/address1_city}'
            Line '{account/emailaddress1} / {account/telephone1}'
            Line ''
            Line 'Parent company' -Heading
            Line '{account/account_parent_account/name}'
            Line '{account/account_parent_account/address1_city}'
            Line ''
            Line 'Primary contact' -Heading
            Line '{account/account_primary_contact/fullname}'
            Line '{account/account_primary_contact/address1_city}'
            Line '{account/account_primary_contact/emailaddress1} / {account/account_primary_contact/telephone1}'
            Line ''
            Line 'Notes' -Heading
            Line '{account/description}'
            Line '{account/account_parent_account/description}'
            Line '{account/account_primary_contact/description}'
            Line ''
            Line 'Prepared by {account/user_accounts/fullname}, {account/user_accounts/title}, {account/user_accounts/address1_city}'
            Line 'Record created {account/createdon}'
        )
    }

    [PSCustomObject]@{
        File        = '02-repeating-sections.docx'
        Title       = 'Account Activity Report'
        RootEntity  = 'account'
        TypeCode    = 1
        Description = 'Two repeating sections whose children collide with the root columns.'
        Body        = @(
            Line 'Account Activity Report' -Title
            Line ''
            Line '{account/name} - {account/description} - {account/createdon}'
            Line ''
            Line 'Contacts' -Heading
            Repeat 'account/contact_customer_accounts' @(
                Line '{account/contact_customer_accounts/fullname} ({account/contact_customer_accounts/jobtitle})'
                Line '{account/contact_customer_accounts/description}'
                Line '{account/contact_customer_accounts/emailaddress1} - added {account/contact_customer_accounts/createdon}'
            )
            Line ''
            Line 'Open tasks' -Heading
            Repeat 'account/Account_Tasks' @(
                Line '{account/Account_Tasks/subject} - due {account/Account_Tasks/scheduledend}'
                Line '{account/Account_Tasks/description}'
                Line 'Raised {account/Account_Tasks/createdon}'
            )
            Line ''
            Line 'Report owner: {account/user_accounts/fullname}'
        )
    }

    [PSCustomObject]@{
        File        = '03-header-footer.docx'
        Title       = 'Letterhead'
        RootEntity  = 'account'
        TypeCode    = 1
        Description = 'The same column name in the body, the header and the footer.'
        Header      = @(
            Line '{account/name} - {account/address1_city}'
        )
        Footer      = @(
            Line '{account/account_parent_account/name} - {account/account_primary_contact/address1_city}'
        )
        Body        = @(
            Line 'Letterhead' -Title
            Line ''
            Line 'Dear {account/account_primary_contact/fullname},'
            Line ''
            Line '{account/description}'
            Line ''
            Line 'Yours sincerely,'
            Line '{account/user_accounts/fullname}'
        )
    }

    [PSCustomObject]@{
        File        = '04-unresolvable-fields.docx'
        Title       = 'Stale Template'
        RootEntity  = 'account'
        TypeCode    = 1
        Description = 'Bindings to a table and columns that no longer exist.'
        Body        = @(
            Line 'Stale Template' -Title
            Line ''
            Line '{account/name}'
            Line '{account/xray_columnthatwasdeleted}'
            Line '{account/xray_relationshipthatwasdeleted/name}'
            Line '{xray_tablethatwasdeleted/xray_name}'
            Line '{account/account_primary_contact/xray_columnthatwasdeleted}'
        )
    }

    [PSCustomObject]@{
        File        = '05-no-fields.docx'
        Title       = 'Plain Document'
        RootEntity  = $null
        TypeCode    = 0
        Description = 'An ordinary Word file with no Dynamics bindings at all.'
        Body        = @(
            Line 'Plain Document' -Title
            Line ''
            Line 'Nothing in this document is bound to Dynamics. The tool should say so rather than fail.'
        )
    }
)

# ── OOXML emitters ──

$FieldToken = [regex]'\{([^{}]+)\}'

function ConvertTo-XPath {
    param([string]$FieldPath, [switch]$NoTrailingIndex)

    $segments = $FieldPath -split '/'
    $xpath = '/ns0:DocumentTemplate[1]'
    for ($i = 0; $i -lt $segments.Count; $i++) {
        $xpath += '/' + $segments[$i]
        if (-not ($NoTrailingIndex -and $i -eq $segments.Count - 1)) { $xpath += '[1]' }
    }
    return $xpath
}

function Get-PlaceholderText {
    param([string]$FieldPath)

    $leaf = ($FieldPath -split '/')[-1]
    if ($leaf.Length -eq 0) { return $leaf }
    return $leaf.Substring(0, 1).ToUpperInvariant() + $leaf.Substring(1)
}

function Get-SdtId { return (Get-Random -Minimum 1 -Maximum 2147483000) }

function Format-XmlText {
    param([string]$Text)
    return $Text.Replace('&', '&amp;').Replace('<', '&lt;').Replace('>', '&gt;')
}

function New-Run {
    param([string]$Text)
    return "<w:r><w:t xml:space=`"preserve`">$(Format-XmlText $Text)</w:t></w:r>"
}

function New-FieldSdt {
    param([string]$FieldPath)

    $xpath = ConvertTo-XPath $FieldPath
    $text = Format-XmlText (Get-PlaceholderText $FieldPath)
    return '<w:sdt><w:sdtPr><w:id w:val="' + (Get-SdtId) + '"/>' +
        '<w:placeholder><w:docPart w:val="DefaultPlaceholder_-1854013440"/></w:placeholder>' +
        '<w15:dataBinding w:prefixMappings="xmlns:ns0=''' + $script:CurrentNs + ''' " w:xpath="' + $xpath + '" w:storeItemID="' + $StoreItemId + '"/>' +
        '</w:sdtPr><w:sdtContent><w:r><w:t xml:space="preserve">' + $text + '</w:t></w:r></w:sdtContent></w:sdt>'
}

function New-Paragraph {
    param([string]$Text, [string]$Style)

    $inner = ''
    $cursor = 0
    foreach ($match in $FieldToken.Matches($Text)) {
        if ($match.Index -gt $cursor) {
            $inner += New-Run $Text.Substring($cursor, $match.Index - $cursor)
        }
        $inner += New-FieldSdt $match.Groups[1].Value
        $cursor = $match.Index + $match.Length
    }
    if ($cursor -lt $Text.Length) { $inner += New-Run $Text.Substring($cursor) }

    $properties = ''
    if ($Style) { $properties = "<w:pPr><w:pStyle w:val=`"$Style`"/></w:pPr>" }
    return "<w:p>$properties$inner</w:p>"
}

function New-RepeatingSectionSdt {
    param([string]$Path, [string]$ChildXml)

    $xpath = ConvertTo-XPath $Path -NoTrailingIndex
    return '<w:sdt><w:sdtPr><w:id w:val="' + (Get-SdtId) + '"/>' +
        '<w15:dataBinding w:prefixMappings="xmlns:ns0=''' + $script:CurrentNs + ''' " w:xpath="' + $xpath + '" w:storeItemID="' + $StoreItemId + '"/>' +
        '<w15:repeatingSection/></w:sdtPr><w:sdtContent>' +
        '<w:sdt><w:sdtPr><w:id w:val="' + (Get-SdtId) + '"/>' +
        '<w:placeholder><w:docPart w:val="DefaultPlaceholder_-1854013435"/></w:placeholder>' +
        '<w15:repeatingSectionItem/></w:sdtPr><w:sdtContent>' + $ChildXml +
        '</w:sdtContent></w:sdt></w:sdtContent></w:sdt>'
}

function ConvertTo-BlockXml {
    param([object[]]$Blocks)

    $xml = ''
    foreach ($block in $Blocks) {
        switch ($block.Kind) {
            'Line'   { $xml += New-Paragraph -Text $block.Text -Style $block.Style }
            'Repeat' { $xml += New-RepeatingSectionSdt -Path $block.Path -ChildXml (ConvertTo-BlockXml $block.Children) }
        }
    }
    return $xml
}

function Get-FieldPaths {
    param([object[]]$Blocks)

    $paths = [System.Collections.Generic.List[string]]::new()
    foreach ($block in $Blocks) {
        if ($block.Kind -eq 'Line') {
            foreach ($match in $FieldToken.Matches($block.Text)) { $paths.Add($match.Groups[1].Value) }
        }
        elseif ($block.Kind -eq 'Repeat') {
            $paths.Add($block.Path)
            foreach ($child in (Get-FieldPaths $block.Children)) { $paths.Add($child) }
        }
    }
    return $paths.ToArray()
}

# ── Package parts ──

$WordNamespaces = 'xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" ' +
    'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" ' +
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
    'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
    'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" ' +
    'xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" ' +
    'xmlns:w16sdtdh="http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash" ' +
    'mc:Ignorable="w14 w15 w16sdtdh"'

function New-CustomXmlData {
    param([string[]]$FieldPaths, [string]$Namespace)

    # Word resolves a binding by walking this document with the sdt's xpath, so the
    # element tree here has to mirror every path the document binds to.
    $root = [ordered]@{}
    foreach ($path in ($FieldPaths | Sort-Object -Unique)) {
        $node = $root
        foreach ($segment in ($path -split '/')) {
            if (-not $node.Contains($segment)) { $node[$segment] = [ordered]@{} }
            $node = $node[$segment]
        }
    }

    function Write-Nodes {
        param([System.Collections.Specialized.OrderedDictionary]$Node, [int]$Depth)

        $xml = ''
        foreach ($key in $Node.Keys) {
            $child = $Node[$key]
            $indent = "`n" + ('  ' * $Depth)
            # The root entity element resets to the empty namespace, the way Dynamics writes it.
            $nsAttribute = $(if ($Depth -eq 1) { ' xmlns=""' } else { '' })
            if (@($child.Keys).Count -gt 0) {
                $xml += "$indent<$key$nsAttribute>" + (Write-Nodes -Node $child -Depth ($Depth + 1)) + "$indent</$key>"
            }
            else {
                $xml += "$indent<$key$nsAttribute>$key</$key>"
            }
        }
        return $xml
    }

    return '<?xml version="1.0" encoding="utf-8"?><DocumentTemplate xmlns="' + $Namespace + '">' +
        (Write-Nodes -Node $root -Depth 1) + "`n</DocumentTemplate>"
}

$StylesXml = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles $WordNamespaces>
<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/><w:sz w:val="22"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="120"/></w:pPr></w:pPrDefault></w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>
<w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/><w:basedOn w:val="Normal"/><w:qFormat/><w:pPr><w:spacing w:after="240"/></w:pPr><w:rPr><w:b/><w:sz w:val="48"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:qFormat/><w:pPr><w:spacing w:before="240" w:after="120"/><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:b/><w:color w:val="2F5496"/><w:sz w:val="28"/></w:rPr></w:style>
</w:styles>
"@

$SettingsXml = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings $WordNamespaces><w:defaultTabStop w:val="720"/><w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>
"@

function New-Fixture {
    param([PSObject]$Fixture, [string]$Path)

    $isBound = [bool]$Fixture.RootEntity
    $script:CurrentNs = $(if ($isBound) { "urn:microsoft-crm/document-template/$($Fixture.RootEntity)/$($Fixture.TypeCode)/" } else { '' })

    $hasHeader = ($Fixture.PSObject.Properties.Name -contains 'Header')
    $hasFooter = ($Fixture.PSObject.Properties.Name -contains 'Footer')

    $allPaths = [System.Collections.Generic.List[string]]::new()
    foreach ($p in (Get-FieldPaths $Fixture.Body)) { $allPaths.Add($p) }
    if ($hasHeader) { foreach ($p in (Get-FieldPaths $Fixture.Header)) { $allPaths.Add($p) } }
    if ($hasFooter) { foreach ($p in (Get-FieldPaths $Fixture.Footer)) { $allPaths.Add($p) } }

    # -- word/document.xml --
    $sectionProperties = ''
    if ($hasHeader) { $sectionProperties += '<w:headerReference w:type="default" r:id="rIdHeader1"/>' }
    if ($hasFooter) { $sectionProperties += '<w:footerReference w:type="default" r:id="rIdFooter1"/>' }
    $sectionProperties += '<w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1418" w:right="1418" w:bottom="1418" w:left="1418" w:header="709" w:footer="709" w:gutter="0"/>'

    $documentXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        "<w:document $WordNamespaces><w:body>" + (ConvertTo-BlockXml $Fixture.Body) +
        "<w:sectPr>$sectionProperties</w:sectPr></w:body></w:document>"

    # -- relationships and content types --
    $documentRels = '<Relationship Id="rIdStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' +
        '<Relationship Id="rIdSettings" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>'
    $overrides = '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
        '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>' +
        '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>' +
        '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>' +
        '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>'

    $parts = [ordered]@{}
    $parts['word/document.xml'] = $documentXml
    $parts['word/styles.xml'] = $StylesXml
    $parts['word/settings.xml'] = $SettingsXml

    if ($isBound) {
        $documentRels += '<Relationship Id="rIdCustomXml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml" Target="../customXml/item1.xml"/>'
        $overrides += '<Override PartName="/customXml/itemProps1.xml" ContentType="application/vnd.openxmlformats-officedocument.customXmlProperties+xml"/>'

        $parts['customXml/item1.xml'] = New-CustomXmlData -FieldPaths $allPaths.ToArray() -Namespace $script:CurrentNs
        $parts['customXml/itemProps1.xml'] = '<?xml version="1.0" encoding="UTF-8" standalone="no"?>' +
            '<ds:datastoreItem ds:itemID="' + $StoreItemId + '" xmlns:ds="http://schemas.openxmlformats.org/officeDocument/2006/customXml">' +
            '<ds:schemaRefs><ds:schemaRef ds:uri="' + $script:CurrentNs + '"/><ds:schemaRef ds:uri=""/></ds:schemaRefs></ds:datastoreItem>'
        $parts['customXml/_rels/item1.xml.rels'] = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXmlProps" Target="itemProps1.xml"/></Relationships>'
    }

    if ($hasHeader) {
        $parts['word/header1.xml'] = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
            "<w:hdr $WordNamespaces>" + (ConvertTo-BlockXml $Fixture.Header) + '</w:hdr>'
        $documentRels += '<Relationship Id="rIdHeader1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>'
        $overrides += '<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>'
    }
    if ($hasFooter) {
        $parts['word/footer1.xml'] = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
            "<w:ftr $WordNamespaces>" + (ConvertTo-BlockXml $Fixture.Footer) + '</w:ftr>'
        $documentRels += '<Relationship Id="rIdFooter1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>'
        $overrides += '<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>'
    }

    $parts['word/_rels/document.xml.rels'] = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' + $documentRels + '</Relationships>'

    $parts['[Content_Types].xml'] = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
        '<Default Extension="xml" ContentType="application/xml"/>' + $overrides + '</Types>'

    $parts['_rels/.rels'] = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>' +
        '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>' +
        '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/></Relationships>'

    $parts['docProps/core.xml'] = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" ' +
        'xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">' +
        '<dc:title>' + (Format-XmlText $Fixture.Title) + '</dc:title>' +
        '<dc:description>' + (Format-XmlText $Fixture.Description) + '</dc:description>' +
        '<dc:creator>Document Template X-Ray fixtures</dc:creator></cp:coreProperties>'

    $parts['docProps/app.xml'] = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" ' +
        'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">' +
        '<Application>Microsoft Office Word</Application><AppVersion>16.0000</AppVersion></Properties>'

    # -- zip it --
    if (Test-Path $Path) { Remove-Item $Path -Force }
    $encoding = New-Object System.Text.UTF8Encoding($false)
    $archive = [System.IO.Compression.ZipFile]::Open($Path, [System.IO.Compression.ZipArchiveMode]::Create)
    try {
        foreach ($partName in $parts.Keys) {
            $entry = $archive.CreateEntry($partName, [System.IO.Compression.CompressionLevel]::Optimal)
            $stream = $entry.Open()
            try {
                $bytes = $encoding.GetBytes($parts[$partName])
                $stream.Write($bytes, 0, $bytes.Length)
            }
            finally { $stream.Dispose() }
        }
    }
    finally { $archive.Dispose() }

    return @($allPaths | Where-Object { $_ }).Count
}

# ── Main ──

if (-not (Test-Path $OutputDirectory)) { New-Item -ItemType Directory -Path $OutputDirectory -Force | Out-Null }

foreach ($fixture in $Fixtures) {
    if ($fixture.File -notlike $Name) { continue }

    $path = Join-Path $OutputDirectory $fixture.File
    $count = New-Fixture -Fixture $fixture -Path $path

    Write-Host ("  {0,-32} {1,3} binding(s)  " -f $fixture.File, $count) -NoNewline
    Write-Host $fixture.Description -ForegroundColor DarkGray
}
