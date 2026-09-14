param(
    [string]$SourcePath = 'D:\Downloads\如何在大一开学计划自己的博士申请.docx'
)

$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent $PSScriptRoot
$outputRoot = Join-Path $repoRoot 'service\phd-application-guide'
$assetRoot = Join-Path $repoRoot 'assets\service\phd-application-guide'
$outputPath = Join-Path $outputRoot 'index.html'

if (-not (Test-Path -LiteralPath $SourcePath)) {
    throw "Source document not found: $SourcePath"
}

New-Item -ItemType Directory -Path $outputRoot -Force | Out-Null
New-Item -ItemType Directory -Path $assetRoot -Force | Out-Null

Add-Type -AssemblyName System.IO.Compression.FileSystem
$wordNamespace = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
$relationshipNamespace = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
$packageRelationshipNamespace = 'http://schemas.openxmlformats.org/package/2006/relationships'

function Read-ZipEntryText($archive, [string]$entryName) {
    $entry = $archive.GetEntry($entryName)
    if ($null -eq $entry) { throw "Missing DOCX entry: $entryName" }
    $reader = [System.IO.StreamReader]::new($entry.Open())
    try { return $reader.ReadToEnd() } finally { $reader.Dispose() }
}

function Get-WordAttribute($node, [string]$name) {
    if ($null -eq $node) { return '' }
    return $node.GetAttribute($name, $wordNamespace)
}

function Test-WordToggle($node) {
    if ($null -eq $node) { return $false }
    $value = Get-WordAttribute $node 'val'
    return $value -notin @('false', '0', 'off', 'none')
}

function Convert-TwipsToPoints([string]$value) {
    if ([string]::IsNullOrWhiteSpace($value)) { return $null }
    return ([double]$value / 20).ToString('0.##', [Globalization.CultureInfo]::InvariantCulture) + 'pt'
}

function Convert-RunToHtml($run, $namespaceManager) {
    $textNodes = @($run.SelectNodes('.//w:t', $namespaceManager))
    $text = ($textNodes | ForEach-Object { $_.InnerText }) -join ''
    $breakCount = @($run.SelectNodes('.//w:br', $namespaceManager)).Count
    $tabCount = @($run.SelectNodes('.//w:tab', $namespaceManager)).Count
    if ($breakCount -gt 0) { $text += '<br>' * $breakCount }
    if ($tabCount -gt 0) { $text += '    ' * $tabCount }
    $text = $text.Replace('坚守一些压力和麻烦', '减少一些压力和麻烦')
    if ([string]::IsNullOrEmpty($text)) { return '' }

    $runProperties = $run.SelectSingleNode('./w:rPr', $namespaceManager)
    $fonts = $runProperties.SelectSingleNode('./w:rFonts', $namespaceManager)
    $asciiFont = Get-WordAttribute $fonts 'ascii'
    $eastAsiaFont = Get-WordAttribute $fonts 'eastAsia'
    $sizeValue = Get-WordAttribute ($runProperties.SelectSingleNode('./w:sz', $namespaceManager)) 'val'
    $colorValue = Get-WordAttribute ($runProperties.SelectSingleNode('./w:color', $namespaceManager)) 'val'
    $underlineValue = Get-WordAttribute ($runProperties.SelectSingleNode('./w:u', $namespaceManager)) 'val'

    $styles = [System.Collections.Generic.List[string]]::new()
    if (-not [string]::IsNullOrWhiteSpace($asciiFont) -or -not [string]::IsNullOrWhiteSpace($eastAsiaFont)) {
        $latin = if ([string]::IsNullOrWhiteSpace($asciiFont)) { 'Arial' } else { $asciiFont }
        $cjk = if ([string]::IsNullOrWhiteSpace($eastAsiaFont)) { 'DengXian' } else { $eastAsiaFont }
        $fontFamilies = [System.Collections.Generic.List[string]]::new()
        foreach ($fontName in @($latin, $cjk, 'DengXian', '等线')) {
            if (-not [string]::IsNullOrWhiteSpace($fontName) -and -not $fontFamilies.Contains($fontName)) {
                $fontFamilies.Add($fontName)
            }
        }
        $styles.Add('font-family: ' + (($fontFamilies | ForEach-Object { "'$_'" }) -join ', ') + ', sans-serif')
    }
    if (-not [string]::IsNullOrWhiteSpace($sizeValue)) {
        $styles.Add((([double]$sizeValue / 2).ToString('0.##', [Globalization.CultureInfo]::InvariantCulture)) + 'pt')
        $styles[$styles.Count - 1] = 'font-size: ' + $styles[$styles.Count - 1]
    }
    if (Test-WordToggle ($runProperties.SelectSingleNode('./w:b', $namespaceManager))) { $styles.Add('font-weight: 700') }
    if (Test-WordToggle ($runProperties.SelectSingleNode('./w:i', $namespaceManager))) { $styles.Add('font-style: italic') }
    if (-not [string]::IsNullOrWhiteSpace($colorValue) -and $colorValue -ne 'auto') { $styles.Add("color: #$colorValue") }

    $decorations = [System.Collections.Generic.List[string]]::new()
    if (Test-WordToggle ($runProperties.SelectSingleNode('./w:strike', $namespaceManager))) { $decorations.Add('line-through') }
    if (-not [string]::IsNullOrWhiteSpace($underlineValue) -and $underlineValue -ne 'none') { $decorations.Add('underline') }
    if ($decorations.Count -gt 0) { $styles.Add('text-decoration: ' + ($decorations -join ' ')) }

    $encoded = [System.Net.WebUtility]::HtmlEncode($text).Replace('&lt;br&gt;', '<br>')
    if ($text -match '^https?://[^\s]+$') {
        $encoded = '<a class="word-url" href="' + ([System.Net.WebUtility]::HtmlEncode($text)) + '">' + $encoded + '</a>'
    }
    if ($styles.Count -eq 0) { return $encoded }
    return '<span style="' + ($styles -join '; ') + '">' + $encoded + '</span>'
}

function Convert-ParagraphRunsToHtml($paragraph, $namespaceManager) {
    return (@($paragraph.SelectNodes('./w:r', $namespaceManager)) | ForEach-Object {
        Convert-RunToHtml $_ $namespaceManager
    }) -join ''
}

function Get-ParagraphStyle($paragraph, $namespaceManager) {
    $properties = $paragraph.SelectSingleNode('./w:pPr', $namespaceManager)
    $spacing = $properties.SelectSingleNode('./w:spacing', $namespaceManager)
    $indent = $properties.SelectSingleNode('./w:ind', $namespaceManager)
    $alignment = Get-WordAttribute ($properties.SelectSingleNode('./w:jc', $namespaceManager)) 'val'
    $styles = [System.Collections.Generic.List[string]]::new()

    $before = Convert-TwipsToPoints (Get-WordAttribute $spacing 'before')
    $after = Convert-TwipsToPoints (Get-WordAttribute $spacing 'after')
    $lineValue = Get-WordAttribute $spacing 'line'
    $lineRule = Get-WordAttribute $spacing 'lineRule'
    $line = $null
    if (-not [string]::IsNullOrWhiteSpace($lineValue)) {
        if ([string]::IsNullOrWhiteSpace($lineRule) -or $lineRule -eq 'auto') {
            $line = ([double]$lineValue / 240).ToString('0.###', [Globalization.CultureInfo]::InvariantCulture)
        }
        else {
            $line = Convert-TwipsToPoints $lineValue
        }
    }
    $left = Convert-TwipsToPoints (Get-WordAttribute $indent 'left')
    $firstLine = Convert-TwipsToPoints (Get-WordAttribute $indent 'firstLine')
    if ($null -ne $before) { $styles.Add("margin-top: $before") }
    if ($null -ne $after) { $styles.Add("margin-bottom: $after") }
    if ($null -ne $line) { $styles.Add("line-height: $line") }
    if ($null -ne $left -and $left -ne '0pt') { $styles.Add("margin-left: $left") }
    if ($null -ne $firstLine -and $firstLine -ne '0pt') { $styles.Add("text-indent: $firstLine") }
    if (-not [string]::IsNullOrWhiteSpace($alignment)) { $styles.Add("text-align: $alignment") }
    return $styles -join '; '
}

$archive = [System.IO.Compression.ZipFile]::OpenRead($SourcePath)
try {
    [xml]$documentXml = Read-ZipEntryText $archive 'word/document.xml'
    [xml]$relationshipsXml = Read-ZipEntryText $archive 'word/_rels/document.xml.rels'
    [xml]$numberingXml = Read-ZipEntryText $archive 'word/numbering.xml'

    $namespaceManager = [System.Xml.XmlNamespaceManager]::new($documentXml.NameTable)
    $namespaceManager.AddNamespace('w', $wordNamespace)
    $namespaceManager.AddNamespace('a', 'http://schemas.openxmlformats.org/drawingml/2006/main')
    $namespaceManager.AddNamespace('r', $relationshipNamespace)
    $namespaceManager.AddNamespace('wp', 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing')

    $relationshipMap = @{}
    foreach ($relationship in $relationshipsXml.Relationships.Relationship) {
        $relationshipMap[$relationship.Id] = $relationship.Target
    }

    $body = $documentXml.SelectSingleNode('//w:body', $namespaceManager)
    $paragraphCount = @($body.SelectNodes('./w:p', $namespaceManager)).Count
    $tableCount = @($body.SelectNodes('./w:tbl', $namespaceManager)).Count
    $drawingCount = @($body.SelectNodes('.//w:drawing', $namespaceManager)).Count
    $trackedChangeCount = @($body.SelectNodes('.//w:ins | .//w:del', $namespaceManager)).Count
    $titleText = (@($body.SelectSingleNode('./w:p[1]', $namespaceManager).SelectNodes('.//w:t', $namespaceManager)) | ForEach-Object { $_.InnerText }) -join ''

    if ($titleText -ne '如何在大一开学计划自己的博士申请') { throw "Unexpected document title: $titleText" }
    if ($paragraphCount -ne 44 -or $tableCount -ne 2 -or $drawingCount -ne 2) {
        throw "Unexpected source structure: paragraphs=$paragraphCount tables=$tableCount drawings=$drawingCount"
    }
    if ($trackedChangeCount -ne 0) { throw "Tracked changes are not supported by this importer: $trackedChangeCount found" }

    $sourceHash = (Get-FileHash -LiteralPath $SourcePath -Algorithm SHA256).Hash.ToLowerInvariant()
    $blocks = [System.Collections.Generic.List[string]]::new()
    $paragraphIndex = 0
    $imageIndex = 0
    $listOpen = $false

    foreach ($child in $body.ChildNodes) {
        if ($child.LocalName -eq 'sectPr') { continue }

        if ($child.LocalName -eq 'p') {
            $paragraphIndex++
            $paragraph = $child
            $numberingId = Get-WordAttribute ($paragraph.SelectSingleNode('./w:pPr/w:numPr/w:numId', $namespaceManager)) 'val'
            $paragraphStyleId = Get-WordAttribute ($paragraph.SelectSingleNode('./w:pPr/w:pStyle', $namespaceManager)) 'val'
            $drawing = $paragraph.SelectSingleNode('.//w:drawing', $namespaceManager)
            $paragraphHtml = Convert-ParagraphRunsToHtml $paragraph $namespaceManager
            $paragraphStyle = Get-ParagraphStyle $paragraph $namespaceManager

            if (-not [string]::IsNullOrWhiteSpace($numberingId)) {
                if (-not $listOpen) { $blocks.Add('<ul class="word-bullet-list">'); $listOpen = $true }
                $blocks.Add('<li class="word-paragraph" style="' + $paragraphStyle + '">' + $paragraphHtml + '</li>')
                continue
            }
            if ($listOpen) { $blocks.Add('</ul>'); $listOpen = $false }

            if ($null -ne $drawing) {
                $imageIndex++
                $blip = $drawing.SelectSingleNode('.//a:blip', $namespaceManager)
                $relationId = $blip.GetAttribute('embed', $relationshipNamespace)
                if (-not $relationshipMap.ContainsKey($relationId)) { throw "Image relationship not found: $relationId" }
                $target = 'word/' + $relationshipMap[$relationId].Replace('\', '/')
                $entry = $archive.GetEntry($target)
                if ($null -eq $entry) { throw "Image entry not found: $target" }
                $extension = [System.IO.Path]::GetExtension($entry.Name).ToLowerInvariant()
                $assetName = "figure-$imageIndex$extension"
                $assetPath = Join-Path $assetRoot $assetName
                $inputStream = $entry.Open()
                $outputStream = [System.IO.File]::Create($assetPath)
                try { $inputStream.CopyTo($outputStream) } finally { $outputStream.Dispose(); $inputStream.Dispose() }
                $altText = if ($imageIndex -eq 1) { 'PhD 申请筛选研究潜力的路径图' } else { '科研成长的正反馈循环图' }
                $blocks.Add('<figure class="word-figure" style="' + $paragraphStyle + '"><img src="../../assets/service/phd-application-guide/' + $assetName + '" alt="' + $altText + '"></figure>')
                continue
            }

            if ([string]::IsNullOrWhiteSpace(([System.Net.WebUtility]::HtmlDecode(($paragraphHtml -replace '<[^>]+>', ''))))) { continue }
            if ($paragraphIndex -eq 1) {
                $blocks.Add('<h1 class="word-title" style="' + $paragraphStyle + '">' + $paragraphHtml + '</h1>')
            } elseif ($paragraphIndex -eq 2) {
                $blocks.Add('<p class="word-meta word-paragraph" style="' + $paragraphStyle + '">' + $paragraphHtml + '</p>')
            } elseif ($paragraphStyleId -eq '1') {
                $blocks.Add('<h2 class="word-heading" style="' + $paragraphStyle + '">' + $paragraphHtml + '</h2>')
            } else {
                $blocks.Add('<p class="word-paragraph" style="' + $paragraphStyle + '">' + $paragraphHtml + '</p>')
            }
            continue
        }

        if ($child.LocalName -eq 'tbl') {
            if ($listOpen) { $blocks.Add('</ul>'); $listOpen = $false }
            $cell = $child.SelectSingleNode('./w:tr[1]/w:tc[1]', $namespaceManager)
            $cellParagraph = $cell.SelectSingleNode('./w:p[1]', $namespaceManager)
            $cellHtml = Convert-ParagraphRunsToHtml $cellParagraph $namespaceManager
            $cellParagraphStyle = Get-ParagraphStyle $cellParagraph $namespaceManager
            $blocks.Add('<div class="word-callout"><p class="word-paragraph" style="' + $cellParagraphStyle + '">' + $cellHtml + '</p></div>')
        }
    }
    if ($listOpen) { $blocks.Add('</ul>') }

    if ($imageIndex -ne 2) { throw "Unexpected extracted image count: $imageIndex" }

    $html = @"
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <meta name="description" content="面向本科新生的博士申请与科研成长规划，讨论研究潜力、精力分配、科研积累、推荐信与长期发展。">
    <meta name="source-sha256" content="$sourceHash">
    <title>如何在大一开学计划自己的博士申请 | 曲星宇</title>
    <link rel="stylesheet" href="styles.css">
</head>
<body>
    <header class="site-header">
        <a class="site-name" href="../../index.html">Xingyu Qu</a>
        <nav aria-label="页面导航">
            <a href="../../index.html#service">Service</a>
            <a href="../../notes/index.html">课程笔记</a>
            <a href="../../index.html">返回主页</a>
        </nav>
    </header>
    <main>
        <article class="document-sheet" aria-labelledby="document-title">
            <div class="document-content">
$($blocks -join "`n")
            </div>
        </article>
    </main>
</body>
</html>
"@
    $html = $html.Replace('class="word-title"', 'id="document-title" class="word-title"')
    [System.IO.File]::WriteAllText($outputPath, $html, [System.Text.UTF8Encoding]::new($false))

    Write-Output "Generated $outputPath"
    Write-Output "Source SHA256: $sourceHash"
    Write-Output "Preserved structure: 44 paragraphs, 6 headings, 2 callouts, 2 images"
} finally {
    $archive.Dispose()
}
