param(
    [string]$MachineLearningTexDirectory = 'D:\WHU2025-2026\机器学习\机器学习章节总结LaTeX',
    [string]$MachineLearningPdfPath = 'D:\WHU2025-2026\机器学习\机器学习期末复习.pdf',
    [string]$CppPdfPath = 'D:\WHU2024-2025\C++\cpp期末.pdf'
)

$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent $PSScriptRoot
$notesRoot = Join-Path $repoRoot 'notes'
$machineLearningRoot = Join-Path $notesRoot 'machine-learning'
$cppRoot = Join-Path $notesRoot 'cpp'

function Escape-Html([string]$text) {
    return [System.Net.WebUtility]::HtmlEncode($text)
}

function Convert-LatexInline([string]$text) {
    if ($null -eq $text) { return '' }
    $normalized = $text.Trim().Replace('\bm{', '\mathbf{').Replace('``', '“').Replace("''", '”')
    return Escape-Html $normalized
}

function Convert-LatexDocument([string]$path) {
    $lines = @(Get-Content -LiteralPath $path -Encoding UTF8)
    $output = [System.Collections.Generic.List[string]]::new()
    $insideDocument = $false
    $listType = $null
    $insideDisplayMath = $false
    $mathLines = [System.Collections.Generic.List[string]]::new()

    foreach ($rawLine in $lines) {
        $line = $rawLine.Trim()
        if ($line -eq '\begin{document}') { $insideDocument = $true; continue }
        if (-not $insideDocument) { continue }
        if ($line -eq '\end{document}') { break }
        if ($line -eq '\maketitle' -or [string]::IsNullOrWhiteSpace($line)) { continue }

        if ($insideDisplayMath) {
            if ($line -eq '\]') {
                $formula = ($mathLines | ForEach-Object { Escape-Html ($_.Replace('\bm{', '\mathbf{')) }) -join "`n"
                $output.Add('<div class="formula-block">\[' + "`n" + $formula + "`n" + '\]</div>')
                $mathLines.Clear()
                $insideDisplayMath = $false
            } else {
                $mathLines.Add($rawLine.TrimEnd())
            }
            continue
        }

        if ($line -eq '\[') {
            $insideDisplayMath = $true
            continue
        }

        $sectionMatch = [regex]::Match($line, '^\\section\*\{(.+)\}$')
        if ($sectionMatch.Success) {
            $output.Add('<h2>' + (Convert-LatexInline $sectionMatch.Groups[1].Value) + '</h2>')
            continue
        }
        $subsectionMatch = [regex]::Match($line, '^\\subsection\*\{(.+)\}$')
        if ($subsectionMatch.Success) {
            $output.Add('<h3>' + (Convert-LatexInline $subsectionMatch.Groups[1].Value) + '</h3>')
            continue
        }

        $beginList = [regex]::Match($line, '^\\begin\{(itemize|enumerate|description)\}(?:\[[^]]+\])?$')
        if ($beginList.Success) {
            $listType = switch ($beginList.Groups[1].Value) {
                'itemize' { 'ul' }
                'enumerate' { 'ol' }
                'description' { 'dl' }
            }
            $output.Add("<$listType>")
            continue
        }
        $endList = [regex]::Match($line, '^\\end\{(itemize|enumerate|description)\}$')
        if ($endList.Success) {
            if ($null -ne $listType) { $output.Add("</$listType>") }
            $listType = $null
            continue
        }

        $itemMatch = [regex]::Match($line, '^\\item(?:\[([^]]+)\])?\s*(.*)$')
        if ($itemMatch.Success) {
            $label = Convert-LatexInline $itemMatch.Groups[1].Value
            $body = Convert-LatexInline $itemMatch.Groups[2].Value
            if ($listType -eq 'dl') {
                $output.Add('<dt>' + $label + '</dt><dd>' + $body + '</dd>')
            } else {
                $output.Add('<li>' + $body + '</li>')
            }
            continue
        }

        $output.Add('<p>' + (Convert-LatexInline $line) + '</p>')
    }

    if ($insideDisplayMath) { throw "Unclosed display-math block in $path" }
    if ($null -ne $listType) { throw "Unclosed list environment in $path" }
    return $output -join "`n"
}

function Join-WrappedText([System.Collections.Generic.List[string]]$parts) {
    if ($parts.Count -eq 0) { return '' }
    $joined = $parts[0]
    for ($index = 1; $index -lt $parts.Count; $index++) {
        $next = $parts[$index]
        $separator = if ($joined -match '[A-Za-z0-9]$' -and $next -match '^[A-Za-z0-9]') { ' ' } else { '' }
        $joined += $separator + $next
    }
    return $joined
}

function Convert-CppText([string[]]$lines) {
    $output = [System.Collections.Generic.List[string]]::new()
    $paragraph = [System.Collections.Generic.List[string]]::new()
    $code = [System.Collections.Generic.List[string]]::new()
    $listOpen = [ref]$false
    $codeOpen = [ref]$false

    function Flush-Paragraph {
        if ($paragraph.Count -gt 0) {
            $text = Join-WrappedText $paragraph
            $output.Add('<p>' + (Escape-Html $text) + '</p>')
            $paragraph.Clear()
        }
    }
    function Close-List {
        if ($listOpen.Value) {
            $output.Add('</ul>')
            $listOpen.Value = $false
        }
    }
    function Flush-Code {
        if ($code.Count -gt 0) {
            $output.Add('<pre><code>' + (Escape-Html ($code -join "`n")) + '</code></pre>')
            $code.Clear()
        }
        $codeOpen.Value = $false
    }

    foreach ($rawLine in $lines) {
        $line = $rawLine.TrimEnd()
        $trimmed = $line.Trim()
        if ($trimmed -eq '未完待续...') { continue }

        if ([string]::IsNullOrWhiteSpace($trimmed)) {
            if ($codeOpen.Value) { Flush-Code }
            Flush-Paragraph
            Close-List
            continue
        }

        $heading2 = [regex]::Match($trimmed, '^(\d+)\.\s*(.+)$')
        $heading3 = [regex]::Match($trimmed, '^[（(](\d+)[）)]\s*(.+)$')
        $heading4 = [regex]::Match($trimmed, '^([a-z])\.\s*(.+)$')
        if ($heading2.Success -or $heading3.Success -or $heading4.Success) {
            if ($codeOpen.Value) { Flush-Code }
            Flush-Paragraph
            Close-List
            if ($heading2.Success) {
                $output.Add('<h2>' + (Escape-Html ($heading2.Groups[1].Value + '. ' + $heading2.Groups[2].Value.Trim())) + '</h2>')
            } elseif ($heading3.Success) {
                $output.Add('<h3>' + (Escape-Html ('（' + $heading3.Groups[1].Value + '）' + $heading3.Groups[2].Value.Trim())) + '</h3>')
            } else {
                $output.Add('<h4>' + (Escape-Html ($heading4.Groups[1].Value + '. ' + $heading4.Groups[2].Value.Trim())) + '</h4>')
            }
            continue
        }

        $codeMatch = [regex]::Match($line, '^\s*\d+\s(.*)$')
        if ($codeMatch.Success) {
            Flush-Paragraph
            Close-List
            $codeOpen.Value = $true
            $code.Add($codeMatch.Groups[1].Value.TrimEnd())
            continue
        }
        if ($codeOpen.Value) {
            if ($trimmed -match '^\d+$') { $code.Add(''); continue }
            $code.Add($trimmed)
            continue
        }

        $bulletMatch = [regex]::Match($trimmed, '^(?:[•●▪◼]|—>|->)\s*(.+)$')
        if ($bulletMatch.Success) {
            Flush-Paragraph
            if (-not $listOpen.Value) { $output.Add('<ul>'); $listOpen.Value = $true }
            $output.Add('<li>' + (Escape-Html $bulletMatch.Groups[1].Value.Trim()) + '</li>')
            continue
        }

        $numberedPoint = [regex]::Match($trimmed, '^(\d+[）)])\s*(.+)$')
        if ($numberedPoint.Success) {
            Flush-Paragraph
            Close-List
            $output.Add('<p class="numbered-point"><strong>' + (Escape-Html $numberedPoint.Groups[1].Value) + '</strong> ' + (Escape-Html $numberedPoint.Groups[2].Value.Trim()) + '</p>')
            continue
        }

        if ($listOpen.Value) { Close-List }
        $paragraph.Add($trimmed)
    }
    if ($codeOpen.Value) { Flush-Code }
    Flush-Paragraph
    Close-List
    return $output -join "`n"
}

function New-ArticlePage {
    param(
        [string]$Course,
        [string]$CourseSlug,
        [int]$Number,
        [int]$Total,
        [string]$UnitLabel,
        [string]$Title,
        [string]$Summary,
        [string]$AcademicTerm,
        [string]$SourceLabel,
        [string]$Body,
        [bool]$UseMathJax
    )
    $previous = if ($Number -gt 1) { '<a href="chapter-' + ($Number - 1) + '.html">← 第 ' + ($Number - 1) + ' ' + $UnitLabel + '</a>' } else { '<span></span>' }
    $next = if ($Number -lt $Total) { '<a href="chapter-' + ($Number + 1) + '.html">第 ' + ($Number + 1) + ' ' + $UnitLabel + ' →</a>' } else { '<span></span>' }
    $mathScripts = if ($UseMathJax) {
@'
    <script>
        window.MathJax = { tex: { inlineMath: [['$', '$'], ['\\(', '\\)']] }, svg: { fontCache: 'global' } };
    </script>
    <script defer src="https://cdn.jsdelivr.net/npm/mathjax@3/es5/tex-svg.js"></script>
'@
    } else { '' }
    return @"
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <meta name="description" content="$(Escape-Html $Summary)">
    <title>$(Escape-Html "$Course · 第 $Number $UnitLabel：$Title") | 曲星宇</title>
    <link rel="stylesheet" href="../styles.css">
$mathScripts
</head>
<body>
    <header class="site-header">
        <a class="site-name" href="../../index.html">Xingyu Qu</a>
        <nav aria-label="笔记导航">
            <a href="../index.html">课程笔记</a>
            <a aria-current="page" href="index.html">本课程</a>
            <a href="../../index.html#service">返回主页</a>
        </nav>
    </header>
    <main class="article-shell">
        <article class="note-article">
            <header class="article-header">
                <p class="eyebrow">$(Escape-Html $Course) · 期末复习</p>
                <h1>$(Escape-Html $Title)</h1>
                <p class="article-summary">$(Escape-Html $Summary)</p>
                <div class="article-meta">
                    <span>$(Escape-Html "第 $Number $UnitLabel")</span>
                    <span>$(Escape-Html $AcademicTerm)</span>
                    <span>$(Escape-Html $SourceLabel)</span>
                </div>
            </header>
            <div class="article-body text-note-body">
$Body
            </div>
        </article>
        <nav class="article-pagination" aria-label="相邻章节">
            $previous
            <a href="index.html">系列目录</a>
            $next
        </nav>
    </main>
</body>
</html>
"@
}

function New-SeriesIndex {
    param(
        [string]$Course,
        [string]$Description,
        [string]$AcademicTerm,
        [string]$UnitLabel,
        [object[]]$Definitions
    )
    $cards = foreach ($definition in $Definitions) {
@"
            <article class="note-card">
                <div class="note-card-meta"><span>$(Escape-Html "第 $($definition.Number) $UnitLabel")</span><span>$(Escape-Html $AcademicTerm)</span></div>
                <h2><a href="chapter-$($definition.Number).html">$(Escape-Html $definition.Title)</a></h2>
                <p>$(Escape-Html $definition.Summary)</p>
                <a class="read-link" href="chapter-$($definition.Number).html">阅读笔记 <span aria-hidden="true">→</span></a>
            </article>
"@
    }
    return @"
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <meta name="description" content="$(Escape-Html $Description)">
    <title>$(Escape-Html $Course) | 课程笔记</title>
    <link rel="stylesheet" href="../styles.css">
</head>
<body>
    <header class="site-header">
        <a class="site-name" href="../../index.html">Xingyu Qu</a>
        <nav aria-label="笔记导航">
            <a href="../index.html">课程笔记</a>
            <a aria-current="page" href="index.html">本课程</a>
            <a href="../../index.html#service">返回主页</a>
        </nav>
    </header>
    <main class="notes-shell">
        <section class="series-intro">
            <p class="eyebrow">武汉大学计算机学院课程笔记</p>
            <h1>$(Escape-Html $Course)</h1>
            <p>$(Escape-Html $Description)</p>
            <div class="series-meta">
                <span>$($Definitions.Count) 篇笔记</span>
                <span>$(Escape-Html $AcademicTerm)</span>
                <span>文本整理版</span>
            </div>
        </section>
        <section class="note-grid" aria-label="本课程笔记">
$($cards -join "`n")
        </section>
    </main>
</body>
</html>
"@
}

$machineLearningDefinitions = @(
    [pscustomobject]@{ Number = 1; Title = '概论'; Summary = '梳理机器学习的核心概念、学习范式、建模流程、常见任务与易错点。' }
    [pscustomobject]@{ Number = 2; Title = '模型选择与评估'; Summary = '覆盖 VC 维、偏差与方差、交叉验证，以及分类、回归和不平衡数据的评价指标。' }
    [pscustomobject]@{ Number = 3; Title = '线性模型'; Summary = '整理线性回归、岭回归、对数几率回归、LDA 与 Softmax 多分类。' }
    [pscustomobject]@{ Number = 4; Title = '贝叶斯分类器'; Summary = '从贝叶斯决策与极大似然出发，介绍朴素贝叶斯、贝叶斯网和 EM 算法。' }
    [pscustomobject]@{ Number = 5; Title = '支持向量机'; Summary = '围绕间隔、对偶问题、KKT 条件、软间隔和核方法整理 SVM 的完整推导链。' }
    [pscustomobject]@{ Number = 6; Title = '决策树'; Summary = '比较信息增益、增益率与基尼指数，并整理连续属性、缺失值、剪枝和回归树。' }
    [pscustomobject]@{ Number = 7; Title = '集成学习'; Summary = '介绍 AdaBoost、梯度提升树、Bagging、随机森林与学习器多样性。' }
    [pscustomobject]@{ Number = 8; Title = '降维与度量学习'; Summary = '整理 k 近邻、维数灾难、MDS、PCA、核化降维、流形学习和度量学习。' }
    [pscustomobject]@{ Number = 9; Title = '聚类'; Summary = '覆盖聚类度量、距离计算、k-means、LVQ、GMM、DBSCAN 与层次聚类。' }
    [pscustomobject]@{ Number = 10; Title = '特征选择与稀疏学习'; Summary = '整理过滤式、包裹式和嵌入式选择，以及稀疏表示、压缩感知和矩阵补全。' }
    [pscustomobject]@{ Number = 11; Title = '半监督学习'; Summary = '介绍半监督假设、生成式方法、半监督 SVM、图方法、协同训练与半监督聚类。' }
    [pscustomobject]@{ Number = 12; Title = '强化学习'; Summary = '从 MDP 与价值函数出发，梳理策略迭代、价值迭代、Sarsa、Q-learning 与 RLHF。' }
)

$cppDefinitions = @(
    [pscustomobject]@{ Number = 1; Title = '类与对象初步'; Summary = '整理构造、复制、移动与析构函数，以及默认函数和删除函数。' }
    [pscustomobject]@{ Number = 2; Title = '继承与派生'; Summary = '覆盖基类与派生类、类型兼容、构造析构顺序和虚基类。' }
    [pscustomobject]@{ Number = 3; Title = '多态'; Summary = '梳理编译时与运行时多态、运算符重载、虚函数、虚析构和纯虚函数。' }
    [pscustomobject]@{ Number = 4; Title = '数据的共享和保护'; Summary = '介绍作用域、生存期、类的静态成员，以及共享数据的访问保护。' }
    [pscustomobject]@{ Number = 5; Title = '函数'; Summary = '集中整理函数原型说明与内联函数的基本规则。' }
    [pscustomobject]@{ Number = 6; Title = '数组、指针、字符串'; Summary = '覆盖数组、指针变量、指针运算、函数指针与 this 指针。' }
    [pscustomobject]@{ Number = 7; Title = 'STL'; Summary = '介绍标准模板库的基本组件、迭代器，以及顺序容器和关联容器接口。' }
    [pscustomobject]@{ Number = 8; Title = '函数与类的模板'; Summary = '整理函数模板、类模板与动态数组类模板示例。' }
    [pscustomobject]@{ Number = 9; Title = '流类库与输入输出'; Summary = '介绍标准 I/O 流、格式控制、文件读写和字符串流。' }
    [pscustomobject]@{ Number = 10; Title = '异常处理'; Summary = '梳理异常抛出与捕获、匹配规则、栈展开和异常再抛出。' }
    [pscustomobject]@{ Number = 11; Title = '其他细节知识点'; Summary = '汇总标识符、整型与实型常量、字符和字符串常量等语言细节。' }
)

if (-not (Test-Path -LiteralPath $MachineLearningTexDirectory -PathType Container)) { throw "Machine-learning LaTeX directory not found: $MachineLearningTexDirectory" }
if (-not (Test-Path -LiteralPath $MachineLearningPdfPath -PathType Leaf)) { throw "Machine-learning PDF not found: $MachineLearningPdfPath" }
if (-not (Test-Path -LiteralPath $CppPdfPath -PathType Leaf)) { throw "C++ PDF not found: $CppPdfPath" }

$texFiles = @(Get-ChildItem -LiteralPath $MachineLearningTexDirectory -Filter '*.tex' -File | Sort-Object Name)
if ($texFiles.Count -ne $machineLearningDefinitions.Count) { throw "Expected 12 machine-learning TeX files, found $($texFiles.Count)" }

New-Item -ItemType Directory -Force -Path $machineLearningRoot,$cppRoot | Out-Null

foreach ($definition in $machineLearningDefinitions) {
    $prefix = '{0:D2}_' -f $definition.Number
    $source = $texFiles | Where-Object { $_.Name.StartsWith($prefix) } | Select-Object -First 1
    if ($null -eq $source) { throw "Missing machine-learning source for chapter $($definition.Number)" }
    $body = Convert-LatexDocument $source.FullName
    if ($body -notmatch '<h2>知识框架</h2>') { throw "Knowledge framework missing from $($source.Name)" }
    $page = New-ArticlePage -Course '机器学习' -CourseSlug 'machine-learning' -Number $definition.Number -Total $machineLearningDefinitions.Count -UnitLabel '讲' -Title $definition.Title -Summary $definition.Summary -AcademicTerm '2026 春季' -SourceLabel 'LaTeX 文本整理版' -Body $body -UseMathJax $true
    Set-Content -LiteralPath (Join-Path $machineLearningRoot "chapter-$($definition.Number).html") -Value $page -Encoding UTF8
    Write-Output "Generated machine-learning chapter $($definition.Number): $($definition.Title)"
}

$machineLearningDescription = '机器学习期末复习共 12 讲，以 LaTeX 源码为正文依据，按概念、模型、算法与常见问法组织，并以网页文本和可缩放公式呈现。'
$machineLearningIndex = New-SeriesIndex -Course '机器学习' -Description $machineLearningDescription -AcademicTerm '2026 春季' -UnitLabel '讲' -Definitions $machineLearningDefinitions
Set-Content -LiteralPath (Join-Path $machineLearningRoot 'index.html') -Value $machineLearningIndex -Encoding UTF8

$pdfToText = Get-Command pdftotext.exe -ErrorAction SilentlyContinue
if ($null -eq $pdfToText) { $pdfToText = Get-Command pdftotext -ErrorAction Stop }
$temporaryText = [IO.Path]::GetTempFileName()
try {
    & $pdfToText.Source -layout -enc UTF-8 $CppPdfPath $temporaryText
    if ($LASTEXITCODE -ne 0) { throw "pdftotext failed with exit code $LASTEXITCODE" }
    $cppRaw = [IO.File]::ReadAllText($temporaryText, [Text.Encoding]::UTF8)
} finally {
    Remove-Item -LiteralPath $temporaryText -Force -ErrorAction SilentlyContinue
}
$cppRaw = $cppRaw.Replace([string][char]0x200B, '').Replace([string][char]0xFEFF, '').Replace([string][char]12, '')
$cppLines = @($cppRaw -split '\r?\n')
$chapterStarts = [System.Collections.Generic.List[object]]::new()
for ($index = 0; $index -lt $cppLines.Count; $index++) {
    $match = [regex]::Match($cppLines[$index], '^\s*（[一二三四五六七八九十]+）\s*(.+?)\s*$')
    if ($match.Success) { $chapterStarts.Add([pscustomobject]@{ Index = $index; Title = $match.Groups[1].Value.Trim() }) }
}
if ($chapterStarts.Count -ne $cppDefinitions.Count) { throw "Expected 11 C++ sections, found $($chapterStarts.Count)" }
for ($position = 0; $position -lt $chapterStarts.Count; $position++) {
    $definition = $cppDefinitions[$position]
    $normalizedSourceTitle = $chapterStarts[$position].Title.Replace('模版', '模板')
    if ($normalizedSourceTitle -ne $definition.Title) { throw "C++ section $($definition.Number) title mismatch: $($chapterStarts[$position].Title)" }
    $start = $chapterStarts[$position].Index + 1
    $end = if ($position -lt $chapterStarts.Count - 1) { $chapterStarts[$position + 1].Index - 1 } else { $cppLines.Count - 1 }
    $bodyLines = if ($end -ge $start) { @($cppLines[$start..$end]) } else { @() }
    $body = Convert-CppText $bodyLines
    $page = New-ArticlePage -Course 'C++ 程序设计' -CourseSlug 'cpp' -Number $definition.Number -Total $cppDefinitions.Count -UnitLabel '章' -Title $definition.Title -Summary $definition.Summary -AcademicTerm '2024 秋季' -SourceLabel 'PDF 文本整理版' -Body $body -UseMathJax $false
    Set-Content -LiteralPath (Join-Path $cppRoot "chapter-$($definition.Number).html") -Value $page -Encoding UTF8
    Write-Output "Generated C++ chapter $($definition.Number): $($definition.Title)"
}

$cppDescription = 'C++ 程序设计期末复习共 11 章，由原 PDF 提取为网页文本，按照面向对象、指针与容器、模板、输入输出和异常处理等主题重建层级。'
$cppIndex = New-SeriesIndex -Course 'C++ 程序设计' -Description $cppDescription -AcademicTerm '2024 秋季' -UnitLabel '章' -Definitions $cppDefinitions
Set-Content -LiteralPath (Join-Path $cppRoot 'index.html') -Value $cppIndex -Encoding UTF8

$mlPdfHash = (Get-FileHash -LiteralPath $MachineLearningPdfPath -Algorithm SHA256).Hash.ToLowerInvariant()
$cppPdfHash = (Get-FileHash -LiteralPath $CppPdfPath -Algorithm SHA256).Hash.ToLowerInvariant()
Write-Output 'Generated machine-learning and C++ series indexes.'
Write-Output "Machine-learning PDF SHA256: $mlPdfHash"
Write-Output "C++ PDF SHA256: $cppPdfHash"
