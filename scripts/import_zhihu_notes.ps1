param(
    [string]$ColumnId = 'c_1949874045978936715'
)

$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'

$repoRoot = Split-Path -Parent $PSScriptRoot
$notesRoot = Join-Path $repoRoot 'notes'
$seriesRoot = Join-Path $notesRoot 'computer-organization'
$assetRoot = Join-Path $repoRoot 'assets\notes'
$apiUrl = "https://www.zhihu.com/api/v4/columns/$ColumnId/articles?limit=20&offset=0"

function Escape-Html([string]$text) {
    if ($null -eq $text) { return '' }
    return [System.Net.WebUtility]::HtmlEncode($text)
}

function Get-PlainText([string]$html) {
    $text = [regex]::Replace($html, '<[^>]+>', ' ')
    $text = [System.Net.WebUtility]::HtmlDecode($text)
    return ([regex]::Replace($text, '\s+', ' ')).Trim()
}

function Remove-ParagraphContaining([string]$html, [string]$needle) {
    $pattern = '<p\b[^>]*>(?:(?!</p>)[\s\S])*?' + [regex]::Escape($needle) + '(?:(?!</p>)[\s\S])*?</p>'
    return [regex]::Replace($html, $pattern, '', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
}

function Invoke-CurlText([string]$url) {
    $parts = & curl.exe -sS -L --max-time 45 -A 'Mozilla/5.0' -H 'Referer: https://www.zhihu.com/' $url
    if ($LASTEXITCODE -ne 0) { throw "curl failed for $url" }
    return ($parts -join "`n")
}

function Download-Image([string]$url, [string]$destination) {
    $parent = Split-Path -Parent $destination
    New-Item -ItemType Directory -Force -Path $parent | Out-Null
    if ((Test-Path -LiteralPath $destination) -and (Get-Item -LiteralPath $destination).Length -ge 512) {
        return
    }
    & curl.exe -sS -L --max-time 45 -A 'Mozilla/5.0' -H 'Referer: https://zhuanlan.zhihu.com/' -o $destination $url
    if ($LASTEXITCODE -ne 0 -or -not (Test-Path -LiteralPath $destination) -or (Get-Item -LiteralPath $destination).Length -lt 512) {
        throw "image download failed for $url"
    }
}

$summaries = @{
    0  = '介绍计算机体系结构的八个核心设计思想，并说明执行时间、CPU 时间、CPI、时钟频率等性能指标；随后讨论功耗墙、多核处理器与 SPEC 基准测试。'
    1  = '从软件与硬件接口的角度介绍 RISC-V 指令集，梳理算术、数据传送与控制流指令，以及寄存器、内存、立即数和主要指令格式。'
    2  = '逐一讲解 RISC-V 的 R、I、S、B、U、J 型指令，说明立即数编码、访存、分支与跳转语义，并总结字段布局背后的设计原则。'
    3  = '以递归函数为例说明 RISC-V 过程调用、寄存器保存和栈帧变化，同时介绍常见寻址模式，以及使用 LR/SC 实现同步的原子操作。'
    4  = '从一位逻辑与加法单元扩展到 32 位 ALU，分析减法、比较和溢出判断，并推导分层超前进位加法器及其 Verilog 实现。'
    5  = '梳理乘法器与除法器的数据通路、迭代过程和 RISC-V 指令，并介绍 IEEE 754 浮点数表示，以及浮点加法和乘法的执行步骤。'
    6  = '介绍处理器的基本组成和指令执行阶段，分别构建 R 型、load/store 与分支指令的数据通路，最后组合成单周期 RISC-V 完整数据通路。'
    7  = '围绕单周期处理器的控制通路，说明不同指令所需的控制信号、ALU 控制器、组合逻辑与 ROM 控制方案，并分析指令延迟和关键路径。'
    8  = '从多周期处理器过渡到五级流水线，讲解流水线寄存器和数据通路的组织方式，并区分结构冒险、数据冒险与控制冒险。'
    9  = '完善流水线控制通路，介绍前递、暂停与冒险检测机制；进一步讨论一位和两位分支预测器、分支目标缓冲，以及流水线中的异常处理。'
    10 = '从局部性原理出发建立存储层次，比较 SRAM、DRAM、闪存与磁盘的特点，并讲解直接映射 Cache 的地址划分、命中判断和数据访问过程。'
    11 = '讨论 Cache 的写直达与写回策略、AMAT 和相联度，扩展到多级 Cache；随后介绍虚拟内存、页表、缺页处理和 TLB 地址转换。'
    12 = '系统梳理无条件传送、查询、中断和 DMA 四种 I/O 方式，并介绍总线分类、性能、仲裁、定时与系统结构，以及常用接口和外部设备。'
}

$raw = Invoke-CurlText $apiUrl
$payload = $raw | ConvertFrom-Json
$articles = @($payload.data | Sort-Object created)
if ($articles.Count -ne 13) { throw "Expected 13 articles, found $($articles.Count)" }

$byId = @{}
$byLecture = @{}
foreach ($article in $articles) {
    if ($article.title -notmatch 'Lecture(\d+)$') { throw "Unexpected title: $($article.title)" }
    $lecture = [int]$Matches[1]
    $byId[[string]$article.id] = $lecture
    $byLecture[$lecture] = $article
}
if (@($byLecture.Keys | Sort-Object) -join ',' -ne '0,1,2,3,4,5,6,7,8,9,10,11,12') {
    throw 'Lecture sequence is incomplete.'
}

New-Item -ItemType Directory -Force -Path $notesRoot | Out-Null
New-Item -ItemType Directory -Force -Path $seriesRoot | Out-Null
New-Item -ItemType Directory -Force -Path $assetRoot | Out-Null

$commonReplacements = [ordered]@{
    'risc-v' = 'RISC-V'
    'RISCV' = 'RISC-V'
    'github' = 'GitHub'
    '上节课我们一起学习了' = '上一讲介绍了'
    '上一节课我们一起学习了' = '上一讲介绍了'
    '上节课我们学习了' = '上一讲介绍了'
    '上一节课我们学习了' = '上一讲介绍了'
    '这节课我们继续一起学习' = '本讲继续介绍'
    '这节课我们一起学习了' = '本讲介绍了'
    '这节课我们一起学习' = '本讲介绍'
    '这节课我们学习' = '本讲介绍'
    '下节课将一起学习' = '下一讲介绍'
    '我们可以发现' = '可以看到'
    '我们可以看到' = '可以看到'
    '我们可以得到' = '可以得到'
    '我们可以设计' = '可以设计'
    '我们需要' = '需要'
    '我们要' = '需要'
    '我们将' = '本讲将'
    '笔者认为' = '从内容组织上看，'
    '恳请各位指正' = '供参考与校核'
    '失去了失去了' = '失去了'
    '不在来自于' = '不再来自于'
    '我们之关心' = '这里只关心'
    '~！' = '。'
    '！' = '。'
}

$specificReplacements = @{
    0 = [ordered]@{
        '坚持更新，努力完成课程要求。' = '本讲概述计算机体系结构中的核心设计思想，并介绍性能、功耗与成本的基本度量方法。'
    }
    1 = [ordered]@{
        '这一讲主要是介绍RISC-V的特性和整体的指令架构，还没有过多关注细节。' = '本讲概述 RISC-V 的主要特性与整体指令架构，具体编码细节将在下一讲展开。'
        '（1）算术指令——计算机必须能计算啊。' = '（1）算术指令：完成基本算术与逻辑运算。'
        '正如教授所言，无论之后研究计算机的什么方向，这些经典又生动的思想会伴随我们始终。' = '这些设计原则也适用于后续的体系结构与系统研究。'
    }
    2 = [ordered]@{
        '这节课学习RISC-V整体指令架构的细节，从内容组织上看，学习指令主要是两部分：一部分是指令本身的结构，这与后面纯硬件部分的设计息息相关；另一部分是一些设计中的“匠心”，这些设计理念给人启迪，帮助我们发现计算机设计的优美，甚至给我们idea上的启示。但是交叉介绍两部分会破坏内容的连贯性，所以我们在最后统一阐释我第二部分的想法。' = '本讲首先介绍 RISC-V 指令结构及其硬件含义，再集中讨论编码规整性、扩展性等设计原则。'
        '如果你对内容不感兴趣，希望这张表也能帮到你。' = '下表汇总了本讲涉及的 RISC-V 指令类型。'
        '这种别扭的指令格式本讲将在最后讨论。' = '这种非连续的立即数字段布局将在本讲末尾讨论。'
        '对于过程调用和其他细节的详细介绍在下一讲中会详细说明。' = '过程调用及相关细节将在下一讲介绍。'
        '这也回应了前文本讲埋下的伏笔。' = '该布局与 S 型指令保持较高的一致性。'
    }
    3 = [ordered]@{
        '这一届我们学习了RISC-V过程调用，寻址模式，同步机制。由于本文已经学习了《深入理解计算机系统基础》，对于原书介绍的系统方面的知识没有过多介绍，感兴趣的读者可以自行查阅。' = '本讲总结 RISC-V 过程调用、寻址模式与同步机制；更完整的系统背景可结合体系结构教材与系统教材阅读。'
        '此外，这节课带我第一次认识到在汇编层面的锁机制，原文中的挖掘还不够细致和深刻，后期我可能会在评论中补充。' = '汇编层面的原子指令是构建锁等同步原语的基础，实际实现还需结合内存模型与并发语义。'
        'if (n < 1) return f;' = 'if (n < 1) return 1;'
    }
    4 = [ordered]@{
        '<h3>too slow</h3>' = '<h3>串行进位的延迟</h3>'
        '简单将 32 个 1 位 ALU 串联起来能不能得到 32 位 ALU 呢？还需要考虑一些重要问题来修改本讲设计的 ALU。' = '串联 32 个 1 位 ALU 可以形成基本的 32 位 ALU，但还需要加入溢出检测与比较逻辑。'
        '这个结论非常好，因为 carryout（结果进位标志）恰好是本讲设计中的一部分，可以直接利用制作出 overflow' = 'carryout（结果进位标志）已经是该设计的一部分，因此可以直接用于构造 overflow'
        '上面的 ALU 实现时，每一个 1 位 ALU 计算都需要依赖于前一位结果的进位情况，这种无法用于时序要求严格的硬件，所以需要快速进位以实现快速加法。' = '在上述 ALU 中，每一位都依赖前一位的进位，关键路径较长，因此需要使用快速进位结构。'
        '但是这样会导致硬件过于复杂，本讲无法承受这种开销。' = '直接展开到 32 位会显著增加硬件复杂度。'
        'c3=g2+p1·c2+p2·p1·c1+p2·p1·p0·c0' = 'c3=g2+p2·g1+p2·p1·g0+p2·p1·p0·c0'
        '共迎国庆，欢度中秋，祝大家双节快乐。' = ''
    }
    5 = [ordered]@{
        '并行进位加法器、浮点数加法器、乘法器的实现都将上传到GitHub的仓库中，供参考与校核。' = '并行进位加法器、浮点加法器与乘法器的参考实现见配套 GitHub 仓库。'
        '由于数学部分过于复杂，在后面实现除法器的时候我再考虑补充这部分内容。' = 'SRT 除法涉及更复杂的商选择与纠错机制，本文不展开。'
        '本讲介绍了乘法器、除法器的设计和浮点计算的加法和乘法设计。本讲经常在一起讨论，这些多样复杂的硬件图如何能够印在本讲的脑海中？我的回答是自己实现一遍。通过编写verilog代码，本讲不但能够理解这些硬件的功能和结构，甚至可以最大程度了解细节，这也是我觉得在vibe coding时代下比较有趣有价值的编码体验。后面我会尽量多实现硬件，上传到仓库中。' = '本讲介绍乘法器、除法器以及浮点加法与乘法。通过 Verilog 实现这些数据通路，有助于验证结构理解并进一步掌握控制细节。'
    }
    7 = [ordered]@{
        '依稀记得在前面的Lecture中，本讲在指令架构的学习中，就有操作码（Opcode）的字段，这些操作码在这里就成为了处理器的分流依据。' = '在前述指令架构中，操作码（Opcode）字段用于区分指令类型；在处理器中，它进一步成为控制信号生成的主要输入。'
        '当然，这还需要您翻看之前的“指令实现的步骤”。' = '具体路径可结合前述指令执行步骤分析。'
        '三日三更实属不易，希望锅炉的熟悉人和陌生人能一键三连。' = ''
    }
    8 = [ordered]@{
        '细忖度之，可以看到，add指令在 EX阶段结束之后 x19就不会再动了，可以直接被后面的指令使用，不需要在寄存器中再存取一次了。' = '进一步观察可知，add 指令在 EX 阶段结束后已经产生 x19 的新值，可以直接转发给后续指令，无需等待寄存器写回。'
        '当然，现代的流水线采用预测的方式解决beq，最新的预测技术成功率可达98%。' = '现代处理器通常使用分支预测来降低控制冒险开销。'
    }
    9 = [ordered]@{
        '问题来了，本讲如何实现插入bubble的操作？' = '插入 bubble 需要同时控制流水线后段和前段的状态。'
        '相关分支预测器的核心想法很简单：' = '相关分支预测器的核心思想是：'
        '它会维护一个短短的历史比特串' = '它维护一段有限长度的历史比特串'
        '锦标赛预测器的思路是：没有一种预测器能在所有场景里表现最好，那就干脆让几种预测器一起上，谁更靠谱就听谁的。' = '锦标赛预测器并行使用多种预测器，并通过选择器动态采用近期更准确的结果。'
        'RISC-V体系结构中使用000000001C090000作为例外入口地址。' = '具体的异常入口地址由实现与运行环境决定。'
        '对于这些部分的内容，本文打算在寒假学习，届时可能会继续更新本专栏的内容，供参考与校核各位大佬指正。' = ''
    }
    11 = [ordered]@{
        '本讲介绍完成了存储器的学习，对于重要的几张大图，他们是理解Memory Architecture的关键。对于一些其他内容（例如汉明编码、虚拟存储中的保护），这里由于笔力有限未做介绍，在寒假可能会补充一部分这些内容。这篇笔记内容较多，如有疏漏错误之处敬请批评指正。' = '本讲以缓存与虚拟存储的关键结构收束存储器部分。汉明编码与虚拟存储保护等主题不在本文范围内。'
    }
    12 = [ordered]@{
        'CPU 和设备“想发就发，想收就收”，不检查准备情况；' = 'CPU 与设备直接收发数据，不检查对方是否已经准备就绪；'
        '实际上是“盲发/盲收”。' = '这种方式不包含显式握手机制。'
        '很容易丢数据或读到垃圾数据。' = '容易丢失数据或读取无效数据。'
        '计组的学习告一段落，期末复习如期而至。回看学习的历程和笔记，记录的点滴对我理解整个课程或者说计算机的Architecture有了很大帮助，也为我复习做题计算提供了理论支持。感谢各位的对12+1篇文章的关注和支持，希望我之后会传递更加优质的内容和教学资源，帮助更多同学学习课程。' = '本讲完成 I/O 与总线部分，也为本系列的 13 篇课程笔记收尾。'
    }
}

# Corrections for the first three notes were reviewed against their complete
# public article bodies and all twenty original figures, not only excerpts.
$firstThreeCorrections = @{
    0 = [ordered]@{
        '坚持更新，努力完成课程要求！' = '本讲概述计算机体系结构中的核心设计思想，并介绍性能、功耗与成本的基本度量方法。'
        '课程采用risc-v架构教学，注重实验（动手能力），最高目标是设计出一台微型计算机（与后面的OS等课程组合）。' = '本系列以 RISC-V 为基础，结合数据通路与控制逻辑实验，逐步建立从指令集到微型计算机系统的整体认识。'
        '本文作为预讲课的note，总体上介绍一些认识。' = '本文是系列导论，集中介绍后续内容所需的基本概念。'
        '实验和测量得出计算机的经常性事件，加速之以提升性能。越common越easy。' = '通过实验与测量识别常见路径，优先优化高频事件，以获得更显著的整体性能收益。'
        '流水线工程。' = '将任务划分为多个阶段并重叠执行，以提高吞吐率。'
        '在忽略错误成本的前提下，在预测下进行工作可以获得更好的性能。' = '当预测错误的恢复成本可控时，推测执行可以提升平均性能。'
        '计算机的内存成本是计算机成本的主要部分。速度最快容量最小的最顶层，速度最慢容量最大的最底层，使内存得到了速度与容量的统一。' = '存储层次将小而快的存储置于上层、将大而慢的存储置于下层，在成本、容量与访问速度之间取得平衡。'
        '冗余的组件可以在系统发生故障的时候代替实现故障检测，以应对物理错误。' = '冗余组件可用于故障检测、容错与服务恢复，以提高系统可靠性。'
        '包含了OS, I/O等的时间，决定了整个计算机系统的时间' = '墙钟时间包含操作系统与 I/O 等开销，反映用户观察到的整体响应时间。'
        'CPU处理一个任务的时间。包括用户CPU和系统CPU。' = 'CPU 时间是处理器执行某项任务所消耗的时间，包括用户态 CPU 时间与系统态 CPU 时间。'
        'CPU Time定义了一种单位时间，叫Clock priod，即一个时钟周期的时间。CPU clock period 的时间单位通常很小，如一种CPU的clock period为250ps=0.25ns=250×10e-12s' = 'Clock period（时钟周期）是一个时钟周期的持续时间。例如，250 ps = 0.25 ns = 250 × 10^-12 s。'
        'Clock period的倒数为Clock Frequency（rate），如4GHZ=4000MHZ=4×10e9HZ' = 'Clock frequency（时钟频率）是时钟周期的倒数。例如，4 GHz = 4000 MHz = 4 × 10^9 Hz。'
        '（3）在时钟频率和循环次数之间进行权衡： 更快的clock frequency会导致更多的时钟周期' = '（3）在时钟周期、每条指令的周期数与指令总数之间进行整体权衡。'
        '高主频的设计意味着一条指令可能需要更多个时钟周期才能完成' = '缩短时钟周期可能要求把工作拆分到更多流水级，从而改变每条指令的周期数与相关开销。'
        '优化时钟周期意味着复杂设计，信号需要经过更多晶体管，传输延迟增加，限制了时钟频率' = '增加单周期内的逻辑工作量会拉长关键路径，从而限制最高时钟频率。'
        'FLOPS：以浮点数标志的运行速度' = 'FLOPS：每秒完成的浮点运算次数。'
        'MFLOPS：每秒百万浮点运算数，10e6' = 'MFLOPS：每秒 10^6 次浮点运算。'
        'GFLOPS：每秒十亿浮点运算数，10e9' = 'GFLOPS：每秒 10^9 次浮点运算。'
        'TFLOPS：每秒万亿浮点运算数，10e12' = 'TFLOPS：每秒 10^12 次浮点运算。'
        'PFLOPS：每秒千万亿浮点运算数，10e15' = 'PFLOPS：每秒 10^15 次浮点运算。'
        '笔者认为使用单位法列公式比较稳妥' = '量纲分析可以用于检查性能公式中的单位一致性。'
        '为了追求更大的吞吐率，业界多采用多核处理器，利于并行计算。' = '在功耗与频率扩展受限的背景下，多核处理器通过线程级并行提高系统吞吐率。'
        '由于作者笔力有限，更多没有那么重要的知识在这里没有介绍，请参考原书《计算机组成与设计》（RISC-V版）（机械工业出版社）。' = '更多背景与推导可参考《计算机组成与设计：硬件/软件接口（RISC-V 版）》。'
    }
    1 = [ordered]@{
        '这一讲主要是介绍riscv的特性和整体的指令架构，还没有过多关注细节。' = '本讲概述 RISC-V 的主要特性与整体指令架构，具体编码细节将在下一讲展开。'
        '（1）算术指令——计算机必须能计算啊！' = '（1）算术指令：完成基本算术与逻辑运算。'
        'RISCV指令中的操作数有三种：寄存器、内存、立即数，接下来分别介绍' = 'RISC-V 指令涉及寄存器、内存数据和立即数三类操作对象，下面分别介绍。'
        'RISCV并不需要指令显式对齐，其硬件支持自动对齐。' = 'RISC-V 的基本指令说明了自然对齐访问；非对齐 load/store 是否由硬件透明支持取决于执行环境，也可能触发异常并由软件处理。'
        '立即数直接编码，计算最快；' = '立即数直接编码在指令中，可减少额外的数据读取；'
        '简单源于规整 ——本书RISCV指令全部为定长的32位。' = '简单源于规整：RV32I 基础指令采用固定的 32 位编码。'
        'opcode为操作码；funct3和funct7则为另外两个操作数。' = 'opcode 为主操作码；funct3 与 funct7 用于进一步区分具体操作。'
        'rs1表示另一个源操作数；rd为目的操作数；opcode为操作码；funct3则为另外的操作数。' = 'rs1 表示源寄存器，rd 表示目的寄存器；opcode 与 funct3 共同确定具体操作。'
        '本讲主要介绍了对RISCV的指令和cpu的认识，还学习了三个计算机设计的经典思想：' = '本讲概述 RISC-V 指令与 CPU 的基本关系，并归纳三个经典设计原则：'
        '正如教授所言，无论之后研究计算机的什么方向，这些经典又生动的思想会伴随我们始终。' = '这些设计原则也适用于后续的体系结构与系统研究。'
    }
    2 = [ordered]@{
        '这节课学习RISCV整体指令架构的细节，笔者认为学习指令主要是两部分：一部分是指令本身的结构，这与后面纯硬件部分的设计息息相关；另一部分是一些设计中的“匠心”，这些设计理念给人启迪，帮助我们发现计算机设计的优美，甚至给我们idea上的启示。但是交叉介绍两部分会破坏内容的连贯性，所以我们在最后统一阐释我第二部分的想法。' = '本讲首先介绍 RISC-V 指令结构及其硬件含义，再集中讨论编码规整性、扩展性等设计原则。'
        '如果你对内容不感兴趣，希望这张表也能帮到你。' = '下表汇总了本讲涉及的 RISC-V 指令类型。'
        '（1）R型指令：所有寄存器的计算' = '（1）R 型指令：寄存器之间的整数计算。'
        '（2）I型指令：所有立即数的计算，jalr 指令' = '（2）I 型指令：立即数计算、load 与 JALR 等指令。'
        '注意：对溢出的情况，寄存器简单将溢出的高位舍弃' = '对于整数溢出，基础整数加减指令保留结果的低 XLEN 位，不产生算术溢出异常。'
        '寄存器中的数需要符号扩展。' = '12 位立即数在参与运算前需要符号扩展到 XLEN 位。'
        'lw rd, rs1, imm x(ld)=Mem[x(rs1)+imm]' = 'lw rd, imm(rs1)    x(rd) = Mem[x(rs1) + sext(imm)]'
        'lb加载8位，lh加载16位->按照符号扩展；' = 'lb 加载 8 位，lh 加载 16 位，并对结果进行符号扩展；'
        'lbu/lhu直接0扩展；"lwu"不存在，因为都写32位了，无所谓什么扩展' = 'lbu 与 lhu 对结果进行零扩展；RV32I 不需要 lwu，因为目的寄存器本身只有 32 位。'
        'sw rs2, imm(rs1) Mem[x(rs2)]=Mem[x(rs1)+imm]' = 'sw rs2, imm(rs1)    Mem[x(rs1) + sext(imm)] = x(rs2)[31:0]'
        '写指令不需要"u"的区分，因为写的操作仅限于固定的位数（b8位，h16位），没有扩展一说。' = 'store 指令写入 rs2 的低位，不涉及将读出的窄数据扩展到寄存器，因此没有有符号/无符号版本之分。'
        '目标地址由分支指令的地址加上符号位扩展的偏移量组成，范围是2^13 字节。' = '目标地址由当前 PC 加上符号扩展后的偏移量得到；B 型立即数最低位隐含为 0，可表示约 ±4 KiB 的偶数字节偏移。'
        '这种别扭的指令格式我们将在最后讨论。' = '这种非连续的立即数字段布局将在本讲末尾讨论。'
        'AUIPC rd, imm x(rd)=PC+imm<<12' = 'AUIPC rd, imm    x(rd) = pc + (imm << 12)'
        'jal rd, imm PC+=imm; x(rd)=PC+4' = 'jal rd, imm    x(rd) = pc + 4; pc = pc + sext(imm)'
        'jal rd, label PC=label; x(rd)=PC+4' = 'jal rd, label    x(rd) = pc + 4; pc = label'
        '返回结果（默认返回寄存器为x1）' = '返回结果（标量返回值通常使用 x10/a0，必要时也使用 x11/a1）'
        '对于过程调用和其他细节的详细介绍在下一讲中会详细说明~！' = '过程调用及相关细节将在下一讲介绍。'
        '这一讲我们学习了risc-v所有常用指令，这些指令不只是考试的工具，其设计引人深思：' = '本讲介绍了 RV32I 中的常用指令，并总结其编码设计的两个特点：'
        '这也回应了前文我们埋下的伏笔。' = '这种布局延续了 S 型指令的字段位置。'
    }
}

$articlePages = @()
foreach ($lecture in 0..12) {
    $article = $byLecture[$lecture]
    $content = [string]$article.content

    if ($lecture -eq 0) {
        $content = Remove-ParagraphContaining $content '交流欢迎访问个人主页'
    }
    switch ($lecture) {
        3 { $content = Remove-ParagraphContaining $content '这一届我们学习了' }
        4 { $content = Remove-ParagraphContaining $content '这节课我们关注了RISC-V ALU的设计' }
        5 { $content = Remove-ParagraphContaining $content '我们经常在一起讨论' }
        9 { $content = Remove-ParagraphContaining $content '笔力有限只做粗浅解读' }
        11 { $content = Remove-ParagraphContaining $content '这节课我们一起完成了存储器的学习' }
    }
    if ($firstThreeCorrections.ContainsKey($lecture)) {
        foreach ($entry in $firstThreeCorrections[$lecture].GetEnumerator()) {
            $content = $content.Replace([string]$entry.Key, [string]$entry.Value)
        }
    }

    $content = Remove-ParagraphContaining $content '三日三更实属不易'
    $content = Remove-ParagraphContaining $content '共迎国庆，欢度中秋'
    $content = Remove-ParagraphContaining $content '各位大佬指正'

    foreach ($entry in $commonReplacements.GetEnumerator()) {
        $content = $content.Replace([string]$entry.Key, [string]$entry.Value)
    }
    if ($specificReplacements.ContainsKey($lecture)) {
        foreach ($entry in $specificReplacements[$lecture].GetEnumerator()) {
            $content = $content.Replace([string]$entry.Key, [string]$entry.Value)
        }
    }

    switch ($lecture) {
        0 {
            $content = [regex]::Replace($content, '<h3>1 Moore[^<]*</h3>', '<h3>1 Moore''s Law</h3>')
            $content = $content.Replace('<h3>2 Abstraction design</h3>', '<h3>2 Design for Abstraction</h3>')
            $content = $content.Replace('<h3>4 Parallel</h3>', '<h3>4 Parallelism</h3>')
            $content = $content.Replace('<h3>8 Redundant</h3>', '<h3>8 Dependability via Redundancy</h3>')
            $content = [regex]::Replace($content, '<p\b[^>]*>（3）在时钟频率和循环次数之间进行权衡：\s*<b>更快的clock frequency会导致更多的时钟周期</b></p>', '<p>（3）在时钟周期、每条指令的周期数与指令总数之间进行整体权衡。</p>')
            $content = $content.Replace('CPU的performance', 'CPU 性能')
            $content = $content.Replace('clock frequency', '时钟频率')
            $content = $content.Replace('CPU的设计要兼顾这两者', '处理器设计需要综合考虑关键路径、时钟周期和 CPI。')
            $content = $content.Replace('近些年Intel微处理器的时钟频率和功率放缓了增长趋势，因为其功率已经达到了实际极限，无法再用普通的商用冷却器冷却下来。由于冷却成本已经不能被性能提升带来的收益所覆盖，所以功耗很难再有提升。', '处理器频率与功耗的持续增长受到散热和能效约束，这一限制通常称为“功耗墙”。现代处理器因此更多依赖并行性与专用化来提升性能。')
            $content = $content.Replace('当前在集成电路中占有统治地位的是CMOS（互补型金属氧化半导体），其主要的能耗来源是动态能耗（晶体管开关过程中产生的能耗）。动态能耗取决于每个晶体管的负载电容和工作电压：', '现代数字集成电路主要采用 CMOS。动态能耗来自节点充放电，主要与负载电容和电压平方相关：')
            $content = $content.Replace('SPEC是许多计算机销售商出资建设的合作组织，目的是为现代计算机系统建立基础评测程序集。', 'SPEC 维护一组标准化基准程序，用于在统一规则下评估和比较计算机系统性能。')
        }
        1 {
            $content = $content.Replace('为了实现计算机的基本功能，我们必须设计以下几种指令：', '为实现计算机的基本功能，指令集至少需要支持以下几类操作：')
            $content = $content.Replace('依据冯诺依曼的基本架构，数据的计算和存储是分治的，因此，为了实现数据的流动运转，我们必须设计数据传送的指令。', '在经典存储程序体系结构中，计算与存储由不同部件承担，因此需要数据传送指令在寄存器与内存之间移动数据。')
            $content = $content.Replace('指令介于机器语言和高级语言之间，指令作为中间层架构，既能很好描述机器的操作，又具有一定的可读性。其反映了设计中“抽象”的作用。', '指令集架构（ISA）定义软件与硬件之间的接口；汇编表示使机器操作具有可读形式，体现了抽象在系统设计中的作用。')
            $content = $content.Replace('32个寄存器并不能满足计算机的存储需求，因此需要内存对数据进行存储。', 'RISC-V 采用 load/store 架构：算术运算主要在寄存器中完成，数据通过 load 与 store 指令在寄存器和内存之间移动。')
            $content = $content.Replace('（2）One byte based', '（2）字节寻址（byte-addressed）')
            $content = [regex]::Replace($content, '<p\b[^>]*><b>简单源于规整</b>——本书RISC-V指令全部为定长的32位。</p>', '<p><b>简单源于规整</b>：RV32I 基础指令采用固定的 32 位编码。</p>')
            $content = $content.Replace('带一个常数的算数指令/加载指令', '立即数算术指令与加载指令')
            $content = $content.Replace('immediate：补码，表示整数（立即数）', 'immediate：按指令语义解释并扩展的立即数字段')
            $content = $content.Replace('rs1：寄存器号-&gt;里面存了一个地址    imm：偏置量', 'rs1：保存基址的源寄存器；imm：地址偏移量')
            $content = $content.Replace('rs2:  存储的数据', 'rs2：提供待写入数据的源寄存器')
        }
        2 {
            $content = [regex]::Replace($content, '<p\b[^>]*><b>lw rd, rs1, imm\s+x\(ld\)=Mem\[x\(rs1\)\+imm\]</b></p>', '<p><b>lw rd, imm(rs1) &nbsp; x(rd) = Mem[x(rs1) + sext(imm)]</b></p>')
            $content = [regex]::Replace($content, '<p\b[^>]*><b>lb加载8位，lh加载16位-&gt;按照符号扩展；</b></p>', '<p><b>lb 加载 8 位，lh 加载 16 位，并对结果进行符号扩展；</b></p>')
            $content = [regex]::Replace($content, '<p\b[^>]*><b>sw rs2, imm\(rs1\)\s+Mem\[x\(rs2\)\]=Mem\[x\(rs1\)\+imm\]</b></p>', '<p><b>sw rs2, imm(rs1) &nbsp; Mem[x(rs1) + sext(imm)] = x(rs2)[31:0]</b></p>')
            $content = $content.Replace('beq rs1, rs2, imm                  if(x(rs1)==x(rs2))   goto imm', 'beq rs1, rs2, imm &nbsp; if (x(rs1) == x(rs2)) pc = pc + sext(imm)')
            $content = [regex]::Replace($content, '<p\b[^>]*><b>AUIPC rd, imm\s+x\(rd\)=PC\+imm&lt;&lt;12</b></p>', '<p><b>AUIPC rd, imm &nbsp; x(rd) = pc + (imm &lt;&lt; 12)</b></p>')
            $content = [regex]::Replace($content, '<p\b[^>]*><b>jal rd, imm\s+PC\+=imm; x\(rd\)=PC\+4</b></p>', '<p><b>jal rd, imm &nbsp; x(rd) = pc + 4; pc = pc + sext(imm)</b></p>')
            $content = [regex]::Replace($content, '<p\b[^>]*><b>jal rd, label\s+PC=label; x\(rd\)=PC\+4</b></p>', '<p><b>jal rd, label &nbsp; x(rd) = pc + 4; pc = label</b></p>')
            $content = $content.Replace('jal是RISC-V32中唯一的j型指令，常用于过程调用。因为它将PC修改为任意值，且将PC+4（当前指令的下一条指令）存储在寄存器中。', 'JAL 是 RV32I 中的 J 型指令，可用于过程调用或无条件跳转。它把 pc + 4 写入 rd，并跳转到 J 型立即数给出的 PC 相对目标。')
            $content = $content.Replace('有了LUI指令之后可以将长立即数的高20位存入，再将低12位存入（常规立即数）', 'LUI 可先构造高 20 位，再配合立即数指令形成更宽的常量；低 12 位的符号扩展细节通常由汇编器处理。')
            $content = $content.Replace('比如B型指令，其小端的1~4位，5~10位，符号位都与S型保持高度一致性，这种布局延续了 S 型指令的字段位置。', '例如，B 型与 S 型指令保持 rs1、rs2、funct3 和 opcode 字段位置一致，只重新排列立即数字段以适配分支偏移。')
        }
        3 {
            $content = $content.Replace('通过分析我们可以知道，需要保存的数为n、返回地址，在每次进入函数的时候，需要移动栈指针（2*4=8个字节），在n&gt;=1的分支条件下，需要恢复被保存的值。', '该递归过程需要保存参数 n 与返回地址。每次进入函数时，栈指针向低地址移动 8 字节；返回前再恢复保存的值。')
            $content = $content.Replace('fp（x8）为帧指针，是栈的最高地址；sp（x2）为栈指针，是栈的最低地址。一般我们不改动fp的值。帧指针的存在为寻址提供一种稳定性；此外，当传入的参数超过8个时，RISC-V约定将多余的参数放在帧指针上方（内存中），可以通过帧指针寻址。', 'x2（sp）是栈指针；x8（s0）可按 ABI 约定作为可选的帧指针。帧指针在函数执行期间提供稳定的栈帧基准，超出参数寄存器容量的参数则通过栈传递。')
            $content = $content.Replace('在RISC-V的汇编中，有许多约定俗成的条件，比如传参是x10~x17，返回地址写回x1，返回值写回x10，x11等，所以，在函数存在嵌套时，寄存器们无法同时保留多个值，需要有人将他们保存起来，这就是栈的作用，而栈指针就是栈的管理者，需要通过它来寻找保存在栈上的元素。', 'RISC-V ABI 使用 x10–x17（a0–a7）传递参数，使用 x10–x11（a0–a1）返回标量值，并通常使用 x1（ra）保存返回地址。嵌套调用需要按照调用约定保存必要的寄存器，栈与栈指针用于管理这些调用状态。')
        }
        4 {
            $content = $content.Replace('注：基于真值表的作图当然是提倡的，但是这里我们实现尽可能少的器件版本。', '下述实现依据真值表化简逻辑，并尽量减少所需器件。')
            $content = $content.Replace('简单将 32 个 1 位 ALU 串联起来能不能得到 32 位 ALU 呢？还需要考虑一些重要问题来修改我们设计的 ALU。', '串联 32 个 1 位 ALU 可以形成基本的 32 位 ALU，但还需要加入溢出检测与比较逻辑。')
            $content = $content.Replace('所以我们得出经验性的结论：', '由此可得：')
            $content = $content.Replace('当 32 位加法的符号位和次高位进位情况不同时，说明发生了 overflow .（如果是无符号数，只看最高位即可）', '对于二进制补码加法，当符号位的进位输入与进位输出不同时，发生有符号溢出；无符号溢出则由最高位进位判断。')
            $content = [regex]::Replace($content, '这个结论非常好，因为 carryout（结果进位标志）恰好是我们设计中的一部分，可以直接利用制作出 overflow\s*</p><p>\s*detection\.', 'carryout 已经是数据通路的一部分，因此可以直接用于构造 overflow detection。</p>')
            $content = $content.Replace('为了适应 RISC-V 的指令架构，我们还要赋予 ALU 比较 的功能，让硬件支持 slt（小于置位） 指令。', '为支持 RISC-V 的 slt（小于置位）指令，ALU 还需要实现比较功能。')
            $content = $content.Replace('至此我们完成了对 32 位 ALU 的设计，此后抽象成的部件如下：', '至此得到完整的 32 位 ALU，并可将其封装为下图所示的部件：')
            $content = $content.Replace('我们可以通过不断代换 ci 得到 c31 与 c0 的关系（隐去中间的进位），但是这样会导致 硬件过于复杂 ，我们无法承受这种开销。', '持续代换 ci 可以得到 c31 与 c0 的直接关系，但完全展开会造成较高的硬件复杂度。')
            $content = $content.Replace('gi 负责产生进位信号，pi 负责传递进位信号；这种想法又印证了我们之前所说的：抽象在计算机设计中的重要性.', 'gi 负责生成进位，pi 负责传播进位；这一分解再次体现了分层抽象在硬件设计中的作用。')
            $content = $content.Replace('这时候我们的进位运算就不再来自于ALU，而是我们设计的另外一个模块——超前进位单元', '此时进位由独立的超前进位单元生成，而不是沿各位 ALU 串行传播。')
            $content = $content.Replace('我们用大写的P集成了1个构建块中的传播信号，生成了“超级”传播信号.', '使用大写 P 汇总一个 4 位构建块中的传播信号，形成组传播信号。')
        }
        5 {
            $content = $content.Replace('上节课我们主要讨论了ALU和并行进位加法器的实现，本讲介绍乘法器、除法器的原理和实现，还有RISC-V浮点表示和浮点计算的内容。', '上一讲讨论了 ALU 与超前进位加法器；本讲介绍乘法器、除法器，以及 RISC-V 浮点表示与浮点运算。')
        }
        6 {
            $content = $content.Replace('上节课我们完成了所有运算部分的学习，这节课我们开始学习课程的核心章节之一——处理器。', '在完成算术部件后，本讲进入处理器设计，首先构建单周期数据通路。')
        }
        7 {
            $content = $content.Replace('依稀记得在前面的Lecture中，我们在指令架构的学习中，就有操作码（Opcode）的字段，这些操作码在这里就成为了处理器的分流依据。', '前述指令格式中的操作码（Opcode）字段用于区分指令类型；在处理器中，它进一步成为控制信号生成的主要输入。')
        }
        8 {
            $content = $content.Replace('细忖度之，可以看到，', '进一步观察可知，')
            $content = $content.Replace('细忖度之，可以看到，add指令在 EX阶段结束之后 x19就不会再动了，可以直接被后面的指令使用，不需要在寄存器中再存取一次了。', '进一步观察可知，add 指令在 EX 阶段结束后已经产生 x19 的新值，可以直接转发给后续指令，无需等待寄存器写回。')
        }
        9 {
            $content = $content.Replace('问题来了，我们如何实现插入bubble的操作？', '插入 bubble 需要同时控制流水线前段与后段的状态。')
        }
        10 {
            $content = $content.Replace('从这节课开始，我们开始学习存储器的内容，处理器的设计决定了计算机的时间能力，存储器的设计则决定计算机的空间能力，二者是计算机组成的核心部分，也是课程学习的重点和难点部分。', '本讲进入存储系统，介绍存储层次、常见存储介质与直接映射缓存。处理器和存储器共同决定系统的计算与数据访问能力。')
            $content = $content.Replace('我们利用局部性原理建立了如下的计算机存储架构：', '基于局部性原理，可以建立如下存储层次：')
            $content = $content.Replace('在这一节中，我们从一个简单的cache开始，处理器每次请求为一个字，且每个块由单个字组成。', '先从一个简化缓存开始：处理器每次请求一个字，每个缓存块也只包含一个字。')
            $content = $content.Replace('对于这种“判断题”，我们的设计理念简单而明确，对cache的存储内容添加一个“有效位”（valid bit），因为二进制的0和1本身就是判断电路。', '为区分有效与无效表项，每个缓存块增加一个有效位（valid bit）。')
        }
        11 {
            $content = $content.Replace('我们从architecture的角度看写操作，由于有多层存储结构，要保证写一致很重要。接下来我们介绍几种写的策略。', '从体系结构角度看，缓存写操作的核心问题是维护各层副本的一致性。下面介绍常见写策略。')
            $content = $content.Replace('这很好理解，当数据还在cache中的时候，我们不需要写回Memory，处理器访问cache即可获得正确的指令。', '在缓存块驻留期间，处理器可直接读取缓存中的最新数据，无需立即写回主存。')
        }
    }

    $content = $content.Replace('我们可以', '可以')
    $content = $content.Replace('我们必须', '必须')
    $content = $content.Replace('我们需要', '需要')
    $content = $content.Replace('我们不能', '不能')
    $content = $content.Replace('我们采用', '采用')
    $content = $content.Replace('我们使用', '使用')
    $content = $content.Replace('我们假设', '假设')
    $content = $content.Replace('我们的设计', '该设计')
    $content = $content.Replace('我们的', '该')

    $visibleTextTranslations = @{
        0 = [ordered]@{
            '<h3>1 Moore''s Law</h3>' = '<h3>1. 摩尔定律（Moore''s Law）</h3>'
            '<h3>2 Design for Abstraction</h3>' = '<h3>2. 面向抽象设计</h3>'
            '<h3>3 Make the common case fast</h3>' = '<h3>3. 优先优化常见情形</h3>'
            '<h3>4 Parallelism</h3>' = '<h3>4. 并行</h3>'
            '<h3>5 Pipelining</h3>' = '<h3>5. 流水线</h3>'
            '<h3>6 Prediction</h3>' = '<h3>6. 预测</h3>'
            '<h3>7 Hierarchy of Memory</h3>' = '<h3>7. 存储层次</h3>'
            '<h3>8 Dependability via Redundancy</h3>' = '<h3>8. 通过冗余提高可靠性</h3>'
            '<p>Performance depends on</p>' = '<p>处理器性能主要取决于：</p>'
            '<p>◼ Algorithm: affects IC, possibly CPI</p>' = '<p>◼ 算法：影响指令数（IC），也可能影响 CPI。</p>'
            '<p>◼ Programming language: affects IC, CPI</p>' = '<p>◼ 编程语言：影响指令数和 CPI。</p>'
            '<p>◼ Compiler: affects IC, CPI</p>' = '<p>◼ 编译器：影响指令数和 CPI。</p>'
            '<p>◼ Instruction set architecture: affects IC, CPI, Tc</p>' = '<p>◼ 指令集架构：影响指令数、CPI 与时钟周期。</p>'
        }
        1 = [ordered]@{
            '<h2>（一）Instruction——the Language of Computer</h2>' = '<h2>（一）指令：计算机的语言</h2>'
            '<figcaption>R-format Instruction</figcaption>' = '<figcaption>R 型指令格式</figcaption>'
            '<figcaption>I-format Instruction</figcaption>' = '<figcaption>I 型指令格式</figcaption>'
            '<figcaption>S-format Instruction</figcaption>' = '<figcaption>S 型指令格式</figcaption>'
        }
        4 = [ordered]@{
            '<p>detection.</p>' = ''
            '<figcaption>RISC-V ALU''s Verilog</figcaption>' = '<figcaption>RISC-V ALU 的 Verilog 实现</figcaption>'
        }
        6 = [ordered]@{
            '<figcaption>Full Datapath</figcaption>' = '<figcaption>完整数据通路</figcaption>'
        }
        7 = [ordered]@{
            '<h2>（一）Overview</h2>' = '<h2>（一）整体设计</h2>'
            '<figcaption>Datapath with Control</figcaption>' = '<figcaption>带控制信号的数据通路</figcaption>'
            '<figcaption>R-type Instruction</figcaption>' = '<figcaption>R 型指令数据通路</figcaption>'
            '<figcaption>load Instruction</figcaption>' = '<figcaption>load 指令数据通路</figcaption>'
            '<figcaption>store Instruction</figcaption>' = '<figcaption>store 指令数据通路</figcaption>'
            '<figcaption>beq Instruction</figcaption>' = '<figcaption>beq 指令数据通路</figcaption>'
            '<figcaption>example instruction</figcaption>' = '<figcaption>示例指令的控制信号</figcaption>'
            '<figcaption>Combinational Logic Control for example instructions</figcaption>' = '<figcaption>示例指令的组合逻辑控制器</figcaption>'
            '<figcaption>ROM-based Control</figcaption>' = '<figcaption>基于 ROM 的控制器</figcaption>'
            '<figcaption>Instruction Performance</figcaption>' = '<figcaption>指令执行时间</figcaption>'
        }
        8 = [ordered]@{
            '<figcaption>Time evaluation</figcaption>' = '<figcaption>执行时间评估</figcaption>'
            '<figcaption>multicycle datapath overview</figcaption>' = '<figcaption>多周期数据通路概览</figcaption>'
            '<figcaption>RISC-V pipeline overview</figcaption>' = '<figcaption>RISC-V 流水线概览</figcaption>'
            '<figcaption>Data Hazard</figcaption>' = '<figcaption>数据冒险</figcaption>'
            '<figcaption>Compiler''s Optimization</figcaption>' = '<figcaption>编译器调度优化</figcaption>'
        }
        9 = [ordered]@{
            '<figcaption>Pipelined Control</figcaption>' = '<figcaption>流水线控制信号</figcaption>'
            '<figcaption>Forwarding Unit Design</figcaption>' = '<figcaption>前递单元设计</figcaption>'
            '<figcaption>Data Hazard Condition</figcaption>' = '<figcaption>数据冒险条件</figcaption>'
            '<figcaption>Double Data Hazard</figcaption>' = '<figcaption>多重数据冒险</figcaption>'
            '<figcaption>MEM Hazard Condition</figcaption>' = '<figcaption>MEM 阶段冒险条件</figcaption>'
            '<figcaption>Path with Forwarding Unit</figcaption>' = '<figcaption>带前递单元的数据通路</figcaption>'
            '<figcaption>Load-use hazard Condition</figcaption>' = '<figcaption>load-use 冒险条件</figcaption>'
            '<figcaption>Load-Use Data Hazard Pipeline Example</figcaption>' = '<figcaption>load-use 数据冒险示例</figcaption>'
            '<figcaption>Datapath with Hazard Detection</figcaption>' = '<figcaption>带冒险检测的数据通路</figcaption>'
            '<figcaption>Branch Hazards</figcaption>' = '<figcaption>分支冒险</figcaption>'
            '<figcaption>Branch Taken</figcaption>' = '<figcaption>分支跳转</figcaption>'
            '<figcaption>One Cycle Bubble</figcaption>' = '<figcaption>插入一个周期的气泡</figcaption>'
            '<figcaption>Branch stalled</figcaption>' = '<figcaption>分支停顿</figcaption>'
            '<figcaption>Branch Stalled for 2 Cycles</figcaption>' = '<figcaption>分支停顿两个周期</figcaption>'
            '<figcaption>inner group branches misprediction</figcaption>' = '<figcaption>内层循环的分支预测错误</figcaption>'
            '<figcaption>2-bit Predictor</figcaption>' = '<figcaption>两位分支预测器</figcaption>'
            '<figcaption>Branch Target Buffer</figcaption>' = '<figcaption>分支目标缓冲区</figcaption>'
            '<figcaption>Pipeline with Expections</figcaption>' = '<figcaption>支持异常处理的流水线</figcaption>'
        }
        10 = [ordered]@{
            '<figcaption>IEC Standard Prefixes</figcaption>' = '<figcaption>IEC 标准存储单位前缀</figcaption>'
            '<figcaption>Computer Memory Hierarchy</figcaption>' = '<figcaption>计算机存储层次</figcaption>'
            '<figcaption>Data Transfer between levels</figcaption>' = '<figcaption>存储层次之间的数据传输</figcaption>'
            '<figcaption>SRAM vs DRAM</figcaption>' = '<figcaption>SRAM 与 DRAM 对比</figcaption>'
            '<figcaption>DRAM Architecture</figcaption>' = '<figcaption>DRAM 结构</figcaption>'
            '<figcaption>disk structure</figcaption>' = '<figcaption>磁盘结构</figcaption>'
            '<figcaption>cache reference</figcaption>' = '<figcaption>Cache 访问关系</figcaption>'
            '<figcaption>Directed Map</figcaption>' = '<figcaption>直接映射</figcaption>'
            '<figcaption>cache architecture</figcaption>' = '<figcaption>Cache 结构</figcaption>'
            '<figcaption>cache example</figcaption>' = '<figcaption>Cache 访问示例</figcaption>'
            '<figcaption>Address Subdivision</figcaption>' = '<figcaption>地址字段划分</figcaption>'
        }
        11 = [ordered]@{
            '<figcaption>Memory Stall</figcaption>' = '<figcaption>存储器停顿</figcaption>'
            '<figcaption>AMAT calculation</figcaption>' = '<figcaption>平均存储访问时间计算</figcaption>'
            '<figcaption>different map</figcaption>' = '<figcaption>不同映射方式</figcaption>'
            '<figcaption>Cache Architecture</figcaption>' = '<figcaption>Cache 结构</figcaption>'
            '<figcaption>Address Translation</figcaption>' = '<figcaption>地址转换</figcaption>'
            '<figcaption>Translation Using a Page Table</figcaption>' = '<figcaption>使用页表进行地址转换</figcaption>'
            '<figcaption>Mapping Pages to Storage</figcaption>' = '<figcaption>虚拟页到存储空间的映射</figcaption>'
            '<figcaption>Possible Combinations of Events</figcaption>' = '<figcaption>TLB 与 Cache 访问结果组合</figcaption>'
            '<figcaption>TLB and Cache Interaction</figcaption>' = '<figcaption>TLB 与 Cache 的协同访问</figcaption>'
        }
    }
    if ($visibleTextTranslations.ContainsKey($lecture)) {
        foreach ($entry in $visibleTextTranslations[$lecture].GetEnumerator()) {
            $content = $content.Replace([string]$entry.Key, [string]$entry.Value)
        }
    }
    $literalVisibleTranslations = [ordered]@{
        '2.CPU时间(CPU time)' = '2. CPU 时间'
        '2.1 clock period' = '2.1 时钟周期'
        'Performance depends on' = '处理器性能主要取决于：'
        '◼ Algorithm: affects IC, possibly CPI' = '◼ 算法：影响指令数（IC），也可能影响 CPI。'
        '◼ Programming language: affects IC, CPI' = '◼ 编程语言：影响指令数和 CPI。'
        '◼ Compiler: affects IC, CPI' = '◼ 编译器：影响指令数和 CPI。'
        '◼ Instruction set architecture: affects IC, CPI, Tc' = '◼ 指令集架构：影响指令数、CPI 与时钟周期。'
        'ALU control的构建' = 'ALU 控制器的构建'
        'ALU control' = 'ALU 控制器'
        'Time evaluation' = '执行时间评估'
        '（1）IF: Instruction fetch from memory （取指令）' = '（1）IF：从指令存储器取指'
        '3.Hazard Detection' = '3. 冒险检测'
        '4. Branch Hazards' = '4. 分支冒险'
        'cache reference' = 'Cache 访问关系'
        'Possible Combinations of Events' = 'TLB 与 Cache 访问结果组合'
    }
    foreach ($entry in $literalVisibleTranslations.GetEnumerator()) {
        $content = $content.Replace([string]$entry.Key, [string]$entry.Value)
    }
    $content = [regex]::Replace(
        $content,
        '<p\b[^>]*>这个结论非常好，因为 carryout（结果进位标志）恰好是我们设计中的一部分，可以直接利用制作出 overflow</p>\s*<p\b[^>]*>detection\.</p>',
        '<p>carryout 已经是数据通路的一部分，因此可以直接用于构造溢出检测逻辑。</p>',
        'IgnoreCase'
    )
    $content = $content.Replace('上一个Lecture', '上一讲')
    $content = $content.Replace('最后一篇Lecture', '最后一篇笔记')
    $content = [regex]::Replace($content, '\bLecture\s*(\d+)\b', '第 $1 讲', 'IgnoreCase')

    $content = $content.Replace('<p data-pid="EHCsGx6j"><b>前言</b></p>', '<h2>概述</h2>')
    $content = [regex]::Replace($content, '<p\b[^>]*>\s*<b>前言</b>\s*</p>', '<h2>概述</h2>', 'IgnoreCase')
    $content = [regex]::Replace($content, '<p\b[^>]*>\s*<b>结语</b>\s*</p>', '<h2>总结</h2>', 'IgnoreCase')
    $content = [regex]::Replace($content, '<p\b[^>]*>\s*<br\s*/?>\s*</p>', '', 'IgnoreCase')

    $imageCounter = [ref]0
    $content = [regex]::Replace($content, '<img\b[^>]*>', {
        param($match)
        $tag = $match.Value
        $altMatch = [regex]::Match($tag, 'alt="([^"]*)"', 'IgnoreCase')
        $alt = if ($altMatch.Success) { [System.Net.WebUtility]::HtmlDecode($altMatch.Groups[1].Value).Trim() } else { '' }
        if ($tag -match 'eeimg="1"' -or $tag -match 'www\.zhihu\.com/equation') {
            if (-not $alt) { return '' }
            return '<code class="inline-equation">' + (Escape-Html $alt) + '</code>'
        }

        $sourceMatch = [regex]::Match($tag, 'data-original="([^"]+)"', 'IgnoreCase')
        if (-not $sourceMatch.Success) { $sourceMatch = [regex]::Match($tag, 'src="([^"]+)"', 'IgnoreCase') }
        if (-not $sourceMatch.Success) { return '' }
        $source = [System.Net.WebUtility]::HtmlDecode($sourceMatch.Groups[1].Value)
        if ($source -notmatch '^https://') { return '' }

        $imageCounter.Value++
        $baseName = ('figure-{0:D2}' -f $imageCounter.Value)
        $jpgRelative = "assets/notes/lecture-$lecture/$baseName.jpg"
        $pngRelative = "assets/notes/lecture-$lecture/$baseName.png"
        $jpgDestination = Join-Path $repoRoot ($jpgRelative -replace '/', '\')
        $pngDestination = Join-Path $repoRoot ($pngRelative -replace '/', '\')

        if (Test-Path -LiteralPath $pngDestination) {
            $destination = $pngDestination
            $relativeAsset = $pngRelative
        } else {
            Download-Image $source $jpgDestination
            $signature = [IO.File]::ReadAllBytes($jpgDestination)
            if ($signature.Length -ge 8 -and $signature[0] -eq 0x89 -and $signature[1] -eq 0x50 -and $signature[2] -eq 0x4e -and $signature[3] -eq 0x47) {
                Move-Item -LiteralPath $jpgDestination -Destination $pngDestination -Force
                $destination = $pngDestination
                $relativeAsset = $pngRelative
            } else {
                $destination = $jpgDestination
                $relativeAsset = $jpgRelative
            }
        }
        $safeAlt = if ($alt) { Escape-Html $alt } else { "第 $lecture 讲技术示意图" }
        return '<img src="../../' + $relativeAsset + '" alt="' + $safeAlt + '" loading="lazy">'
    }, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)

    $content = [regex]::Replace($content, '\s+data-[\w-]+="[^"]*"', '', 'IgnoreCase')
    $content = [regex]::Replace($content, '\s+(?:width|height)="[^"]*"', '', 'IgnoreCase')
    $content = $content.Replace('class="highlight"', 'class="code-block"')
    $content = [regex]::Replace($content, ' class="(?:origin_image|content_image|zh-lightbox-thumb)[^"]*"', '', 'IgnoreCase')

    foreach ($linkedId in $byId.Keys) {
        $targetLecture = $byId[$linkedId]
        $content = $content.Replace("https://zhuanlan.zhihu.com/p/$linkedId", "lecture-$targetLecture.html")
    }
    $content = [regex]::Replace($content, '<a\b[^>]*href="https?://[^"]+"[^>]*>', {
        param($match)
        $tag = $match.Value
        if ($tag -notmatch '\btarget="_blank"') {
            $tag = $tag.Substring(0, $tag.Length - 1) + ' target="_blank">'
        }
        if ($tag -match '\brel="([^"]*)"') {
            $relations = @($Matches[1] -split '\s+' | Where-Object { $_ })
            if ($relations -notcontains 'noopener') { $relations += 'noopener' }
            if ($relations -notcontains 'noreferrer') { $relations += 'noreferrer' }
            $normalizedRel = ($relations | Select-Object -Unique) -join ' '
            return [regex]::Replace($tag, '\brel="[^"]*"', 'rel="' + $normalizedRel + '"', 'IgnoreCase')
        }
        return $tag.Substring(0, $tag.Length - 1) + ' rel="noopener noreferrer">'
    }, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
    $content = [regex]::Replace($content, '[ \t]+(?=\r?\n)', '')

    $created = [DateTimeOffset]::FromUnixTimeSeconds([int64]$article.created).ToOffset([TimeSpan]::FromHours(8)).ToString('yyyy年M月d日')
    $canonicalTitle = "计算机组成与体系结构 · 第 $lecture 讲"
    $previous = if ($lecture -gt 0) { '<a href="lecture-' + ($lecture - 1) + '.html">← 第 ' + ($lecture - 1) + ' 讲</a>' } else { '<span></span>' }
    $next = if ($lecture -lt 12) { '<a href="lecture-' + ($lecture + 1) + '.html">第 ' + ($lecture + 1) + ' 讲 →</a>' } else { '<span></span>' }
    $originalUrl = [string]$article.url
    $summary = $summaries[$lecture]

    $page = @"
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <meta name="description" content="$(Escape-Html $summary)">
    <title>$(Escape-Html $canonicalTitle) | 曲星宇</title>
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
    <main class="article-shell">
        <article class="note-article">
            <header class="article-header">
                <p class="eyebrow">计算机组成与体系结构</p>
                <h1>$(Escape-Html $canonicalTitle)</h1>
                <p class="article-summary">$(Escape-Html $summary)</p>
                <div class="article-meta">
                    <time datetime="$([DateTimeOffset]::FromUnixTimeSeconds([int64]$article.created).ToString('yyyy-MM-dd'))">$created</time>
                    <span>本站整理版</span>
                    <a href="$originalUrl" target="_blank" rel="noopener noreferrer">查看知乎原文</a>
                </div>
            </header>
            <div class="article-body">
$content
            </div>
        </article>
        <nav class="article-pagination" aria-label="相邻笔记">
            $previous
            <a href="index.html">系列目录</a>
            $next
        </nav>
    </main>
</body>
</html>
"@
    Set-Content -LiteralPath (Join-Path $seriesRoot "lecture-$lecture.html") -Value $page -Encoding UTF8

    $legacyTarget = "computer-organization/lecture-$lecture.html"
    $legacyPage = @"
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <meta http-equiv="refresh" content="0; url=$legacyTarget">
    <link rel="canonical" href="$legacyTarget">
    <title>正在前往第 $lecture 讲 | 曲星宇</title>
</head>
<body>
    <p>这篇笔记已归入“计算机组成与体系结构”课程分支，<a href="$legacyTarget">点击这里继续访问</a>。</p>
</body>
</html>
"@
    Set-Content -LiteralPath (Join-Path $notesRoot "lecture-$lecture.html") -Value $legacyPage -Encoding UTF8
    $articlePages += [pscustomobject]@{
        Lecture = $lecture
        Title = $canonicalTitle
        Created = $created
        Summary = $summary
        Url = "lecture-$lecture.html"
    }
    Write-Output "Generated Lecture $lecture ($($imageCounter.Value) images)"
}

$cards = foreach ($page in ($articlePages | Sort-Object Lecture)) {
@"
            <article class="note-card">
                <div class="note-card-meta"><span>第 $($page.Lecture) 讲</span><time>$($page.Created)</time></div>
                <h2><a href="$($page.Url)">$(Escape-Html $page.Title)</a></h2>
                <p>$(Escape-Html $page.Summary)</p>
                <a class="read-link" href="$($page.Url)">阅读笔记 <span aria-hidden="true">→</span></a>
            </article>
"@
}

$indexPage = @"
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <meta name="description" content="曲星宇的计算机组成与体系结构笔记，涵盖 RISC-V、处理器、流水线、存储系统、总线与 I/O。">
    <title>计算机组成与体系结构 | 课程笔记</title>
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
            <h1>计算机组成与体系结构</h1>
            <p>本系列共 13 篇，按照课程进度从 RISC-V 指令集出发，依次介绍运算部件、处理器数据通路、流水线、存储系统、总线与 I/O。本站版本在保留原有技术结构的基础上，统一了术语和书面表达。</p>
            <div class="series-meta">
                <span>13 篇笔记</span>
                <span>2025 年 9 月至 12 月</span>
                <a href="https://www.zhihu.com/column/$ColumnId" target="_blank" rel="noopener noreferrer">知乎原专栏</a>
            </div>
        </section>
        <section class="note-grid" aria-label="本课程笔记">
$($cards -join "`n")
        </section>
    </main>
</body>
</html>
"@
Set-Content -LiteralPath (Join-Path $seriesRoot 'index.html') -Value $indexPage -Encoding UTF8
Write-Output 'Generated computer organization series index.'

$hubPage = @"
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <meta name="description" content="曲星宇的武汉大学计算机学院课程学习分享，按照课程整理中文笔记、知识总结与学习思考。">
    <title>Lecture Notes | 曲星宇</title>
    <link rel="stylesheet" href="styles.css">
</head>
<body>
    <header class="site-header">
        <a class="site-name" href="../index.html">Xingyu Qu</a>
        <nav aria-label="课程笔记导航">
            <a aria-current="page" href="index.html">课程笔记</a>
            <a href="../index.html#service">返回主页</a>
        </nav>
    </header>
    <main class="notes-shell">
        <section class="series-intro">
            <p class="eyebrow">Lecture Notes</p>
            <h1>课程学习笔记</h1>
            <p>记录我在武汉大学计算机学院课程学习中的知识整理、课程思考与实践总结。所有内容均使用中文，并按照课程划分为独立系列。</p>
            <div class="series-meta">
                <span>武汉大学 · 计算机学院</span>
                <span>中文课程分享</span>
            </div>
        </section>
        <section class="note-grid course-grid" aria-label="课程系列">
            <article class="note-card course-card">
                <div class="note-card-meta"><span>2025 秋季</span><span>13 篇笔记</span></div>
                <h2><a href="computer-organization/index.html">计算机组成与体系结构</a></h2>
                <p>以 RISC-V 为主线，从指令集、运算部件和处理器数据通路出发，逐步介绍流水线、存储层次、总线与输入输出系统。</p>
                <div class="series-meta course-tags" aria-label="课程主题">
                    <span>RISC-V</span>
                    <span>处理器设计</span>
                    <span>存储系统</span>
                </div>
                <a class="read-link" href="computer-organization/index.html">进入课程分支 <span aria-hidden="true">→</span></a>
            </article>
        </section>
    </main>
</body>
</html>
"@
Set-Content -LiteralPath (Join-Path $notesRoot 'index.html') -Value $hubPage -Encoding UTF8
Write-Output 'Generated Lecture Notes hub.'
