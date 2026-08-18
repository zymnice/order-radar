# ============================================================
# 订单雷达 v1.0 — 猿急送订单监控（用户侧运行）
# 功能：抓取猿急送远程/全部订单 → 按技能关键词筛选 → 输出报告
# 用法：右键"使用 PowerShell 运行"，或命令行执行
#       powershell -ExecutionPolicy Bypass -File "订单雷达.ps1"
# 输出：控制台摘要 + Markdown 报告文件（同目录）
# ============================================================

$ErrorActionPreference = 'Continue'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# ---------- 技能关键词（可自行增删） ----------
$keywords = @(
    'ai', 'agent', 'mcp', '大模型', '智能体', '提示词', 'langchain', 'rag',
    'python', '脚本', '自动化',
    '数据', 'excel', 'csv', '清洗', '表格', '报表', '统计', '分析', '数据库',
    '爬虫', '抓取', '采集', '浏览器',
    '网页', '网站', '前端', '后端', 'api', '接口', '部署', '小程序', 'h5',
    '视频', '剪辑', 'ffmpeg', '字幕', '音频', '配音', '转码', '录屏',
    '文档', 'pdf', 'word', '办公',
    '图像', 'ocr', '识别', '图片',
    'git', 'github', 'wordpress'
)

# ---------- 风险词（灰色单提醒） ----------
$riskWords = @('外挂', '破解', '私服', '卡密', '绕过验证', '逆向', '棋牌', '博彩', '彩票', '爬隐私', '个人信息爬取')

# ---------- 抓取 ----------
$urls = @(
    @{ Name = '远程单';  Url = 'https://www.yuanjisong.com/job/remote' },
    @{ Name = '全部单';  Url = 'https://www.yuanjisong.com/job/allcity' }
)

$ua = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0 Safari/537.36'

$jobs = @()
foreach ($src in $urls) {
    Write-Host "抓取 $($src.Name) ..." -ForegroundColor Cyan
    try {
        $resp = Invoke-WebRequest -Uri $src.Url -UseBasicParsing -Headers @{ 'User-Agent' = $ua } -TimeoutSec 20
        $html = $resp.Content
    } catch {
        Write-Host "  !! 抓取失败: $($_.Exception.Message)" -ForegroundColor Red
        continue
    }

    # 提取所有唯一 job id
    $ids = [regex]::Matches($html, '/job/(\d+)') | ForEach-Object { $_.Groups[1].Value } | Select-Object -Unique

    foreach ($id in $ids) {
        if ($jobs.id -contains $id) { continue }

        # 标题：第一个指向该 job 的 <a> 文本
        $title = ''
        $am = [regex]::Match($html, "(?s)<a[^>]*href=`"[^`"]*/job/$id`"[^>]*>(.*?)</a>")
        if ($am.Success) {
            $title = ($am.Groups[1].Value -replace '<[^>]+>', ' ' -replace '\s+', ' ').Trim()
        }

        # 在 job id 周边窗口内找预算/工时
        $pos = $html.IndexOf("/job/$id")
        if ($pos -lt 0) { continue }
        $start = [Math]::Max(0, $pos - 500)
        $window = $html.Substring($start, [Math]::Min(4500, $html.Length - $start))

        $budget = ''
        $bm = [regex]::Match($window, '￥\s*([\d,]+)\s*元')
        if ($bm.Success) { $budget = $bm.Groups[1].Value }

        $days = ''
        $dm = [regex]::Match($window, '工时[：:]\s*([\d.]+)\s*天')
        if ($dm.Success) { $days = $dm.Groups[1].Value }

        $desc = ($window -replace '<[^>]+>', ' ' -replace '\s+', ' ').Trim()
        if ($desc.Length -gt 200) { $desc = $desc.Substring(0, 200) }

        $jobs += [PSCustomObject]@{
            id     = $id
            title  = $title
            budget = $budget
            days   = $days
            desc   = $desc
            url    = "https://www.yuanjisong.com/job/$id"
            source = $src.Name
        }
    }
    Start-Sleep -Milliseconds 500
}

# ---------- 关键词筛选 ----------
$matchJobs = @()
foreach ($j in $jobs) {
    $text = ($j.title + ' ' + $j.desc).ToLower()
    $hits = @($keywords | Where-Object { $text.Contains($_.ToLower()) })
    $risks = @($riskWords | Where-Object { $text.Contains($_) })
    $j | Add-Member -NotePropertyName hits  -NotePropertyValue ($hits -join '、')
    $j | Add-Member -NotePropertyName risks -NotePropertyValue ($risks -join '、')
    if ($hits.Count -gt 0) { $matchJobs += $j }
}

# ---------- 输出 ----------
$ts = Get-Date -Format 'yyyy-MM-dd HH:mm'
$stamp = Get-Date -Format 'yyyyMMdd-HHmm'
$outFile = Join-Path $PSScriptRoot "订单雷达报告-$stamp.md"

Write-Host ""
Write-Host "======== 订单雷达报告 ($ts) ========" -ForegroundColor Green
Write-Host "抓到订单总数: $($jobs.Count)    匹配技能: $($matchJobs.Count)" -ForegroundColor Yellow
Write-Host ""

if ($matchJobs.Count -eq 0) {
    Write-Host "暂无匹配技能的新单。稍后再跑一次试试。" -ForegroundColor Gray
}

$md = @("# 订单雷达报告", "", "生成时间: $ts", "", "订单总数: $($jobs.Count) | 匹配技能: $($matchJobs.Count)", "")

if ($matchJobs.Count -gt 0) {
    $md += "## 匹配你的技能 ($($matchJobs.Count) 单)", ""
    $md += "| 标题 | 预算(元) | 工时(天) | 命中关键词 | 风险提示 | 链接 |", "|---|---|---|---|---|---|"
    foreach ($j in $matchJobs) {
        $risk = if ($j.risks) { "⚠️ $($j.risks)" } else { '' }
        $md += "| $($j.title) | $($j.budget) | $($j.days) | $($j.hits) | $risk | [查看]($($j.url)) |"
        Write-Host "✔ [$($j.budget)元/$($j.days)天] $($j.title)" -ForegroundColor White
        Write-Host "    命中: $($j.hits)  $risk" -ForegroundColor DarkGray
    }
    $md += ""
}

if ($jobs.Count -gt $matchJobs.Count) {
    $md += "## 全部订单 ($($jobs.Count) 单)", ""
    foreach ($j in $jobs) {
        $mark = if ($j.hits) { '✔' } else { '·' }
        $md += "- $mark [$($j.budget)元/$($j.days)天] $($j.title) — $($j.url)"
    }
}

$md | Out-File -FilePath $outFile -Encoding UTF8
Write-Host ""
Write-Host "报告已保存: $outFile" -ForegroundColor Cyan
Write-Host "提示: 预算空=时间制/面议; 风险⚠️的单请谨慎评估（按你的原则不接灰色单）" -ForegroundColor DarkYellow
