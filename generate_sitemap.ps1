$baseDir = Get-Location
$sitemapPath = Join-Path $baseDir "sitemap.xml"
$robotsPath = Join-Path $baseDir "robots.txt"

$domains = @(
    "https://jichang365.com",
    "https://yunguidaohang.com"
)

# Root HTML files
$rootFiles = @(
    "about.html",
    "airports.html",
    "article.html",
    "avoid-scam.html",
    "knowledge.html",
    "ranking.html",
    "404.html"
)

# Get all article files
$articlesDir = Join-Path $baseDir "articles"
$articleFiles = Get-ChildItem -Path $articlesDir -Filter "*.html" | Select-Object -ExpandProperty Name

$dateStr = (Get-Date).ToString("yyyy-MM-dd")

$urls = @()

foreach ($domain in $domains) {
    # 1. Homepage
    $urls += @{ "loc" = "$domain/"; "priority" = "1.0"; "freq" = "daily" }
    $urls += @{ "loc" = "$domain/index.html"; "priority" = "1.0"; "freq" = "daily" }

    # 2. Root Pages
    foreach ($rf in $rootFiles) {
        $cleanName = $rf -replace '\.html$', ''
        $prio = if ($cleanName -in @("airports", "ranking", "knowledge")) { "0.9" } else { "0.8" }
        
        # Clean URL (without .html)
        $urls += @{ "loc" = "$domain/$cleanName"; "priority" = $prio; "freq" = "weekly" }
        # Full URL (with .html)
        $urls += @{ "loc" = "$domain/$rf"; "priority" = $prio; "freq" = "weekly" }
    }

    # 3. Article Pages
    foreach ($af in $articleFiles) {
        $cleanArtName = $af -replace '\.html$', ''
        
        # Clean URL (without .html)
        $urls += @{ "loc" = "$domain/articles/$cleanArtName"; "priority" = "0.8"; "freq" = "weekly" }
        # Full URL (with .html)
        $urls += @{ "loc" = "$domain/articles/$af"; "priority" = "0.8"; "freq" = "weekly" }
    }
}

# Build XML Content
$xml = '<?xml version="1.0" encoding="UTF-8"?>' + "`n"
$xml += '<urlset xmlns="http://www.sitemaps.org/schemas/sitemap/0.9">' + "`n"

foreach ($u in $urls) {
    $loc = $u.loc
    $prio = $u.priority
    $freq = $u.freq
    
    $xml += "    <url>`n"
    $xml += "        <loc>$loc</loc>`n"
    $xml += "        <lastmod>$dateStr</lastmod>`n"
    $xml += "        <changefreq>$freq</changefreq>`n"
    $xml += "        <priority>$prio</priority>`n"
    $xml += "    </url>`n"
}

$xml += '</urlset>'

$utf8NoBom = New-Object System.Text.UTF8Encoding($false)
[System.IO.File]::WriteAllText($sitemapPath, $xml, $utf8NoBom)
Write-Output "Successfully compiled sitemap.xml with $($urls.Count) URLs!"

# Build robots.txt Content
$robots = @"
User-agent: *
Allow: /

# Domestic AI Models & Search Spiders (国内 AI 大模型及搜索引擎)
# 字节跳动 / 豆包 AI
User-agent: Bytespider
User-agent: Bytespider-AI
Allow: /

# 百度 / 文心一言 & 搜索引擎
User-agent: Baiduspider
User-agent: Baiduspider-render
User-agent: Baiduspider-image
User-agent: Baiduspider-video
Allow: /

# 阿里 / 通义千问 & 夸克 & 神马搜索
User-agent: YisouSpider
User-agent: ShenmaSpider
Allow: /

# 搜狗搜索 (腾讯 / 微信 AI 搜索底层)
User-agent: Sogou web spider
User-agent: Sogou inst spider
Allow: /

# 360 智脑 / 奇虎 360 搜索
User-agent: 360Spider
Allow: /

# 深度求索 / DeepSeek AI
User-agent: DeepSeek-Bot
User-agent: DeepSeek
Allow: /

# 月之暗面 / Kimi AI
User-agent: Moonshot-Bot
User-agent: Kimi-Bot
Allow: /

# 智谱 AI / 智谱清言 (GLM-4)
User-agent: Zhipu-AI
User-agent: GLM-Bot
Allow: /

# 百川智能 / 零一万物 / 其它国内大模型 Agent
User-agent: Baichuan-Bot
User-agent: Yi-Bot
User-agent: MiniMax-Bot
Allow: /

# Global AI Models & Search Spiders (国外主流 AI 大模型及搜索引擎)
# OpenAI / ChatGPT / SearchGPT
User-agent: GPTBot
User-agent: ChatGPT-User
User-agent: OAI-SearchBot
Allow: /

# Anthropic / Claude 3.5
User-agent: ClaudeBot
User-agent: Claude-Web
Allow: /

# Google / Gemini & Search
User-agent: Googlebot
User-agent: Googlebot-Image
User-agent: Google-Extended
Allow: /

# Perplexity AI
User-agent: PerplexityBot
User-agent: Perplexity-Bot
Allow: /

# Microsoft / Bing & Copilot
User-agent: Bingbot
User-agent: Microsoft-Bot
User-agent: BingPreview
Allow: /

# Apple / Apple Intelligence
User-agent: Applebot
User-agent: Applebot-Extended
Allow: /

# xAI / Grok
User-agent: GrokBot
User-agent: xAIBot
Allow: /

# You.com / DuckDuckGo / Amazon / Meta AI
User-agent: YouBot
User-agent: DuckDuckGoBot
User-agent: Amazonbot
User-agent: Meta-ExternalAgent
User-agent: FacebookExternalHit
Allow: /

Sitemap: https://jichang365.com/sitemap.xml
Sitemap: https://yunguidaohang.com/sitemap.xml
"@

[System.IO.File]::WriteAllText($robotsPath, $robots, $utf8NoBom)
Write-Output "Successfully updated robots.txt!"
