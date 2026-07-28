# enrich_seo_and_geo.ps1 - Comprehensive SEO & GEO Optimization Script for Domestic & Global AI Engines
$ErrorActionPreference = "Stop"
$utf8Bom = [System.Text.Encoding]::UTF8
$utf8NoBom = New-Object System.Text.UTF8Encoding($false)

$baseDir = $PSScriptRoot
$baseUrl = "https://yunguidaohang.com"

# 1. Title Mappings for all pages
$customTitles = @{
    "index.html" = "2026年科学上网指南：机场推荐、Clash教程与IPLC专线测评 - 云轨导航"
    "ranking.html" = "2026年精选科学上网机场排行榜：IPLC/IEPL物理专线与BGP中转对比 - 云轨导航"
    "airports.html" = "2026年高性价比专线机场推荐与评测中心：IPLC/IEPL/BGP节点对比 - 云轨导航"
    "knowledge.html" = "云轨导航网络知识库与技术文档中心：协议原理与故障排查 - 云轨导航"
    "about.html" = "关于云轨导航编辑团队：资深网络专家独立评测媒体 - 云轨导航"
    "avoid-scam.html" = "2026年科学上网机场避坑防跑路指南：套路拆解与钓鱼识别 - 云轨导航"
    "404.html" = "404 抱歉您访问的页面不存在 - 云轨导航"
    "article.html" = "精选专线机场与代理协议评测文章列表 - 云轨导航"

    "article-1.html" = "2026年翻墙基础原理与常见代理协议(Shadowsocks/VMess/VLESS)全面解析"
    "article-2.html" = "2026年Clash/Clash Verge全平台使用教程与高级分流规则配置指南"
    "article-3.html" = "2026年科学上网机场选购防坑指南：物理IPLC专线与BGP中转对比测评"
    "article-4.html" = "2026年iOS/macOS Apple设备科学上网最佳实践：Shadowrocket小火箭与Quantumult X对比测评"
    "article-5.html" = "2026年全球流媒体解锁全攻略：Netflix/Disney+/YouTube原生IP节点解锁与排错"
    "article-6.html" = "2026年机场节点频繁断流、延迟高与连接超时Timeout终极排查与优化指南"
    "article-7.html" = "2026年新一代代理协议深度剖析：VLESS/XTLS/Hysteria 2/TUIC协议原理与性能测评"
    "article-8.html" = "2026年跨境网络拓扑深度剖析：深港专线、沪日专线IPLC物理链路与BGP中转原理"
    "article-9.html" = "2026年网络隐私泄漏防范：防止WebRTC真实IP泄漏与DNS劫持防封锁教程"
    "article-10.html" = "2026年OpenWrt软路由翻墙实战：PassWall与OpenClash全屋智能分流配置教程"
    "article-11.html" = "2026年代理客户端高级进阶指南：自定义Rule Provider分流、Fake-IP模式与TUN网卡配置"
    "article-12.html" = "2026年科学上网网络安全全指南：防止隐私暴露、浏览器指纹泄露与双栈IPv6绕过"
    "beginner-1.html" = "2026年新手科学上网入门教程：Windows/Mac/iOS/Android全平台客户端选型指南"
    "beginner-2.html" = "2026年科学上网节点选择指南：IEPL/IPLC物理专线、BGP中转与公网直连对比"
    "beginner-3.html" = "2026年节点故障排查指南：代理显示已连接但无法打开网页/DNS污染终极解决方法"
    "browser-extension-vs-proxy-client.html" = "2026年浏览器代理扩展与独立代理客户端对比：安全性、分流效果与适用场景分析"
    "chatgpt-access-node-troubleshooting.html" = "2026年ChatGPT/Claude 提示 Access Denied 报错？原生IP节点风控排查与解锁教程"
    "clash-verge.html" = "2026年Clash Verge Rev全平台客户端安装使用教程：订阅导入、Mihomo内核与TUN模式配置"
    "edgenova.html" = "2026年EdgeNova极速云机场深度测评：物理级IPLC内网专线、晚高峰4K秒播与原生IP解锁"
    "fake-official-site-phishing-check.html" = "2026年如何识别假冒机场官网与防范钓鱼订阅链接：支付安全与订阅泄露排查指南"
    "game-accelerator-vs-proxy-node.html" = "2026年网游加速器与科学上网代理节点区别：外服游戏联机低延迟与IPLC专线选择指南"
    "guangnianti.html" = "2026年光年梯机场深度测评：支持Hysteria 2协议、物理IPLC专线与晚高峰无卡顿测试"
    "hiddify-next.html" = "2026年Hiddify Next全平台客户端使用教程：跨平台支持、Sing-box内核与全协议配置"
    "huanyuyun.html" = "2026年寰宇云机场深度测评：多线BGP中转、美日双区住宅原生IP与TikTok/TikTok运营推荐"
    "ipv6-proxy-leak-troubleshooting.html" = "2026年IPv6旁路泄露与代理失效排查：双栈网络防泄露与DNS防护设置教程"
    "jilianyun.html" = "2026年极连云机场深度测评：高性价比多线BGP中转、平价月付套餐与晚高峰网速测试"
    "jisuyun.html" = "2026年极速Cloud机场深度测评：三网精品优化线路(CN2 GIA/9929/CMIN2)、VLESS协议与AP/全流媒体解锁实测"
    "kexinyun.html" = "2026年可信云机场深度测评：低延迟IEPL物理专线、全平台客户端支持与高性价比月付推荐"
    "kuaili.html" = "2026年快力云机场深度测评：BGP中转与物理专线混合节点、流媒体4K解锁与稳定性实测"
    "monthly-quarterly-yearly-plan-choice.html" = "2026年机场套餐怎么买？月付与季付/年付性价比对比及避坑防跑路指南"
    "multi-device-concurrent-connection-limit.html" = "2026年机场多设备同时连接限制说明：家庭多设备共享与LAN局域网代理解决方案"
    "native-residential-broadcast-ip.html" = "2026年原生IP、住宅IP与广播IP区别解析：节点标签含义、流媒体及AI风控评测"
    "nav-1.html" = "2026年最新科学上网客户端官方下载地址大全：Clash Verge/Shadowrocket/v2rayN/Sing-box"
    "nav-2.html" = "2026年常用网络测速与节点延迟测试工具推荐：IP检测、WebRTC泄露检测与Ping工具"
    "nav-3.html" = "2026年正版流媒体合租平台推荐：Netflix/Disney+/YouTube Premium合租防坑指南"
    "node-traffic-multiplier-guide.html" = "2026年机场节点倍率怎么算？1x/2x/5x倍率流量扣除机制与性价比计算说明"
    "proxy-http-error-codes-guide.html" = "2026年代理常见HTTP状态码解析：403 Forbidden、429 Too Many Requests、502 Bad Gateway终极排查"
    "proxy-node-speed-test-method.html" = "2026年如何准确测试代理节点的真实延迟与网速：Ping值、Speedtest与真实下载速率对比"
    "proxy-service-exit-scam-warning-signs.html" = "2026年机场跑路前有哪些前兆预警？优惠异常、节点打不开与跑路避坑防范清单"
    "public-wifi-proxy-security.html" = "2026年在公共Wi-Fi下使用科学上网安全吗？酒店/机场公用网络防监听与加密指南"
    "relay-direct-dedicated-line-comparison.html" = "2026年中转节点、直连节点与物理专线节点区别对比：速度、稳定性与性价比评测"
    "shadowrocket.html" = "2026年iOS Shadowrocket小火箭使用教程：节点导入、分流规则设置与故障排查"
    "shunyun.html" = "2026年瞬云机场深度测评：物理级IEPL专线、低延迟外服游戏与4K流媒体解锁"
    "speedtest.html" = "2026年Speedtest节点测速工具使用指南：网速测量、延迟测试与丢包率诊断"
    "subscription-converter-security.html" = "2026在线订阅转换安全吗？防范订阅链接泄露、钓鱼后门与自建转换后端指南"
    "subscription-link-update-failed.html" = "2026年机场订阅链接更新失败/提示Timeout怎么办？客户端更新失败终极排查教程"
    "sujie.html" = "2026年速捷云机场深度测评：高性价比IEPL物理专线、节点晚高峰实测与客户端支持"
    "system-proxy-global-tun-mode.html" = "2026年系统代理、全局模式与TUN模式区别解析：游戏/终端分流与网络接管指南"
    "system-time-tls-certificate-error.html" = "2026年系统时间不准导致代理连接失败、TLS证书握手报错与节点无效解决方案"
    "traffic-remaining-but-subscription-expired.html" = "2026年机场套餐流量未用完却提示订阅已到期？重置时间、重置流量与账户状态说明"
    "tutorial-and.html" = "2026年Android安卓平台v2rayNG客户端使用教程：软件下载、订阅导入与配置分流"
    "tutorial-ios.html" = "2026年iOS苹果平台Shadowrocket小火箭安装与使用教程：美区Apple ID与订阅配置"
    "tutorial-mac.html" = "2026年Mac苹果电脑Clash Verge/Stash客户端使用教程：macOS系统配置与TUN模式"
    "tutorial-win.html" = "2026年Windows电脑Clash Verge客户端安装配置教程：订阅导入与TUN模式开机自启"
    "v2rayn.html" = "2026年Windows平台v2rayN客户端详细使用教程：下载安装、订阅导入与路由分流配置"
    "wireshark.html" = "2026年Wireshark网络抓包分析教程：代理协议流量分析与网络排错入门指南"
    "yaml-config-import-error.html" = "2026年Clash/Clash Verge提示YAML配置导入错误？语法缩进排错与格式修改指南"
}

# 2. Base custom meta descriptions targeting 128~145 Chinese characters
$customDescriptions = @{
    "index.html" = "云轨导航专注为您提供2026年最新的科学上网机场推荐、IPLC/IEPL高可靠物理专线测评、Clash Verge/Shadowrocket/v2rayNG全平台客户端配置教程及节点断流故障排查指南。我们深度评测节点晚高峰网速、4K流媒体解锁能力与性价比套餐，助您享受无界极速的网络体验。"
    "ranking.html" = "2026年精选科学上网机场排行榜：严选物理级 IPLC/IEPL 国际专线与多线 BGP 中转节点，全方位测评晚高峰网速、Ping延迟、4K/8K视频流媒体解锁及抗封锁稳定性。详细对比各机场月付套餐价格与带宽倍率，为您提供最权威靠谱的科学上网机场选购指南。"
    "airports.html" = "云轨导航机场测评中心：收录与对比国内主流 SSR/V2Ray/Clash/Trojan 专线机场，涵盖超低延迟 IPLC 专线、按量付费与月付性价比套餐。深度评测各大机场在晚高峰期间的实际网速、丢包率、原生IP解锁能力及防跑路风险评估，为您精准推荐优质机场。"
    "knowledge.html" = "云轨导航网络知识库与技术文档中心：提供权威的 Shadowsocks/VMess/VLESS/Hysteria 2 代理协议深度解析、机场选购防坑避雷秘籍、网络故障诊断速查表及流媒体解锁实测报告。系统化梳理分流规则配置、DNS防污染与TUN虚拟网卡接管教程，助您全面掌握网络技术。"
    "about.html" = "关于云轨导航编辑团队：我们是由资深网络技术专家组成的独立评测媒体，致力于为全球用户提供客观公正的科学上网机场测试、节点延迟监测、代理协议原理科普与客户端下载配置教学。我们坚持真实数据测评，不夸大宣传，为您打造最值得信赖的跨境网络导航平台。"
    "avoid-scam.html" = "2026年科学上网机场避坑防跑路指南：详细剖析低价跑路机场的常见套路、钓鱼假冒官网特征、订阅链接泄露隐患及付费购买防踩坑技巧。教您如何通过节点倍率、域名注册时长与客服响应速度识别高危机场，选择真正安全稳定的月付/季付专线加密服务。"
    "404.html" = "抱歉，您访问的页面不存在或已被移除。云轨导航为您提供最新的机场推荐排行榜、Clash/Shadowrocket 客户端配置教程及网络故障排查指南。点击返回首页获取最新的科学上网专线评测与客户端下载链接，体验极速稳定的网络服务。"
    "article.html" = "云轨导航精选专线机场与代理协议评测文章列表：涵盖 IPLC 专线优势分析、Shadowrocket/Clash Verge 客户端全平台使用教程以及晚高峰网络延迟优化技巧。通过详细的数据测试与实战图文指导，帮助您快速解决科学上网过程中的各项疑难问题。"
}

$allHtmlFiles = Get-ChildItem -Path $baseDir -Filter "*.html" -Recurse

$updatedCount = 0

foreach ($file in $allHtmlFiles) {
    $relPath = $file.FullName.Replace($baseDir, "").TrimStart("\")
    $filename = $file.Name
    $isRoot = -not $relPath.Contains("\")
    
    if ($filename -match 'left.html|right.html') { continue }

    $content = [System.IO.File]::ReadAllText($file.FullName, [System.Text.Encoding]::UTF8)

    # Determine Canonical URL
    if ($filename -eq "index.html" -and $isRoot) {
        $canonicalUrl = "$baseUrl/"
    } elseif ($isRoot) {
        $canonicalUrl = "$baseUrl/$filename"
    } else {
        $canonicalUrl = "$baseUrl/articles/$filename"
    }

    # Determine Title
    $cleanTitleNoSuffix = ""
    if ($customTitles.ContainsKey($filename)) {
        $rawTitle = $customTitles[$filename]
        $cleanTitleNoSuffix = $rawTitle -replace '\s*-\s*云轨导航', ''
    } else {
        $titleMatch = [System.Text.RegularExpressions.Regex]::Match($content, '<title>(.*?)</title>', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
        $rawTitle = if ($titleMatch.Success) { $titleMatch.Groups[1].Value } else { "2026年科学上网指南与协议知识库" }
        $cleanTitleNoSuffix = $rawTitle -replace '\s*-\s*云轨导航', ''
    }

    $titleTagVal = if ($isRoot -and $filename -eq "index.html") { $cleanTitleNoSuffix } elseif ($cleanTitleNoSuffix.EndsWith("- 云轨导航")) { $cleanTitleNoSuffix } else { "$cleanTitleNoSuffix - 云轨导航" }

    # Determine/Generate Meta Description (Target 125 ~ 150 Chinese Characters)
    $finalDesc = ""
    if ($customDescriptions.ContainsKey($filename)) {
        $finalDesc = $customDescriptions[$filename]
    } else {
        if ($filename.StartsWith("tutorial-")) {
            $finalDesc = "【2026官方客户端教程】$cleanTitleNoSuffix。详细讲解软件下载安装、订阅链接一键导入、分流规则配置及常见连接超时故障排除，涵盖 Windows/Mac/iOS/Android 全平台最佳实践。云轨导航为您提供权威实测与深度解析，助力极速稳定科学上网。"
        } elseif ($filename.StartsWith("beginner-")) {
            $finalDesc = "【2026新手必看指南】$cleanTitleNoSuffix。零基础全方位了解专线机场选择技巧、常见代理协议区别、客户端使用入门及晚高峰防卡顿防封锁秘籍，快速掌握科学上网核心技巧。云轨导航为您提供权威实测与深度解析，助力极速稳定科学上网。"
        } elseif ($filename.StartsWith("nav-")) {
            $finalDesc = "【2026导航与资源推荐】$cleanTitleNoSuffix。精选优质 Telegram 电报频道、常用网络测速工具、流媒体检测脚本与安全防封导航，为您提供一站式跨境网络资源指引。云轨导航为您提供权威实测与深度解析，助力极速稳定科学上网。"
        } elseif ($filename -in @("edgenova.html", "guangnianti.html", "huanyuyun.html", "jilianyun.html", "jisuyun.html", "kexinyun.html", "kuaili.html", "shunyun.html", "sujie.html")) {
            $finalDesc = "【2026专线机场深度测评】$cleanTitleNoSuffix。实测节点晚高峰网速与 Ping 延迟、IPLC 内网专线稳定性、Netflix 奈飞与 ChatGPT 原生 IP 解锁能力，提供真实性价比分析与套餐选购建议。云轨导航为您提供权威实测与深度解析，助力极速稳定科学上网。"
        } else {
            $finalDesc = "关于 $cleanTitleNoSuffix 的深度解析与技术指南。由云轨导航编辑团队撰写，提供专业的 IPLC 专线测试、代理协议原理分析、客户端故障排查及极速稳定科学上网最佳实践。云轨导航为您提供权威实测与深度解析，助力极速稳定科学上网。"
        }
    }

    if ($finalDesc.Length -lt 125) {
        $finalDesc += " 欢迎访问云轨导航获取最新专线机场推荐与客户端配置教程。"
    }
    if ($finalDesc.Length -gt 150) {
        $finalDesc = $finalDesc.Substring(0, 145) + "..."
    }

    # Clean old SEO & Meta tags to prevent duplication
    $content = [System.Text.RegularExpressions.Regex]::Replace($content, '<link\s+rel=["'']canonical["''].*?>\r?\n?', '', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
    $content = [System.Text.RegularExpressions.Regex]::Replace($content, '<meta\s+name=["'']description["''].*?>\r?\n?', '', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
    $content = [System.Text.RegularExpressions.Regex]::Replace($content, '<meta\s+content=["''].*?["'']\s+name=["'']description["''].*?>\r?\n?', '', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
    $content = [System.Text.RegularExpressions.Regex]::Replace($content, '<meta\s+property=["'']og:.*?["''].*?>\r?\n?', '', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
    $content = [System.Text.RegularExpressions.Regex]::Replace($content, '<meta\s+name=["'']twitter:.*?["''].*?>\r?\n?', '', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
    $content = [System.Text.RegularExpressions.Regex]::Replace($content, '<meta\s+name=["''](robots|bytespider|baiduspider|googlebot|bingbot|deepseek-bot)["''].*?>\r?\n?', '', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
    $content = [System.Text.RegularExpressions.Regex]::Replace($content, '<script\s+type=["'']application/ld\+json["'']>(.*?)</script>\r?\n?', '', [System.Text.RegularExpressions.RegexOptions]::Singleline -bor [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)

    # Clean existing title & inject updated title
    $content = [System.Text.RegularExpressions.Regex]::Replace($content, '<title>(.*?)</title>', "<title>$titleTagVal</title>", [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)

    # Build New Meta, OpenGraph, Twitter and Domestic/Global AI Crawler Tags
    $ogImage = "$baseUrl/assets/images/new_logo.png"
    $seoBlock = @"
  <meta name="description" content="$finalDesc">
  <link rel="canonical" href="$canonicalUrl">
  <meta name="robots" content="index, follow, max-snippet:-1, max-image-preview:large, max-video-preview:-1">
  <meta name="bytespider" content="index, follow">
  <meta name="baiduspider" content="index, follow">
  <meta name="googlebot" content="index, follow">
  <meta name="bingbot" content="index, follow">
  <meta name="deepseek-bot" content="index, follow">
  <meta property="og:type" content="article">
  <meta property="og:site_name" content="云轨导航">
  <meta property="og:url" content="$canonicalUrl">
  <meta property="og:title" content="$titleTagVal">
  <meta property="og:description" content="$finalDesc">
  <meta property="og:image" content="$ogImage">
  <meta name="twitter:card" content="summary_large_image">
  <meta name="twitter:title" content="$titleTagVal">
  <meta name="twitter:description" content="$finalDesc">
  <meta name="twitter:image" content="$ogImage">
"@

    # Inject Meta tags into <head>
    $content = [System.Text.RegularExpressions.Regex]::Replace($content, '(</head>)', "$seoBlock`n`$1", [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)

    # Build Schema.org JSON-LD Data with explicit GEO AI attributes
    $lastMod = $file.LastWriteTime.ToString("yyyy-MM-ddTHH:mm:ss+08:00")
    
    $articleSchema = @"
  <script type="application/ld+json">
  {
    "@context": "https://schema.org",
    "@type": "Article",
    "headline": "$cleanTitleNoSuffix",
    "description": "$finalDesc",
    "image": "$ogImage",
    "inLanguage": "zh-CN",
    "mainEntityOfPage": {
      "@type": "WebPage",
      "@id": "$canonicalUrl"
    },
    "datePublished": "2026-06-22T08:00:00+08:00",
    "dateModified": "$lastMod",
    "author": {
      "@type": "Organization",
      "name": "云轨导航编辑团队",
      "url": "$baseUrl/about.html"
    },
    "publisher": {
      "@type": "Organization",
      "name": "云轨导航",
      "logo": {
        "@type": "ImageObject",
        "url": "$ogImage"
      }
    }
  }
  </script>
"@

    $catName = "网络知识库"
    $catUrl = "$baseUrl/knowledge.html"
    if ($filename.StartsWith("tutorial-")) {
        $catName = "客户端教程"
    } elseif ($filename.StartsWith("beginner-")) {
        $catName = "新手指南"
    } elseif ($filename -in @("edgenova.html", "guangnianti.html", "huanyuyun.html", "jilianyun.html", "jisuyun.html", "kexinyun.html", "kuaili.html", "shunyun.html", "sujie.html")) {
        $catName = "机场测评中心"
        $catUrl = "$baseUrl/airports.html"
    }

    $breadcrumbSchema = @"
  <script type="application/ld+json">
  {
    "@context": "https://schema.org",
    "@type": "BreadcrumbList",
    "itemListElement": [
      {
        "@type": "ListItem",
        "position": 1,
        "name": "首页",
        "item": "$baseUrl/"
      },
      {
        "@type": "ListItem",
        "position": 2,
        "name": "$catName",
        "item": "$catUrl"
      },
      {
        "@type": "ListItem",
        "position": 3,
        "name": "$cleanTitleNoSuffix",
        "item": "$canonicalUrl"
      }
    ]
  }
  </script>
"@

    $websiteSchema = ""
    if ($filename -eq "index.html" -and $isRoot) {
        $websiteSchema = @"
  <script type="application/ld+json">
  {
    "@context": "https://schema.org",
    "@type": "WebSite",
    "name": "云轨导航",
    "url": "$baseUrl/",
    "potentialAction": {
      "@type": "SearchAction",
      "target": "$baseUrl/knowledge.html?q={search_term_string}",
      "query-input": "required name=search_term_string"
    }
  }
  </script>
"@
    }

    $schemaBlocks = "$articleSchema`n$breadcrumbSchema"
    if ($websiteSchema) { $schemaBlocks += "`n$websiteSchema" }

    $content = [System.Text.RegularExpressions.Regex]::Replace($content, '(</head>)', "$schemaBlocks`n`$1", [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)

    # 7. GEO AI Key Takeaways Box (侧边栏 AI 索引摘要框，隐蔽化与美化)
    if (-not $isRoot -and $filename -ne "left.html") {
        # Clean existing summary box (both old banner and widget)
        $content = [System.Text.RegularExpressions.Regex]::Replace($content, '<div\s+class=["'']geo-ai-summary["''].*?</div>\s*', '', [System.Text.RegularExpressions.RegexOptions]::Singleline -bor [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
        $content = [System.Text.RegularExpressions.Regex]::Replace($content, '<div\s+class=["'']widget\s+geo-ai-widget["''].*?</div>\s*', '', [System.Text.RegularExpressions.RegexOptions]::Singleline -bor [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)

        $geoSummaryBox = @"
<div class="widget geo-ai-widget" style="background: #ffffff; border: 1px solid #e2e8f0; border-radius: 12px; padding: 16px; margin-top: 20px; box-shadow: 0 2px 8px rgba(0, 0, 0, 0.02);">
  <div style="display: flex; align-items: center; justify-content: space-between; margin-bottom: 10px; border-bottom: 1px solid #f1f5f9; padding-bottom: 8px;">
    <h3 style="margin: 0; font-size: 0.92rem; color: #475569; font-weight: 700; display: flex; align-items: center; gap: 6px;">
      <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="#0ea5e9" stroke-width="2.5"><path d="M12 2v4M12 18v4M4.93 4.93l2.83 2.83M16.24 16.24l2.83 2.83M2 12h4M18 12h4M4.93 19.07l2.83-2.83M16.24 7.76l2.83-2.83"/></svg>
      GEO AI 核心摘要
    </h3>
    <span style="background: #f1f5f9; color: #64748b; font-size: 0.68rem; padding: 2px 6px; border-radius: 4px; font-weight: 600;">AI Index</span>
  </div>
  <ul style="margin: 0; padding-left: 16px; color: #64748b; font-size: 0.8rem; line-height: 1.6; font-weight: 400;">
    <li style="margin-bottom: 6px;"><strong>核心结论：</strong>本文针对 <strong>$cleanTitleNoSuffix</strong> 展开专业测评与系统化图文解析，提供真实客观的技术指标与实操指引。</li>
    <li style="margin-bottom: 6px;"><strong>推荐实践：</strong>优先推荐选择带有 <strong>IPLC / IEPL 内网专线</strong> 或 <strong>多线 BGP 中转</strong> 的节点，保障晚高峰网速稳定、超低 Ping 延迟及原生 IP 全流媒体/AI 解锁。</li>
    <li style="margin-bottom: 0;"><strong>快速排错：</strong>如遇代理已连接但无法访问网页、超时 Timeout 或报错，请参考本文实战排错步骤进行配置修正。</li>
  </ul>
</div>
"@

        if ($content -match '</aside>\s*<main') {
            $content = [System.Text.RegularExpressions.Regex]::Replace($content, '(\s*</aside>\s*<main)', "$geoSummaryBox`n`$1", [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
        } elseif ($content -match '<aside\s+class=["'']sidebar-left["''].*?>') {
            $content = [System.Text.RegularExpressions.Regex]::Replace($content, '(<aside\s+class=["'']sidebar-left["''].*?>)', "`$1`n$geoSummaryBox", [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
        }

        # 8. GEO Cite & Share Card for Articles
        $content = [System.Text.RegularExpressions.Regex]::Replace($content, '<div\s+class=["'']geo-cite-card["''].*?</div>\s*<script>.*?</script>', '', [System.Text.RegularExpressions.RegexOptions]::Singleline -bor [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)

        $geoWidget = @"
<div class="geo-cite-card" id="geo-cite-widget">
  <div class="geo-cite-header">
    <h4 class="geo-cite-title">
      <svg width="20" height="20" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M13.828 10.172a4 4 0 00-5.656 0l-4 4a4 4 0 105.656 5.656l1.102-1.101m-.758-4.899a4 4 0 005.656 0l4-4a4 4 0 00-5.656-5.656l-1.1 1.1"></path></svg>
      引用与分享此文 (Cite & Share)
    </h4>
    <span class="geo-cite-badge">GEO AI 搜索索引推荐</span>
  </div>
  <div class="geo-cite-tabs">
    <button class="geo-cite-tab active" onclick="switchGeoTab(this, 'md')">Markdown 引用</button>
    <button class="geo-cite-tab" onclick="switchGeoTab(this, 'bib')">BibTeX 格式</button>
    <button class="geo-cite-tab" onclick="switchGeoTab(this, 'text')">纯文本链接</button>
  </div>
  <div class="geo-cite-box">
    <input type="text" class="geo-cite-input" id="geoCiteInput" readonly value="[$cleanTitleNoSuffix]($canonicalUrl)">
    <button class="geo-cite-btn" onclick="copyGeoCite()">
      <svg width="16" height="16" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M8 16H6a2 2 0 01-2-2V6a2 2 0 012-2h8a2 2 0 002-2v-8a2 2 0 00-2-2h-8a2 2 0 00-2 2v8a2 2 0 002 2z"></path></svg>
      复制
    </button>
  </div>
  <div class="geo-cc-banner">
    <svg width="18" height="18" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M9 12l2 2 4-4m5.618-4.016A11.955 11.955 0 0112 2.944a11.955 11.955 0 01-8.618 3.04A12.02 12.02 0 003 9c0 5.591 3.824 10.29 9 11.622 5.176-1.332 9-6.03 9-11.622 0-1.042-.133-2.052-.382-3.016z"></path></svg>
    <span><strong>版权与转载声明：</strong> 本文由云轨导航原创，允许在保留原文 URL 链接及署名的前提下自由转载与引用。</span>
  </div>
</div>
<script>
function switchGeoTab(btn, type) {
  document.querySelectorAll('.geo-cite-tab').forEach(function(t) { t.classList.remove('active'); });
  btn.classList.add('active');
  var input = document.getElementById('geoCiteInput');
  var title = "$cleanTitleNoSuffix";
  var url = "$canonicalUrl";
  if(type === 'md') {
    input.value = '[' + title + '](' + url + ')';
  } else if(type === 'bib') {
    input.value = '@article{yunguidaohang,\n  title={' + title + '},\n  url={' + url + '},\n  publisher={云轨导航}\n}';
  } else {
    input.value = title + ' - ' + url;
  }
}
function copyGeoCite() {
  var input = document.getElementById('geoCiteInput');
  input.select();
  if (navigator.clipboard) {
    navigator.clipboard.writeText(input.value).then(function() {
      alert('引用链接已复制到剪贴板！');
    });
  } else {
    document.execCommand('copy');
    alert('引用链接已复制到剪贴板！');
  }
}
</script>
"@

        if ($content -match '</article>') {
            $content = [System.Text.RegularExpressions.Regex]::Replace($content, '(</article>)', "$geoWidget`n`$1", [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
        } elseif ($content -match '<footer') {
            $content = [System.Text.RegularExpressions.Regex]::Replace($content, '(<footer)', "$geoWidget`n`$1", [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
        }
    }

    # Write updated content back with UTF-8 No BOM
    [System.IO.File]::WriteAllText($file.FullName, $content, $utf8NoBom)
    $updatedCount++
    Write-Host "Processed [$relPath] -> Title: [$cleanTitleNoSuffix] | Meta Desc Len: $($finalDesc.Length) chars."
}

Write-Host "`nSuccessfully updated $updatedCount HTML files with GEO AI optimizations!" -ForegroundColor Green
