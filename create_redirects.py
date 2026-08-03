import os

redirects = {
    'airport-recommend': 'airports.html',
    'airport-rank': 'ranking.html',
    'reviews': 'knowledge.html',
    'tutorials': 'knowledge.html',
    'clash-tutorial': 'articles/tutorial-win.html',
    'shadowrocket-tutorial': 'articles/tutorial-ios.html',
    'client-download': 'articles/nav-1.html',
    'avoid-scam': 'avoid-scam.html'
}

html_template = """<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta http-equiv="refresh" content="0;url=../{url}">
    <title>跳转中...</title>
    <script>
        window.location.replace("../{url}");
    </script>
</head>
<body>
    <p>正在为您跳转到相关页面，请稍候... <a href="../{url}">点击这里直接跳转</a></p>
</body>
</html>"""

base_dir = r"c:\Users\Administrator\Desktop\博客"

for folder, target_url in redirects.items():
    folder_path = os.path.join(base_dir, folder)
    if not os.path.exists(folder_path):
        os.makedirs(folder_path)
    with open(os.path.join(folder_path, 'index.html'), 'w', encoding='utf-8') as f:
        f.write(html_template.format(url=target_url))
    print(f'Created clean relative redirect for {folder} -> ../{target_url}')
