#!/usr/bin/env python3
"""
wechat2docx_app.py - 微信文章转 Word 的 Web 界面

启动：
    python3 wechat2docx_app.py
然后打开 http://localhost:5000
"""

import sys
import tempfile
import os
from urllib.parse import quote
from flask import Flask, request, jsonify, render_template_string

# 复用 wechat2docx 里的核心函数
sys.path.insert(0, os.path.dirname(__file__))
from wechat2docx import fetch_article, extract_meta, build_docx

app = Flask(__name__)

HTML = """
<!DOCTYPE html>
<html lang="zh-CN">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>微信文章 → Word</title>
  <style>
    *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }

    body {
      font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
      background: #f0f4f8;
      min-height: 100vh;
      display: flex;
      align-items: center;
      justify-content: center;
      padding: 24px;
    }

    .card {
      background: #fff;
      border-radius: 16px;
      box-shadow: 0 4px 24px rgba(0,0,0,.08);
      padding: 40px 44px;
      width: 100%;
      max-width: 560px;
    }

    .logo {
      display: flex;
      align-items: center;
      gap: 10px;
      margin-bottom: 28px;
    }
    .logo-icon {
      width: 40px; height: 40px;
      background: #07c160;
      border-radius: 10px;
      display: flex; align-items: center; justify-content: center;
    }
    .logo-icon svg { width: 24px; height: 24px; fill: #fff; }
    .logo-text h1 { font-size: 18px; font-weight: 700; color: #1a1a1a; }
    .logo-text p  { font-size: 13px; color: #888; margin-top: 2px; }

    label { display: block; font-size: 14px; font-weight: 600; color: #333; margin-bottom: 8px; }

    .input-row { display: flex; gap: 8px; }

    input[type="url"] {
      flex: 1;
      padding: 12px 14px;
      border: 1.5px solid #d0d7de;
      border-radius: 10px;
      font-size: 14px;
      color: #1a1a1a;
      outline: none;
      transition: border-color .2s;
    }
    input[type="url"]:focus { border-color: #07c160; }
    input[type="url"]::placeholder { color: #bbb; }

    #btn {
      padding: 12px 20px;
      background: #07c160;
      color: #fff;
      border: none;
      border-radius: 10px;
      font-size: 14px;
      font-weight: 600;
      cursor: pointer;
      white-space: nowrap;
      transition: background .2s, transform .1s;
      display: flex; align-items: center; gap: 6px;
      min-width: 110px; justify-content: center;
    }
    #btn:hover  { background: #06ad56; }
    #btn:active { transform: scale(.97); }
    #btn:disabled { background: #9ed6b8; cursor: not-allowed; }
    #btn svg { width: 16px; height: 16px; fill: #fff; flex-shrink: 0; }

    .hint { margin-top: 10px; font-size: 12px; color: #aaa; line-height: 1.6; }

    #msg { display: none; margin-top: 20px; padding: 12px 16px; border-radius: 10px; font-size: 14px; line-height: 1.6; }
    #msg.error   { display: block; background: #fff2f2; color: #c0392b; border: 1px solid #f5c6cb; }
    #msg.success { display: block; background: #f0fdf4; color: #1a7a40; border: 1px solid #bbf7d0; }

    .spinner {
      display: none;
      width: 15px; height: 15px;
      border: 2px solid rgba(255,255,255,.4);
      border-top-color: #fff;
      border-radius: 50%;
      animation: spin .7s linear infinite;
      flex-shrink: 0;
    }
    @keyframes spin { to { transform: rotate(360deg); } }

    .divider { border: none; border-top: 1px solid #eee; margin: 24px 0; }
    .features { display: grid; grid-template-columns: 1fr 1fr; gap: 12px; }
    .feature-item { display: flex; align-items: flex-start; gap: 8px; font-size: 13px; color: #555; }
    .feature-item .dot { width: 6px; height: 6px; border-radius: 50%; background: #07c160; flex-shrink: 0; margin-top: 5px; }
  </style>
</head>
<body>
  <div class="card">
    <div class="logo">
      <div class="logo-icon">
        <svg viewBox="0 0 24 24"><path d="M8.7 10.3a1 1 0 1 0 0-2 1 1 0 0 0 0 2zm4.6 0a1 1 0 1 0 0-2 1 1 0 0 0 0 2zM12 2C6.48 2 2 6.48 2 12s4.48 10 10 10 10-4.48 10-10S17.52 2 12 2zm4.93 12.12c-.36.1-.74.16-1.13.16-.27 0-.53-.03-.79-.07l-1.56.93.26-1.35c-1.23-.72-2.02-1.83-2.02-3.07 0-2.1 2.35-3.8 5.24-3.8s5.24 1.7 5.24 3.8c0 1.43-1.05 2.69-2.63 3.35l-.61-.95zm-9.46.04a5.4 5.4 0 0 1-.79.06c-.39 0-.77-.06-1.13-.16l-.61.95C3.36 14.35 2.31 13.1 2.31 11.66c0-2.1 2.35-3.8 5.24-3.8s5.24 1.7 5.24 3.8c0 1.24-.79 2.35-2.02 3.07l.26 1.35-1.56-.92z"/></svg>
      </div>
      <div class="logo-text">
        <h1>微信文章 → Word</h1>
        <p>粘贴公众号链接，一键下载 docx</p>
      </div>
    </div>

    <label for="url">文章链接</label>
    <div class="input-row">
      <input id="url" type="url" required
        placeholder="https://mp.weixin.qq.com/s/..."
        autocomplete="off">
      <button id="btn" type="button">
        <div class="spinner" id="spinner"></div>
        <svg id="btn-icon" viewBox="0 0 24 24"><path d="M5 20h14v-2H5v2zm7-18L5.33 10h4.34v6h4.66v-6h4.34L12 2z"/></svg>
        <span id="btn-text">转换下载</span>
      </button>
    </div>
    <p class="hint">支持微信公众号文章 · 自动保留正文图片 · 无需登录</p>

    <div id="msg"></div>

    <hr class="divider">

    <div class="features">
      <div class="feature-item"><div class="dot"></div><span>保留正文全部图片</span></div>
      <div class="feature-item"><div class="dot"></div><span>保留标题层级结构</span></div>
      <div class="feature-item"><div class="dot"></div><span>自动提取元信息</span></div>
      <div class="feature-item"><div class="dot"></div><span>输出标准 .docx 格式</span></div>
    </div>
  </div>

  <script>
    const urlInput = document.getElementById('url');
    const btn      = document.getElementById('btn');
    const spinner  = document.getElementById('spinner');
    const btnIcon  = document.getElementById('btn-icon');
    const btnText  = document.getElementById('btn-text');
    const msg      = document.getElementById('msg');

    function showMsg(type, text) {
      msg.className = type;
      msg.textContent = text;
    }

    btn.addEventListener('click', async () => {
      const url = urlInput.value.trim();
      if (!url) { urlInput.focus(); return; }

      // Loading state
      btn.disabled = true;
      spinner.style.display = 'block';
      btnIcon.style.display = 'none';
      msg.className = '';
      msg.textContent = '';

      const t0 = Date.now();
      const timer = setInterval(() => {
        btnText.textContent = `转换中… ${Math.round((Date.now()-t0)/1000)}s`;
      }, 500);

      try {
        const resp = await fetch('/convert', {
          method: 'POST',
          headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
          body: 'url=' + encodeURIComponent(url),
        });

        clearInterval(timer);

        if (!resp.ok) {
          const data = await resp.json().catch(() => ({ error: resp.statusText }));
          showMsg('error', '❌ ' + (data.error || '转换失败'));
          return;
        }

        // 文件下载 — 从响应头取文件名
        const cd = resp.headers.get('Content-Disposition') || '';
        let filename = 'article.docx';
        const m1 = cd.match(/filename\\*=UTF-8''([^;\\s]+)/i);
        const m2 = cd.match(/filename="([^"]+)"/i);
        if (m1) filename = decodeURIComponent(m1[1]);
        else if (m2) filename = m2[1];

        const blob = await resp.blob();
        const a = document.createElement('a');
        a.href = URL.createObjectURL(blob);
        a.download = filename;
        document.body.appendChild(a);
        a.click();
        a.remove();
        URL.revokeObjectURL(a.href);

        showMsg('success', `✅ 已下载：${filename}`);

      } catch (err) {
        clearInterval(timer);
        showMsg('error', '❌ 网络错误：' + err.message);
      } finally {
        btn.disabled = false;
        spinner.style.display = 'none';
        btnIcon.style.display = 'block';
        btnText.textContent = '转换下载';
      }
    });

    // Enter 键触发
    urlInput.addEventListener('keydown', e => { if (e.key === 'Enter') btn.click(); });
  </script>
</body>
</html>
"""


@app.get("/")
def index():
    return render_template_string(HTML)


@app.post("/convert")
def convert():
    url = request.form.get("url", "").strip()

    if not url:
        return jsonify({"error": "请输入文章链接"}), 400

    if "mp.weixin.qq.com" not in url:
        return jsonify({"error": "请输入有效的微信公众号文章链接（mp.weixin.qq.com）"}), 400

    try:
        soup = fetch_article(url)
        meta = extract_meta(soup)

        tmp = tempfile.NamedTemporaryFile(suffix=".docx", delete=False)
        tmp.close()
        build_docx(soup, meta, tmp.name)

        # 用标题做文件名，过滤非法字符
        safe_title = "".join(c for c in meta["title"] if c not in r'\/:*?"<>|').strip() or "article"
        filename = f"{safe_title}.docx"
        encoded_filename = quote(filename, safe='')

        with open(tmp.name, "rb") as f:
            data = f.read()
        os.unlink(tmp.name)

        from flask import Response
        # filename= 只能用 ASCII（latin-1），中文标题放 filename*=UTF-8'' 里
        ascii_fallback = "article.docx"
        return Response(
            data,
            mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            headers={
                "Content-Disposition": (
                    f"attachment; filename=\"{ascii_fallback}\"; "
                    f"filename*=UTF-8''{encoded_filename}"
                )
            },
        )

    except Exception as e:
        error_msg = str(e)
        if "403" in error_msg or "Forbidden" in error_msg:
            error_msg = "微信拒绝访问（403），请稍后重试，或文章已设置付费/仅粉丝可见"
        elif "timeout" in error_msg.lower():
            error_msg = "请求超时，请检查网络或稍后重试"
        return jsonify({"error": f"转换失败：{error_msg}"}), 500


if __name__ == "__main__":
    print("启动服务：http://localhost:5001")
    app.run(host="0.0.0.0", port=5001, debug=False)
