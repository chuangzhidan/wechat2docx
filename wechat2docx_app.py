#!/usr/bin/env python3
"""
wechat2docx_app.py - 微信/CSDN 文章转 Word/Markdown，Word 转 Markdown Web 界面

启动：
    python3 wechat2docx_app.py
然后打开 https://localhost:5001
"""

import sys
import tempfile
import os
from urllib.parse import quote
from flask import Flask, request, jsonify, render_template_string

sys.path.insert(0, os.path.dirname(__file__))
from wechat2docx import (
  fetch_article,
  extract_meta,
  build_docx,
  SUPPORTED_BODY_FONTS,
)

app = Flask(__name__)

HTML = """
<!DOCTYPE html>
<html lang="zh-CN">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>文章工具箱</title>
  <style>
    *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }
    body {
      font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
      background: #f0f4f8;
      min-height: 100vh;
      display: flex;
      align-items: flex-start;
      justify-content: center;
      padding: 32px 24px;
    }
    .card {
      background: #fff;
      border-radius: 16px;
      box-shadow: 0 4px 24px rgba(0,0,0,.08);
      padding: 36px 40px;
      width: 100%;
      max-width: 600px;
    }

    /* Logo */
    .logo { display: flex; align-items: center; gap: 10px; margin-bottom: 28px; }
    .logo-icon {
      width: 40px; height: 40px; background: #07c160;
      border-radius: 10px; display: flex; align-items: center; justify-content: center;
    }
    .logo-icon svg { width: 24px; height: 24px; fill: #fff; }
    .logo-text h1 { font-size: 18px; font-weight: 700; color: #1a1a1a; }
    .logo-text p  { font-size: 13px; color: #888; margin-top: 2px; }

    /* Tabs */
    .tabs { display: flex; gap: 0; border-bottom: 2px solid #eee; margin-bottom: 24px; }
    .tab {
      padding: 10px 18px; font-size: 14px; font-weight: 600; color: #888;
      cursor: pointer; border-bottom: 2px solid transparent; margin-bottom: -2px;
      transition: color .2s, border-color .2s;
    }
    .tab.active { color: #07c160; border-bottom-color: #07c160; }
    .tab:hover:not(.active) { color: #555; }

    /* Panels */
    .panel { display: none; }
    .panel.active { display: block; }

    label { display: block; font-size: 14px; font-weight: 600; color: #333; margin-bottom: 8px; }

    /* Format selector */
    .format-row {
      display: flex; gap: 10px; margin-bottom: 14px;
    }
    .fmt-btn {
      flex: 1; padding: 9px 0; border: 1.5px solid #d0d7de; border-radius: 8px;
      font-size: 13px; font-weight: 600; color: #555; cursor: pointer;
      background: #fff; transition: all .2s; text-align: center;
    }
    .fmt-btn.active { border-color: #07c160; background: #f0fdf4; color: #07c160; }
    .fmt-btn:hover:not(.active) { border-color: #aaa; }

    .option-row { display: grid; grid-template-columns: 1fr 1fr; gap: 10px; margin-bottom: 14px; }
    .option-field label { font-size: 13px; margin-bottom: 6px; }
    select {
      width: 100%; padding: 10px 12px; border: 1.5px solid #d0d7de;
      border-radius: 10px; font-size: 14px; color: #1a1a1a;
      background: #fff; outline: none; transition: border-color .2s;
    }
    select:focus { border-color: #07c160; }

    /* Input row */
    .input-row { display: flex; gap: 8px; }
    input[type="url"], input[type="text"] {
      flex: 1; padding: 12px 14px; border: 1.5px solid #d0d7de;
      border-radius: 10px; font-size: 14px; color: #1a1a1a;
      outline: none; transition: border-color .2s;
    }
    input:focus { border-color: #07c160; }
    input::placeholder { color: #bbb; }

    /* File upload */
    .upload-area {
      border: 2px dashed #d0d7de; border-radius: 10px; padding: 28px;
      text-align: center; cursor: pointer; transition: border-color .2s, background .2s;
      margin-bottom: 14px; position: relative;
    }
    .upload-area:hover, .upload-area.drag { border-color: #07c160; background: #f0fdf4; }
    .upload-area input[type="file"] {
      position: absolute; inset: 0; opacity: 0; cursor: pointer; width: 100%; height: 100%;
    }
    .upload-icon { font-size: 32px; margin-bottom: 8px; }
    .upload-area p { font-size: 14px; color: #666; }
    .upload-area .hint { font-size: 12px; color: #aaa; margin-top: 4px; }
    .selected-file { font-size: 13px; color: #07c160; margin-top: 8px; font-weight: 600; }

    /* Buttons */
    .btn-primary {
      width: 100%; padding: 12px 20px; background: #07c160; color: #fff;
      border: none; border-radius: 10px; font-size: 14px; font-weight: 600;
      cursor: pointer; transition: background .2s, transform .1s;
      display: flex; align-items: center; justify-content: center; gap: 8px;
    }
    .btn-primary:hover  { background: #06ad56; }
    .btn-primary:active { transform: scale(.98); }
    .btn-primary:disabled { background: #9ed6b8; cursor: not-allowed; }
    .btn-primary svg { width: 16px; height: 16px; fill: #fff; flex-shrink: 0; }

    .hint { margin-top: 10px; font-size: 12px; color: #aaa; line-height: 1.6; }

    /* Message */
    #msg-url, #msg-upload {
      display: none; margin-top: 16px; padding: 12px 16px;
      border-radius: 10px; font-size: 14px; line-height: 1.6;
    }
    .error   { display: block !important; background: #fff2f2; color: #c0392b; border: 1px solid #f5c6cb; }
    .success { display: block !important; background: #f0fdf4; color: #1a7a40; border: 1px solid #bbf7d0; }

    .spinner {
      display: none; width: 15px; height: 15px;
      border: 2px solid rgba(255,255,255,.4); border-top-color: #fff;
      border-radius: 50%; animation: spin .7s linear infinite; flex-shrink: 0;
    }
    @keyframes spin { to { transform: rotate(360deg); } }

    .divider { border: none; border-top: 1px solid #eee; margin: 24px 0; }
    .features { display: grid; grid-template-columns: 1fr 1fr; gap: 10px; }
    .feature-item { display: flex; align-items: flex-start; gap: 8px; font-size: 13px; color: #555; }
    .feature-item .dot { width: 6px; height: 6px; border-radius: 50%; background: #07c160; flex-shrink: 0; margin-top: 5px; }
  </style>
</head>
<body>
<div class="card">

  <!-- Logo -->
  <div class="logo">
    <div class="logo-icon">
      <svg viewBox="0 0 24 24"><path d="M8.7 10.3a1 1 0 1 0 0-2 1 1 0 0 0 0 2zm4.6 0a1 1 0 1 0 0-2 1 1 0 0 0 0 2zM12 2C6.48 2 2 6.48 2 12s4.48 10 10 10 10-4.48 10-10S17.52 2 12 2z"/></svg>
    </div>
    <div class="logo-text">
      <h1>文章工具箱</h1>
      <p>URL 抓取 · Word 转 Markdown · 格式自由选择</p>
    </div>
  </div>

  <!-- Tabs -->
  <div class="tabs">
    <div class="tab active" data-tab="url">🌐 URL 转文档</div>
    <div class="tab" data-tab="upload">📄 Word → Markdown</div>
  </div>

  <!-- Panel 1: URL 转 Word / MD -->
  <div class="panel active" id="panel-url">
    <label>目标格式</label>
    <div class="format-row">
      <div class="fmt-btn active" data-fmt="docx">📝 Word (.docx)</div>
      <div class="fmt-btn" data-fmt="md">🗒️ Markdown (.md)</div>
    </div>

    <div class="option-row" id="docx-options">
      <div class="option-field">
        <label for="font-select">导出字体</label>
        <select id="font-select">
          <option value="宋体-简" selected>宋体-简（默认）</option>
          <option value="宋体">宋体</option>
          <option value="Times New Roman">Times New Roman</option>
          <option value="楷体">楷体</option>
        </select>
      </div>
      <div class="option-field">
        <label for="font-size-select">正文字号</label>
        <select id="font-size-select">
          <option value="" selected>默认（10.5 pt）</option>
          <option value="9">9 pt</option>
          <option value="10.5">10.5 pt</option>
          <option value="12">12 pt</option>
          <option value="14">14 pt</option>
        </select>
      </div>
    </div>

    <label for="url-input">文章链接</label>
    <div class="input-row">
      <input id="url-input" type="url" required
        placeholder="https://mp.weixin.qq.com/s/... 或 https://blog.csdn.net/..."
        autocomplete="off">
      <button class="btn-primary" id="url-btn" type="button" style="width:auto;min-width:110px;">
        <div class="spinner" id="url-spinner"></div>
        <svg id="url-icon" viewBox="0 0 24 24"><path d="M5 20h14v-2H5v2zm7-18L5.33 10h4.34v6h4.66v-6h4.34L12 2z"/></svg>
        <span id="url-btn-text">转换下载</span>
      </button>
    </div>
    <p class="hint">支持微信公众号 & CSDN 博客 · 自动保留图片 · 图片合集可提取文字 · 字体/字号可选 · 雪球需浏览器验证暂不支持</p>
    <div id="msg-url"></div>

    <hr class="divider">
    <div class="features">
      <div class="feature-item"><div class="dot"></div><span>保留正文全部图片</span></div>
      <div class="feature-item"><div class="dot"></div><span>保留标题层级结构</span></div>
      <div class="feature-item"><div class="dot"></div><span>自动提取元信息</span></div>
      <div class="feature-item"><div class="dot"></div><span>输出 docx / md 可选</span></div>
    </div>
  </div>

  <!-- Panel 2: Word 上传 → Markdown -->
  <div class="panel" id="panel-upload">
    <label>上传 Word 文件</label>
    <div class="upload-area" id="drop-zone">
      <input type="file" id="file-input" accept=".docx">
      <div class="upload-icon">📄</div>
      <p>点击选择或拖拽 .docx 文件</p>
      <div class="hint">转换为 Markdown（图片嵌入 base64）</div>
      <div class="selected-file" id="selected-name"></div>
    </div>
    <button class="btn-primary" id="upload-btn" type="button">
      <div class="spinner" id="upload-spinner"></div>
      <svg id="upload-icon" viewBox="0 0 24 24"><path d="M5 20h14v-2H5v2zm7-18L5.33 10h4.34v6h4.66v-6h4.34L12 2z"/></svg>
      <span id="upload-btn-text">转换为 Markdown</span>
    </button>
    <div id="msg-upload"></div>
  </div>

</div>

<script>
// ——— Tab 切换 ———
document.querySelectorAll('.tab').forEach(tab => {
  tab.addEventListener('click', () => {
    document.querySelectorAll('.tab').forEach(t => t.classList.remove('active'));
    document.querySelectorAll('.panel').forEach(p => p.classList.remove('active'));
    tab.classList.add('active');
    document.getElementById('panel-' + tab.dataset.tab).classList.add('active');
  });
});

// ——— 格式按钮 ———
let selectedFmt = 'docx';
const docxOptions = document.getElementById('docx-options');
document.querySelectorAll('.fmt-btn').forEach(btn => {
  btn.addEventListener('click', () => {
    document.querySelectorAll('.fmt-btn').forEach(b => b.classList.remove('active'));
    btn.classList.add('active');
    selectedFmt = btn.dataset.fmt;
    docxOptions.style.display = selectedFmt === 'docx' ? 'grid' : 'none';
  });
});

// ——— URL 转换 ———
const urlInput  = document.getElementById('url-input');
const urlBtn    = document.getElementById('url-btn');
const urlSpinner = document.getElementById('url-spinner');
const urlIcon   = document.getElementById('url-icon');
const urlBtnTxt = document.getElementById('url-btn-text');
const msgUrl    = document.getElementById('msg-url');
const fontSelect = document.getElementById('font-select');
const fontSizeSelect = document.getElementById('font-size-select');

function showMsg(el, type, text) {
  el.className = type;
  el.textContent = text;
}

urlBtn.addEventListener('click', async () => {
  const url = urlInput.value.trim();
  if (!url) { urlInput.focus(); return; }

  urlBtn.disabled = true;
  urlSpinner.style.display = 'block';
  urlIcon.style.display = 'none';
  msgUrl.className = ''; msgUrl.textContent = '';

  const t0 = Date.now();
  const timer = setInterval(() => {
    urlBtnTxt.textContent = '转换中… ' + Math.round((Date.now()-t0)/1000) + 's';
  }, 500);

  try {
    const params = new URLSearchParams();
    params.set('url', url);
    params.set('fmt', selectedFmt);
    if (selectedFmt === 'docx') {
      params.set('font', fontSelect.value);
      if (fontSizeSelect.value) params.set('font_size', fontSizeSelect.value);
    }
    const resp = await fetch('/convert', {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: params.toString(),
    });
    clearInterval(timer);

    if (!resp.ok) {
      const data = await resp.json().catch(() => ({ error: resp.statusText }));
      showMsg(msgUrl, 'error', '❌ ' + (data.error || '转换失败'));
      return;
    }

    const cd = resp.headers.get('Content-Disposition') || '';
    let filename = selectedFmt === 'md' ? 'article.md' : 'article.docx';
    const m1 = cd.match(/filename\\*=UTF-8''([^;\\s]+)/i);
    const m2 = cd.match(/filename="([^"]+)"/i);
    if (m1) filename = decodeURIComponent(m1[1]);
    else if (m2) filename = m2[1];

    const blob = await resp.blob();
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob); a.download = filename;
    document.body.appendChild(a); a.click(); a.remove();
    URL.revokeObjectURL(a.href);
    showMsg(msgUrl, 'success', '✅ 已下载：' + filename);

  } catch(err) {
    clearInterval(timer);
    showMsg(msgUrl, 'error', '❌ 网络错误：' + err.message);
  } finally {
    urlBtn.disabled = false;
    urlSpinner.style.display = 'none';
    urlIcon.style.display = 'block';
    urlBtnTxt.textContent = '转换下载';
  }
});
urlInput.addEventListener('keydown', e => { if (e.key === 'Enter') urlBtn.click(); });

// ——— 剪贴板自动检测 ———
let _lastClipUrl = '';
let _clipToast = null;

function _dismissToast() {
  if (_clipToast) { _clipToast.remove(); _clipToast = null; }
}

function _showClipToast(url) {
  _dismissToast();
  const toast = document.createElement('div');
  toast.style.cssText = [
    'position:fixed','top:16px','left:50%','transform:translateX(-50%)',
    'background:#1a1a1a','color:#fff','border-radius:12px',
    'padding:10px 16px','font-size:13px','display:flex',
    'align-items:center','gap:10px','z-index:9999',
    'box-shadow:0 4px 20px rgba(0,0,0,.25)','max-width:90vw',
  ].join(';');
  const shortUrl = url.length > 48 ? url.slice(0, 45) + '…' : url;
  toast.innerHTML = [
    '<span>📋 检测到链接：<b>' + shortUrl + '</b></span>',
    '<button id="toast-use" style="background:#07c160;color:#fff;border:none;border-radius:6px;padding:4px 12px;cursor:pointer;font-size:13px;white-space:nowrap;">使用</button>',
    '<button id="toast-dismiss" style="background:transparent;color:#aaa;border:none;cursor:pointer;font-size:18px;line-height:1;padding:0 2px;">✕</button>',
  ].join('');
  document.body.appendChild(toast);
  _clipToast = toast;
  toast.querySelector('#toast-use').addEventListener('click', () => {
    urlInput.value = url;
    _lastClipUrl = url;
    _dismissToast();
    // 切换到 URL Tab
    document.querySelectorAll('.tab').forEach(t => t.classList.remove('active'));
    document.querySelectorAll('.panel').forEach(p => p.classList.remove('active'));
    document.querySelector('.tab[data-tab="url"]').classList.add('active');
    document.getElementById('panel-url').classList.add('active');
  });
  toast.querySelector('#toast-dismiss').addEventListener('click', () => {
    _lastClipUrl = url;
    _dismissToast();
  });
  // 10s 自动消失
  setTimeout(_dismissToast, 10000);
}

async function _checkClipboard() {
  try {
    const text = (await navigator.clipboard.readText()).trim();
    if (!text || text === _lastClipUrl) return;
    const isArticle = text.includes('mp.weixin.qq.com') || text.includes('blog.csdn.net');
    if ((text.startsWith('http://') || text.startsWith('https://')) && isArticle) {
      if (text !== urlInput.value.trim()) _showClipToast(text);
    }
  } catch (_) { /* 权限未授予或浏览器限制，静默失败 */ }
}

window.addEventListener('focus', _checkClipboard);
document.addEventListener('visibilitychange', () => { if (!document.hidden) _checkClipboard(); });
// 页面加载后也检查一次
document.addEventListener('DOMContentLoaded', () => setTimeout(_checkClipboard, 500));

// ——— 文件上传 ———
const fileInput    = document.getElementById('file-input');
const dropZone     = document.getElementById('drop-zone');
const selectedName = document.getElementById('selected-name');
const uploadBtn    = document.getElementById('upload-btn');
const uploadSpinner = document.getElementById('upload-spinner');
const uploadIcon   = document.getElementById('upload-icon');
const uploadBtnTxt = document.getElementById('upload-btn-text');
const msgUpload    = document.getElementById('msg-upload');

fileInput.addEventListener('change', () => {
  selectedName.textContent = fileInput.files[0] ? '已选择：' + fileInput.files[0].name : '';
});

['dragover','dragenter'].forEach(ev => {
  dropZone.addEventListener(ev, e => { e.preventDefault(); dropZone.classList.add('drag'); });
});
['dragleave','drop'].forEach(ev => {
  dropZone.addEventListener(ev, e => {
    e.preventDefault(); dropZone.classList.remove('drag');
    if (ev === 'drop' && e.dataTransfer.files.length) {
      fileInput.files = e.dataTransfer.files;
      selectedName.textContent = '已选择：' + fileInput.files[0].name;
    }
  });
});

uploadBtn.addEventListener('click', async () => {
  if (!fileInput.files.length) {
    showMsg(msgUpload, 'error', '❌ 请先选择 .docx 文件');
    return;
  }
  uploadBtn.disabled = true;
  uploadSpinner.style.display = 'block';
  uploadIcon.style.display = 'none';
  msgUpload.className = ''; msgUpload.textContent = '';

  const t0 = Date.now();
  const timer = setInterval(() => {
    uploadBtnTxt.textContent = '转换中… ' + Math.round((Date.now()-t0)/1000) + 's';
  }, 500);

  try {
    const fd = new FormData();
    fd.append('file', fileInput.files[0]);
    const resp = await fetch('/docx2md', { method: 'POST', body: fd });
    clearInterval(timer);

    if (!resp.ok) {
      const data = await resp.json().catch(() => ({ error: resp.statusText }));
      showMsg(msgUpload, 'error', '❌ ' + (data.error || '转换失败'));
      return;
    }

    const cd = resp.headers.get('Content-Disposition') || '';
    let filename = 'output.md';
    const m1 = cd.match(/filename\\*=UTF-8''([^;\\s]+)/i);
    const m2 = cd.match(/filename="([^"]+)"/i);
    if (m1) filename = decodeURIComponent(m1[1]);
    else if (m2) filename = m2[1];

    const blob = await resp.blob();
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob); a.download = filename;
    document.body.appendChild(a); a.click(); a.remove();
    URL.revokeObjectURL(a.href);
    showMsg(msgUpload, 'success', '✅ 已下载：' + filename);

  } catch(err) {
    clearInterval(timer);
    showMsg(msgUpload, 'error', '❌ 网络错误：' + err.message);
  } finally {
    uploadBtn.disabled = false;
    uploadSpinner.style.display = 'none';
    uploadIcon.style.display = 'block';
    uploadBtnTxt.textContent = '转换为 Markdown';
  }
});
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
    fmt = request.form.get("fmt", "docx").strip()  # docx 或 md
    body_font = request.form.get("font", "").strip() or None
    body_font_size = request.form.get("font_size", "").strip() or None

    if not url:
        return jsonify({"error": "请输入文章链接"}), 400

    if body_font and body_font not in SUPPORTED_BODY_FONTS:
        return jsonify({"error": "不支持的导出字体"}), 400

    if body_font_size:
        try:
            size_value = float(body_font_size)
        except ValueError:
            return jsonify({"error": "字号必须是数字"}), 400
        if not 6 <= size_value <= 24:
            return jsonify({"error": "字号必须在 6 到 24 pt 之间"}), 400

    is_wechat = "mp.weixin.qq.com" in url
    is_csdn   = "blog.csdn.net" in url
    if not (is_wechat or is_csdn):
        return jsonify({"error": "请输入有效的文章链接（支持微信公众号 mp.weixin.qq.com 或 CSDN blog.csdn.net）"}), 400

    try:
        soup = fetch_article(url)
        meta = extract_meta(soup)
        safe_title = "".join(c for c in meta["title"] if c not in r'\/:*?"<>|').strip() or "article"

        from flask import Response

        if fmt == "md":
            # URL → docx（临时），再 docx → md
            tmp_docx = tempfile.NamedTemporaryFile(suffix=".docx", delete=False)
            tmp_docx.close()
            build_docx(soup, meta, tmp_docx.name, body_font=body_font, body_font_size=body_font_size)
            md_text = _docx_to_md(tmp_docx.name)
            os.unlink(tmp_docx.name)

            filename = safe_title + ".md"
            encoded = quote(filename, safe='')
            return Response(
                md_text.encode("utf-8"),
                mimetype="text/markdown; charset=utf-8",
                headers={
                    "Content-Disposition": (
                        f'attachment; filename="article.md"; filename*=UTF-8\'\'{encoded}'
                    )
                },
            )
        else:
            # URL → docx
            tmp = tempfile.NamedTemporaryFile(suffix=".docx", delete=False)
            tmp.close()
            build_docx(soup, meta, tmp.name, body_font=body_font, body_font_size=body_font_size)

            filename = safe_title + ".docx"
            encoded = quote(filename, safe='')
            with open(tmp.name, "rb") as f:
                data = f.read()
            os.unlink(tmp.name)

            return Response(
                data,
                mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                headers={
                    "Content-Disposition": (
                        f'attachment; filename="article.docx"; filename*=UTF-8\'\'{encoded}'
                    )
                },
            )

    except Exception as e:
        error_msg = str(e)
        if "403" in error_msg or "Forbidden" in error_msg:
            error_msg = "微信拒绝访问（403），请稍后重试，或文章已付费/仅粉丝可见"
        elif "timeout" in error_msg.lower():
            error_msg = "请求超时，请检查网络或稍后重试"
        return jsonify({"error": f"转换失败：{error_msg}"}), 500


@app.post("/docx2md")
def docx2md():
    if "file" not in request.files:
        return jsonify({"error": "请上传 .docx 文件"}), 400

    f = request.files["file"]
    if not f.filename.lower().endswith(".docx"):
        return jsonify({"error": "只支持 .docx 格式"}), 400

    tmp = tempfile.NamedTemporaryFile(suffix=".docx", delete=False)
    try:
        f.save(tmp.name)
        tmp.close()
        md_text = _docx_to_md(tmp.name)
    except Exception as e:
        return jsonify({"error": f"转换失败：{e}"}), 500
    finally:
        os.unlink(tmp.name)

    safe_stem = "".join(c for c in os.path.splitext(f.filename)[0] if c not in r'\/:*?"<>|').strip() or "output"
    filename = safe_stem + ".md"
    encoded = quote(filename, safe='')

    from flask import Response
    return Response(
        md_text.encode("utf-8"),
        mimetype="text/markdown; charset=utf-8",
        headers={
            "Content-Disposition": (
                f'attachment; filename="output.md"; filename*=UTF-8\'\'{encoded}'
            )
        },
    )


def _docx_to_md(docx_path: str) -> str:
    """将 .docx 转为 Markdown（图片嵌入 base64），复用 doc2md.py 的逻辑"""
    import base64
    from docx import Document as DocxDoc
    from docx.oxml.ns import qn

    doc = DocxDoc(docx_path)

    # rId → base64 data URI
    image_map = {}
    for rId, rel in doc.part.rels.items():
        if "image" in rel.reltype:
            blob = rel.target_part.blob
            ctype = rel.target_part.content_type
            b64 = base64.b64encode(blob).decode()
            image_map[rId] = f"data:{ctype};base64,{b64}"

    img_counter = [0]

    def para_to_md(para) -> str:
        result = []
        for child in para._p:
            tag = child.tag.split("}")[-1] if "}" in child.tag else child.tag
            if tag == "r":
                for blip in child.findall(".//" + qn("a:blip")):
                    rId = blip.get(
                        "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed"
                    )
                    if rId and rId in image_map:
                        img_counter[0] += 1
                        result.append(f"\n![图{img_counter[0]}]({image_map[rId]})\n")
                for t in child.findall(qn("w:t")):
                    if t.text:
                        result.append(t.text)
            elif tag == "hyperlink":
                # 提取超链接文本（如有 URL 关系则加链接）
                rId = child.get(qn("r:id"), "")
                href = doc.part.rels.get(rId)
                href_url = href.target_ref if href else None
                texts = [t.text for t in child.findall(".//" + qn("w:t")) if t.text]
                link_text = "".join(texts)
                if href_url and href_url.startswith("http"):
                    result.append(f"[{link_text}]({href_url})")
                else:
                    result.append(link_text)
        return "".join(result)

    lines = []
    for para in doc.paragraphs:
        style = para.style.name
        text = para_to_md(para)
        if not text.strip():
            continue
        if style.startswith("Heading 1") or style == "Title":
            lines.append(f"# {text}")
        elif style.startswith("Heading 2"):
            lines.append(f"## {text}")
        elif style.startswith("Heading 3"):
            lines.append(f"### {text}")
        elif style == "No Spacing":
            # 代码块：用 ``` 包裹
            lines.append(f"```\n{text}\n```")
        elif style in ("Intense Quote", "Quote"):
            lines.append(f"> {text}")
        else:
            lines.append(text)

    print(f"  [docx→md] {img_counter[0]} 张图片，{len(lines)} 个段落", file=sys.stderr)
    return "\n\n".join(lines)


if __name__ == "__main__":
    _dir = os.path.dirname(os.path.abspath(__file__))
    _cert = os.path.join(_dir, "wechat2docx_cert.pem")
    _key  = os.path.join(_dir, "wechat2docx_key.pem")
    if os.path.exists(_cert) and os.path.exists(_key):
        ssl_ctx = (_cert, _key)
    else:
        try:
            import OpenSSL  # noqa: F401
            ssl_ctx = "adhoc"
        except ImportError:
            ssl_ctx = None
    scheme = "https" if ssl_ctx else "http"
    print(f"启动服务：{scheme}://localhost:5001")
    app.run(host="0.0.0.0", port=5001, debug=False, ssl_context=ssl_ctx)

