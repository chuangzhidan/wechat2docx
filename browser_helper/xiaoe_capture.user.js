// ==UserScript==
// @name         小鹅通视频一键抓取 → 本地下载
// @namespace    wechat2docx.video
// @version      0.2.2
// @description  在小鹅通播放页拦截真实 m3u8 地址，一键发送到本地 wechat2docx 工具下载为 mp4。支持 iframe 播放器与白标域名。仅用于下载你已登录/已购的内容。
// @author       wechat2docx
// @match        *://*.xiaoecloud.com/*
// @match        *://*.xiaoe-tech.com/*
// @match        *://*.xiaoeknow.com/*
// @match        *://*.xeknow.com/*
// @match        *://*.citv.cn/*
// @match        *://*.xet.tech/*
// @match        *://*.xet-pc.com/*
// @run-at       document-start
// @grant        GM_openInTab
// ==/UserScript==

(function () {
  'use strict';

  // 本地 wechat2docx 服务地址（默认 https + 5001，与 wechat2docx_app.py 一致）
  const LOCAL_APP = 'https://localhost:5001/';
  const MSG_TAG = '__x2v_m3u8__';

  const isTop = (window.top === window.self);

  // ——————— 抓取逻辑（所有 frame 都跑）———————
  function reportUrl(url) {
    try {
      if (!url || typeof url !== 'string' || url.indexOf('.m3u8') === -1) return;
      const abs = new URL(url, location.href).href;
      if (isTop) {
        onCaptured(abs, document.title);
      } else {
        // 子 frame：把捕获结果发给顶层页面
        window.top.postMessage({ tag: MSG_TAG, url: abs, title: document.title }, '*');
      }
    } catch (_) { /* 忽略非法 URL / 跨域限制 */ }
  }

  // Hook XMLHttpRequest
  const _open = XMLHttpRequest.prototype.open;
  XMLHttpRequest.prototype.open = function (method, url) {
    reportUrl(url);
    return _open.apply(this, arguments);
  };
  // Hook fetch
  const _fetch = window.fetch;
  if (_fetch) {
    window.fetch = function (input) {
      try { reportUrl(typeof input === 'string' ? input : (input && input.url)); } catch (_) {}
      return _fetch.apply(this, arguments);
    };
  }

  // 兜底：定时扫描已加载资源（即使 m3u8 在脚本 hook 之前就请求过也能抓到）
  let scanCount = 0;
  const scanTimer = setInterval(() => {
    try {
      const entries = performance.getEntriesByType('resource') || [];
      for (const e of entries) if (e.name && e.name.indexOf('.m3u8') !== -1) reportUrl(e.name);
      // 同时扫一下 video/source 标签的 src
      document.querySelectorAll('video[src],source[src]').forEach(v => reportUrl(v.src));
    } catch (_) {}
    if (++scanCount > 40) clearInterval(scanTimer);  // 约 60 秒后停止扫描
  }, 1500);

  // ——————— 顶层页面：UI + 收集 ———————
  if (!isTop) return;  // 子 frame 到此为止，只负责抓取并上报

  const captured = [];
  const seen = new Set();
  let box = null;

  window.addEventListener('message', (e) => {
    const d = e.data;
    if (d && d.tag === MSG_TAG && d.url) onCaptured(d.url, d.title);
  });

  function cleanTitle(raw) {
    let t = (raw || document.title || 'video').trim();
    t = t.replace(/[_\-|–—]+\s*小鹅通.*$/, '').trim();
    return t || 'video';
  }

  function onCaptured(url, title) {
    if (seen.has(url)) return;
    seen.add(url);
    captured.push({ url, title: cleanTitle(title) });
    render();
  }

  function sendToLocal(item) {
    const params = new URLSearchParams();
    params.set('m3u8', item.url);
    params.set('title', item.title);
    params.set('referer', location.href);
    const target = LOCAL_APP + '?' + params.toString();
    if (typeof GM_openInTab === 'function') GM_openInTab(target, { active: true, insert: true });
    else window.open(target, '_blank');
  }

  function ensureBox() {
    if (box) return;
    box = document.createElement('div');
    box.style.cssText = [
      'position:fixed', 'right:16px', 'bottom:16px', 'z-index:2147483647',
      'background:#fff', 'border:1px solid #e2e8f0', 'border-radius:12px',
      'box-shadow:0 6px 24px rgba(0,0,0,.18)', 'padding:11px 13px',
      'font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",sans-serif',
      'font-size:13px', 'color:#1a1a1a', 'max-width:300px',
    ].join(';');
    document.body.appendChild(box);
  }

  function render() {
    ensureBox();
    if (captured.length === 0) {
      box.innerHTML =
        '<div style="font-weight:700;">🎬 小鹅通抓取已就绪</div>' +
        '<div style="margin-top:6px;font-size:12px;color:#64748b;">请点击播放视频，捕获到地址后这里会出现下载按钮</div>';
      return;
    }
    box.innerHTML = '<div style="font-weight:700;margin-bottom:8px;">🎬 已捕获 ' + captured.length + ' 个视频流</div><div id="x2v-list"></div>';
    const list = box.querySelector('#x2v-list');
    captured.forEach((item, i) => {
      const btn = document.createElement('button');
      btn.textContent = '⬇️ 发送 #' + (i + 1) + (i === captured.length - 1 ? '（最新）' : '');
      btn.style.cssText = [
        'display:block', 'width:100%', 'margin-top:6px', 'padding:8px 10px',
        'border:0', 'border-radius:8px', 'background:#07c160', 'color:#fff',
        'font-weight:600', 'font-size:12px', 'cursor:pointer', 'text-align:left',
      ].join(';');
      btn.addEventListener('click', () => sendToLocal(item));
      list.appendChild(btn);
    });
  }

  // 进页面就显示"就绪"状态条（看得到=脚本在运行）
  if (document.body) render();
  else document.addEventListener('DOMContentLoaded', render, { once: true });

  // 心跳：状态条被 SPA 页面刷掉时自动重建
  setInterval(() => {
    if (document.body && (!box || !document.body.contains(box))) { box = null; render(); }
  }, 2000);
})();
