# 文章工具箱 wechat2docx

一个本地运行的内容抓取/转换工具箱，提供网页 Web 界面，主要做三件事：

1. **微信公众号 / CSDN 文章 → Word / Markdown**（保留图片、代码高亮、公式、emoji）
2. **Word(.docx) → Markdown**（图片以 base64 内嵌）
3. **小鹅通(xiaoecloud)等加密 HLS 视频 → MP4 下载**

> ⚠️ 仅供个人学习用途，请只下载/转换你**已登录、已购买或有权访问**的内容，遵守各平台服务条款与版权规定。

---

## 功能与实现概览

### 标签一：URL 转文档（`/convert`）
- 核心库 [`wechat2docx.py`](wechat2docx.py)：`requests` 抓取页面 → `BeautifulSoup` 解析 → `python-docx` 生成 Word。
- 自动识别微信(`mp.weixin.qq.com`) / CSDN(`blog.csdn.net`)，分别用对应解析器提取标题/作者/时间与正文。
- 图片用线程池并发预下载到内存；webp/gif 自动转 PNG。
- 代码块保留语法高亮颜色；LaTeX 公式用 matplotlib mathtext 渲染，MathJax SVG 用 resvg 渲染为图片。
- 选 Markdown 时：先生成临时 docx，再走 `_docx_to_md()` 转换。

### 标签二：Word → Markdown（`/docx2md`）
- 遍历段落，按样式映射标题/引用/代码块；图片提取为 `data:` base64 内嵌。
- 命令行版见 [`doc2md.py`](doc2md.py)，还支持 PDF/PPTX/XLSX/图片/HTML/ZIP（依赖 `markitdown`）。

### 标签三：视频下载（`/api/video/*`）
- 下载内核 [`video_download.py`](video_download.py)，对外入口 `download_m3u8()`：
  - **主路径**：调用系统 `ffmpeg` 直下加密 m3u8（`-headers` 带 UA/Referer/Cookie），`-c copy` 无损合并为 mp4。
  - **兜底路径**：ffmpeg 失败时纯 Python 拉分片 → 解析 `#EXT-X-KEY` → 取 AES-128 key → 逐片 CBC 解密 → 拼接 → 封装。
- 视频是分钟级长任务，后端用**内存任务表 + 后台线程**异步执行，前端**轮询** `GET /api/video/jobs/<id>` 刷新进度条与日志。
- 真实 m3u8 地址由用户在浏览器侧提供（见下文两种方式），后端只负责"下载+解密+合并"，不做任何逆向/破解。

### Web 服务
- [`wechat2docx_app.py`](wechat2docx_app.py)：Flask 应用，内联单页 HTML（三个标签），默认 HTTPS 监听 `5001`。

---

## 环境与安装

- **Python** 3.10+
- **系统依赖**：`ffmpeg`（视频下载/合并必需）
  - macOS：`brew install ffmpeg`
- **Python 依赖**：
  ```bash
  pip install -r requirements.txt
  ```
  含：`flask`、`requests`、`beautifulsoup4`、`python-docx`、`pillow`、`matplotlib`、`resvg-py`、`markitdown`、`pycryptodome`（仅视频 AES 兜底用到）。

---

## 启动

### Web 界面（推荐）
```bash
cd /path/to/wechat2docx
python3 wechat2docx_app.py
```
然后浏览器打开 **https://localhost:5001**

- 默认走 HTTPS：若目录下存在 `wechat2docx_cert.pem` / `wechat2docx_key.pem` 则用之，否则尝试 adhoc，再不行降级 HTTP。
- 自签证书首次访问会有安全警告 → 「高级 → 继续前往 localhost」。

### 命令行

文章转 Word：
```bash
python3 wechat2docx.py "https://mp.weixin.qq.com/s/xxxx" -o out.docx --font 楷体 --font-size 12
```

任意文件转 Markdown：
```bash
python3 doc2md.py file.docx        # 或 .pdf/.pptx/.xlsx/...
```

下载加密 m3u8 视频：
```bash
python3 video_download.py "<真实m3u8地址>" -o downloads -t 文件名 [--cookie "..."] [--referer "..."]
```

---

## 视频下载怎么用

视频文件的"真实 m3u8 地址"需要在你**已登录**的浏览器里获取，有两种方式：

### 方式 A · 手动粘贴（稳，随时可用）
1. 浏览器登录平台、打开视频页并**播放**。
2. `F12` → **Network** → 过滤框输入 `m3u8` → 找到 `.m3u8` 请求 → 右键 **Copy → Copy link address**。
3. 粘进「🎬 视频下载」标签的 m3u8 框 → 「开始下载」。
4. 若失败：展开「高级」，把该请求 Headers 里的 **Cookie / Referer** 也填进去再试。

### 方式 B · 油猴一键脚本（装一次，以后点一下）
脚本：[`browser_helper/xiaoe_capture.user.js`](browser_helper/xiaoe_capture.user.js)

1. 安装浏览器扩展 **Tampermonkey**。
2. 服务启动后，浏览器访问 **https://localhost:5001/xiaoe_capture.user.js**，Tampermonkey 会弹出安装页 → 安装。
3. 打开小鹅通视频页，右下角出现「🎬 抓取已就绪」状态条 → **点击播放** → 状态条变成绿色「⬇️ 发送」按钮 → 点一下 → 自动跳回本工具并开始下载。

> 脚本在播放页（含 iframe）拦截 m3u8 网络请求，连同页面标题/Referer 通过本地新标签页发给工具。`@match` 覆盖 `xiaoecloud.com / xiaoe-tech.com / xiaoeknow.com / xeknow.com`，如遇其他播放器域名需在脚本里补 `@match`。

---

## 已知边界
- 视频下载仅"接住"浏览器里已产生的地址，**不破解付费墙 / 不逆向签名**；签名地址会过期，过期后重新抓取即可（方式 B 现抓现下天然规避）。
- 若某课程用了非标 AES 或 DRM 等更强加密，ffmpeg 与标准 AES 兜底可能失败。
- 雪球文章受 WAF JS Challenge 保护，`requests` 抓不到，暂不支持。

---

## 后台常驻服务（macOS，关掉终端/VSCode 也不停、开机自启）

用 launchd 把服务装成"像后台 Mac 程序"一样常驻。

> ⚠️ 注意：项目**不能放在 `~/Documents`、`~/Desktop`、`~/Downloads`** 这些被 macOS 隐私保护(TCC)的目录，否则后台进程无权访问会报 `Operation not permitted`。本项目放在 `~/wechat2docx`。

- 启动脚本：[`start_service.sh`](start_service.sh)（cd 到项目目录并用 pyenv 3.12.3 的 python 启动）
- LaunchAgent 配置：`~/Library/LaunchAgents/com.wechat2docx.app.plist`（`RunAtLoad` 登录自启 + `KeepAlive` 崩溃自拉起）

**常用管理命令**（`UID` 用 `id -u` 的值，一般是 501）：
```bash
# 启动 / 装入（首次或修改 plist 后）
launchctl bootstrap gui/$(id -u) ~/Library/LaunchAgents/com.wechat2docx.app.plist
# 停止 / 卸载
launchctl bootout   gui/$(id -u)/com.wechat2docx.app
# 立即重启
launchctl kickstart -k gui/$(id -u)/com.wechat2docx.app
# 查看状态
launchctl print gui/$(id -u)/com.wechat2docx.app | grep -E 'state|pid'
# 看日志
tail -f ~/wechat2docx/service.out.log
```
装好后浏览器访问 **https://localhost:5001**（建议加书签）。自签证书首次访问点「高级 → 继续前往」。

---

## 目录结构（核心）
```
wechat2docx/
├── wechat2docx.py            # 文章抓取 + Word 生成 核心库 / CLI
├── wechat2docx_app.py        # Flask Web 界面（三个标签 + 视频任务系统）
├── video_download.py         # 加密 m3u8 下载/解密内核 / CLI
├── doc2md.py                 # 通用文件 → Markdown CLI
├── browser_helper/
│   └── xiaoe_capture.user.js # 油猴一键抓取脚本
├── start_service.sh          # launchd 后台服务启动脚本
├── requirements.txt
└── downloads/                # 视频下载输出（已 gitignore）
```
