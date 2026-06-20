#!/usr/bin/env python3
"""
video_download.py - 加密 HLS(m3u8) 视频下载内核

用于下载小鹅通(xiaoecloud)等平台的 AES 加密 m3u8 视频流并合并为 mp4。
真实 m3u8 地址 / cookie / referer 由调用方提供（通常在用户已登录的浏览器里抓到），
本模块只负责"下载 + 解密 + 合并"，不做任何逆向 / 破解。

两条路径：
  1) download_m3u8_ffmpeg : 首选，调用系统 ffmpeg 直下加密 m3u8（-headers 带 cookie/referer）
  2) download_m3u8_python : 兜底，纯 Python 拉分片 + AES-128-CBC 解密 + 拼接（需 pycryptodome）

对外主入口 download_m3u8() 会先试 ffmpeg，失败再自动回退到 Python 路径。
"""

from __future__ import annotations

import os
import re
import shutil
import subprocess
import time
from pathlib import Path
from typing import Callable, Optional
from urllib.parse import urljoin, urlparse

import requests

DEFAULT_UA = (
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
)

# on_progress 回调签名：(percent: float|None, stage: str, log: str|None) -> None
ProgressCb = Callable[[Optional[float], str, Optional[str]], None]


def _noop(percent: Optional[float], stage: str, log: Optional[str]) -> None:  # pragma: no cover
    pass


def sanitize_filename(title: str, default: str = "video") -> str:
    """清洗成安全文件名（去非法字符，限长）。"""
    name = re.sub(r"[^\w\s\-_.（）()【】一-鿿]", "_", title or "").strip(" ._")
    name = re.sub(r"\s+", "_", name)
    return name[:120] or default


def unique_file_path(directory: Path, filename: str) -> Path:
    """若文件已存在则加 -2 / -3 ... 后缀，返回不冲突的路径。"""
    directory.mkdir(parents=True, exist_ok=True)
    stem = Path(filename).stem
    suffix = Path(filename).suffix
    candidate = directory / filename
    i = 2
    while candidate.exists():
        candidate = directory / f"{stem}-{i}{suffix}"
        i += 1
    return candidate


def ffmpeg_available() -> bool:
    return shutil.which("ffmpeg") is not None


def _build_header_string(cookie: str = "", referer: str = "", ua: str = "") -> str:
    """构造 ffmpeg -headers 用的多行 header 串（\\r\\n 分隔）。"""
    lines = [f"User-Agent: {ua or DEFAULT_UA}"]
    if referer:
        lines.append(f"Referer: {referer}")
    if cookie:
        lines.append(f"Cookie: {cookie}")
    return "\r\n".join(lines) + "\r\n"


def _parse_duration_seconds(text: bytes) -> Optional[float]:
    m = re.search(rb"Duration:\s*(\d+):(\d+):(\d+)", text)
    if not m:
        return None
    return int(m[1]) * 3600 + int(m[2]) * 60 + int(m[3])


def download_m3u8_ffmpeg(
    m3u8_url: str,
    out_path: Path,
    *,
    cookie: str = "",
    referer: str = "",
    ua: str = "",
    on_progress: ProgressCb = _noop,
) -> Path:
    """用 ffmpeg 直接下载加密 m3u8 流并合并为 mp4。

    ffmpeg 的 -headers 会带到 key 请求，因此 AES-128 且 key 可访问时通常一步到位。
    """
    if not ffmpeg_available():
        raise RuntimeError("未找到 ffmpeg，请先安装（mac: brew install ffmpeg）")

    out_path.parent.mkdir(parents=True, exist_ok=True)
    headers = _build_header_string(cookie, referer, ua)
    cmd = [
        "ffmpeg", "-nostdin", "-y",
        "-reconnect", "1",
        "-reconnect_streamed", "1",
        "-reconnect_delay_max", "15",
        "-headers", headers,
        "-i", m3u8_url,
        "-c", "copy",
        "-bsf:a", "aac_adtstoasc",
        "-movflags", "+faststart",
        str(out_path),
    ]

    on_progress(5, "ffmpeg 下载中（CDN 可能较慢）", "使用 ffmpeg 直接下载加密 m3u8 流")

    process = subprocess.Popen(
        cmd, stdin=subprocess.DEVNULL,
        stdout=subprocess.DEVNULL, stderr=subprocess.PIPE,
    )

    time_re = re.compile(rb"time=(\d+):(\d+):(\d+)")
    total_seconds: Optional[float] = None
    last_update = 0.0
    buf = b""
    tail = b""  # 累积用于提取 Duration（出现在最前面的几行）

    assert process.stderr is not None
    while True:
        chunk = process.stderr.read(512)
        if not chunk:
            break
        if total_seconds is None:
            tail = (tail + chunk)[-4096:]
            total_seconds = _parse_duration_seconds(tail)
        buf += chunk
        parts = buf.split(b"\r")  # ffmpeg 用 \r 刷新进度行
        buf = parts[-1]
        for part in parts[:-1]:
            m = time_re.search(part)
            if m and time.time() - last_update > 1.5:
                cur = int(m[1]) * 3600 + int(m[2]) * 60 + int(m[3])
                if total_seconds and total_seconds > 0:
                    pct = min(99.0, 5 + cur / total_seconds * 93)
                    on_progress(round(pct, 1), f"下载中 {cur // 60}:{cur % 60:02d}", None)
                else:
                    on_progress(None, f"下载中 {cur // 60}:{cur % 60:02d}", None)
                last_update = time.time()

    process.wait()
    if process.returncode != 0 or not out_path.exists() or out_path.stat().st_size < 1024:
        out_path.unlink(missing_ok=True)
        raise RuntimeError("ffmpeg 下载失败（可能 key 需要特殊处理 / CDN 不稳 / 链接已失效）")

    on_progress(100, "视频已保存", f"本地文件：{out_path.resolve()}")
    return out_path


# ——————————————————— Python 兜底路径 ———————————————————

_KEY_ATTR_RE = re.compile(r'([A-Z0-9-]+)=("[^"]*"|[^,]*)')


def _parse_key_line(line: str) -> dict:
    """解析 #EXT-X-KEY 行，返回 {METHOD, URI, IV}。"""
    attrs = {}
    for m in _KEY_ATTR_RE.finditer(line.split(":", 1)[1] if ":" in line else line):
        attrs[m.group(1)] = m.group(2).strip('"')
    return attrs


def _pick_variant(playlist_text: str, base_url: str) -> Optional[str]:
    """若是 master playlist（含多清晰度），挑带宽最高的子 playlist 绝对地址。"""
    if "#EXT-X-STREAM-INF" not in playlist_text:
        return None
    best_bw = -1
    best_url = None
    lines = playlist_text.splitlines()
    for i, line in enumerate(lines):
        if line.startswith("#EXT-X-STREAM-INF"):
            bw_m = re.search(r"BANDWIDTH=(\d+)", line)
            bw = int(bw_m.group(1)) if bw_m else 0
            # 下一非注释行是子 playlist 地址
            for nxt in lines[i + 1:]:
                if nxt and not nxt.startswith("#"):
                    if bw > best_bw:
                        best_bw = bw
                        best_url = urljoin(base_url, nxt.strip())
                    break
    return best_url


def download_m3u8_python(
    m3u8_url: str,
    out_path: Path,
    *,
    cookie: str = "",
    referer: str = "",
    ua: str = "",
    on_progress: ProgressCb = _noop,
) -> Path:
    """纯 Python 下载 + AES-128 解密 + 拼接。用于 ffmpeg 直下失败时兜底。

    生成 .ts（VLC 可直接播放）；若 ffmpeg 可用再无损 remux 成 .mp4。
    """
    try:
        from Crypto.Cipher import AES  # pycryptodome
    except ImportError as exc:  # pragma: no cover
        raise RuntimeError(
            "兜底解密需要 pycryptodome：pip install pycryptodome"
        ) from exc

    session = requests.Session()
    headers = {"User-Agent": ua or DEFAULT_UA}
    if referer:
        headers["Referer"] = referer
    if cookie:
        headers["Cookie"] = cookie
    session.headers.update(headers)

    on_progress(2, "解析 m3u8", "Python 兜底：拉取 m3u8 播放列表")

    def _get(url: str) -> requests.Response:
        r = session.get(url, timeout=30)
        r.raise_for_status()
        return r

    text = _get(m3u8_url).text
    base_url = m3u8_url
    variant = _pick_variant(text, base_url)
    if variant:
        on_progress(3, "解析 m3u8", f"选取最高清晰度子列表")
        base_url = variant
        text = _get(variant).text

    # 解析分片与加密信息
    segments: list[str] = []
    key_attrs: Optional[dict] = None
    key_bytes: Optional[bytes] = None
    iv_bytes: Optional[bytes] = None
    media_seq = 0
    for line in text.splitlines():
        line = line.strip()
        if line.startswith("#EXT-X-MEDIA-SEQUENCE"):
            try:
                media_seq = int(line.split(":", 1)[1])
            except (ValueError, IndexError):
                pass
        elif line.startswith("#EXT-X-KEY"):
            key_attrs = _parse_key_line(line)
        elif line and not line.startswith("#"):
            segments.append(urljoin(base_url, line))

    if not segments:
        raise RuntimeError("m3u8 中未找到任何分片（可能不是有效的媒体播放列表）")

    method = (key_attrs or {}).get("METHOD", "NONE").upper()
    if key_attrs and method not in ("NONE", "AES-128"):
        raise RuntimeError(f"暂不支持的加密方式：{method}（ffmpeg 路径已失败）")

    if key_attrs and method == "AES-128":
        key_uri = urljoin(base_url, key_attrs["URI"])
        key_bytes = _get(key_uri).content
        if len(key_bytes) != 16:
            raise RuntimeError(f"AES key 长度异常（{len(key_bytes)} 字节，应为 16）")
        iv_hex = key_attrs.get("IV")
        if iv_hex:
            iv_bytes = bytes.fromhex(iv_hex.lower().removeprefix("0x"))
        on_progress(5, "已取得解密密钥", "成功获取 AES-128 key")

    out_path.parent.mkdir(parents=True, exist_ok=True)
    ts_path = out_path.with_suffix(".ts")
    total = len(segments)
    with open(ts_path, "wb") as fout:
        for idx, seg_url in enumerate(segments):
            data = _get(seg_url).content
            if key_bytes is not None:
                iv = iv_bytes if iv_bytes is not None else (media_seq + idx).to_bytes(16, "big")
                data = AES.new(key_bytes, AES.MODE_CBC, iv).decrypt(data)
            fout.write(data)
            pct = round(5 + (idx + 1) / total * 90, 1)
            if idx % 3 == 0 or idx == total - 1:
                on_progress(pct, f"下载分片 {idx + 1}/{total}", None)

    # 尝试 remux 成 mp4（无损），失败就保留 .ts
    if ffmpeg_available():
        on_progress(97, "封装 mp4", "ffmpeg 无损封装为 mp4")
        proc = subprocess.run(
            ["ffmpeg", "-nostdin", "-y", "-i", str(ts_path),
             "-c", "copy", "-movflags", "+faststart", str(out_path)],
            stdin=subprocess.DEVNULL, stdout=subprocess.DEVNULL, stderr=subprocess.PIPE,
        )
        if proc.returncode == 0 and out_path.exists() and out_path.stat().st_size > 1024:
            ts_path.unlink(missing_ok=True)
            on_progress(100, "视频已保存", f"本地文件：{out_path.resolve()}")
            return out_path

    # 没有 ffmpeg 或封装失败：直接用 .ts
    final = out_path.with_suffix(".ts")
    if final != ts_path:
        ts_path.replace(final)
    on_progress(100, "视频已保存", f"本地文件（.ts 可用 VLC 播放）：{final.resolve()}")
    return final


def download_m3u8(
    m3u8_url: str,
    output_dir: Path,
    *,
    title: str = "",
    cookie: str = "",
    referer: str = "",
    ua: str = "",
    on_progress: ProgressCb = _noop,
) -> Path:
    """对外主入口：先试 ffmpeg，失败自动回退 Python 解密路径。"""
    if not m3u8_url or ".m3u8" not in urlparse(m3u8_url).path.lower() + m3u8_url.lower():
        # 宽松校验：允许 query 里带 m3u8 的情况
        if "m3u8" not in m3u8_url.lower():
            raise RuntimeError("地址里没有 m3u8，请确认粘贴的是 .m3u8 播放列表地址")

    out_path = unique_file_path(output_dir, f"{sanitize_filename(title)}.mp4")

    try:
        return download_m3u8_ffmpeg(
            m3u8_url, out_path,
            cookie=cookie, referer=referer, ua=ua, on_progress=on_progress,
        )
    except Exception as exc:
        on_progress(None, "ffmpeg 失败，改用兜底解密", f"ffmpeg 路径失败：{exc}")
        return download_m3u8_python(
            m3u8_url, out_path,
            cookie=cookie, referer=referer, ua=ua, on_progress=on_progress,
        )


if __name__ == "__main__":
    import argparse

    ap = argparse.ArgumentParser(description="下载加密 m3u8 视频为 mp4")
    ap.add_argument("m3u8", help="真实 m3u8 地址")
    ap.add_argument("-o", "--out-dir", default="downloads", help="输出目录")
    ap.add_argument("-t", "--title", default="video", help="文件名")
    ap.add_argument("--cookie", default="", help="Cookie 头")
    ap.add_argument("--referer", default="", help="Referer 头")
    args = ap.parse_args()

    def _cli_progress(pct, stage, log):
        bar = f"{pct:5.1f}%" if pct is not None else "  ... "
        line = f"[{bar}] {stage}"
        if log:
            line += f"  | {log}"
        print(line)

    result = download_m3u8(
        args.m3u8, Path(args.out_dir),
        title=args.title, cookie=args.cookie, referer=args.referer,
        on_progress=_cli_progress,
    )
    print(f"\n完成：{result}")
