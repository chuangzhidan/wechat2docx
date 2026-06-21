#!/bin/bash
# wechat2docx 后台服务启动脚本（供 launchd 调用）
cd /Users/chuangzhidian/wechat2docx || exit 1
exec /Users/chuangzhidian/.pyenv/versions/3.12.3/bin/python3 wechat2docx_app.py
