#!/bin/bash
# 双击启动赫眉盘点网页版
cd "$(dirname "$0")"
# 后台等服务起来后自动开浏览器
( sleep 2; open "http://127.0.0.1:5001" ) &
echo "正在启动盘点网页…  浏览器会自动打开 http://127.0.0.1:5001"
echo "用完后，关掉这个终端窗口即可停止。"
./.venv/bin/python app.py
