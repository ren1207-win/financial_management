#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
财务数据看板启动脚本
使用 Python 内置 HTTP 服务器提供静态文件服务
"""

import http.server
import socketserver
import os
import webbrowser
import sys

# 服务器配置
PORT = 8000
HOST = "localhost"
DIRECTORY = os.path.dirname(os.path.abspath(__file__))


def main():
    # 切换到项目目录
    os.chdir(DIRECTORY)

    # 创建 HTTP 服务器
    handler = http.server.SimpleHTTPRequestHandler

    with socketserver.TCPServer((HOST, PORT), handler) as httpd:
        url = f"http://{HOST}:{PORT}/前端页面框架.html"

        print("=" * 50)
        print("     财务数据看板 - 本地服务器已启动")
        print("=" * 50)
        print()
        print(f"服务器地址: http://{HOST}:{PORT}")
        print(f"访问地址: {url}")
        print()
        print("按 Ctrl+C 停止服务器")
        print("=" * 50)
        print()

        # 自动打开浏览器
        try:
            webbrowser.open(url)
        except Exception as e:
            print(f"自动打开浏览器失败: {e}")

        # 开始服务
        try:
            httpd.serve_forever()
        except KeyboardInterrupt:
            print()
            print("\n服务器已停止")
            sys.exit(0)


if __name__ == "__main__":
    main()