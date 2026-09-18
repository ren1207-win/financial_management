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
import socket

# 服务器配置
PORT = 8000
HOST = "127.0.0.1"
DIRECTORY = os.path.dirname(os.path.abspath(__file__))

def main():
    # 设置 socket.getfqdn 的默认值，避免 UnicodeDecodeError
    # 临时绕过 Python 内置函数的编码问题
    def workaround_getfqdn(name=""):
        try:
            return socket.getfqdn(name)
        except UnicodeDecodeError:
            return "127.0.0.1"

    # 保存原函数并替换
    original_getfqdn = socket.getfqdn
    socket.getfqdn = workaround_getfqdn

    # 切换到项目目录
    os.chdir(DIRECTORY)

    try:
        # 创建 HTTP 服务器
        handler = http.server.SimpleHTTPRequestHandler

        # 设置允许重用地址
        socketserver.TCPServer.allow_reuse_address = True

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
                print(f"服务器正在启动，监听 {HOST}:{PORT}...")
                httpd.serve_forever()
            except KeyboardInterrupt:
                print()
                print("\n服务器已停止")
                sys.exit(0)

    except Exception as e:
        # 还原原始函数
        socket.getfqdn = original_getfqdn
        print(f"服务器启动失败: {e}")
        print("解决方案：")
        print("1. 请使用管理员权限运行命令提示符")
        print("2. 或者尝试在系统环境变量中设置：")
        print("   set PYTHONIOENCODING=utf-8")
        print("3. 如果问题仍存在，请尝试修改系统主机名为英文")
        sys.exit(1)
    finally:
        # 确保还原原始函数
        socket.getfqdn = original_getfqdn


if __name__ == "__main__":
    main()