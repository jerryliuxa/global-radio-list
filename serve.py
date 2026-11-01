# -*- coding: utf-8 -*-
"""本地开发服务器：带 no-cache 头，避免浏览器启发式缓存旧数据。

用法: python serve.py [端口]   （默认 8420，在 radio-guide 目录下运行）
"""
import http.server
import sys

class NoCacheHandler(http.server.SimpleHTTPRequestHandler):
    def end_headers(self):
        self.send_header('Cache-Control', 'no-cache, must-revalidate')
        super().end_headers()

if __name__ == '__main__':
    port = int(sys.argv[1]) if len(sys.argv) > 1 else 8420
    http.server.ThreadingHTTPServer(('127.0.0.1', port), NoCacheHandler).serve_forever()
