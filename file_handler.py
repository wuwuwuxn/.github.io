#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
简易文件上传与分析服务
- POST /upload: 接收Excel文件并保存，调用处理脚本生成 analysis_results.json
- 静态资源: 继承 SimpleHTTPRequestHandler 提供 index.html / upload.html / analysis_results.json 静态访问
"""

import os
import sys
import json
import http.server
import socketserver
import urllib.parse
import cgi
import io
import subprocess
from pathlib import Path
from datetime import datetime

class FileUploadHandler(http.server.SimpleHTTPRequestHandler):
    # 允许跨域（简单处理）
    def end_headers(self):
        self.send_header('Access-Control-Allow-Origin', '*')
        super().end_headers()

    def do_OPTIONS(self):
        self.send_response(204)
        # 仅在此声明允许的方法与头，Origin 由 end_headers 统一添加以避免重复
        self.send_header('Access-Control-Allow-Methods', 'POST, GET, OPTIONS')
        self.send_header('Access-Control-Allow-Headers', 'Content-Type')
        self.end_headers()

    def do_POST(self):
        if self.path == '/upload':
            self.handle_file_upload()
        else:
            self.send_error(404, "Unknown POST endpoint")

    def do_GET(self):
        if self.path == '/history':
            try:
                history_dir = Path.cwd() / 'history'
                items = []
                if history_dir.exists():
                    files = sorted(history_dir.glob('*.json'), key=lambda p: p.stat().st_mtime, reverse=True)
                    for p in files:
                        ts = datetime.fromtimestamp(p.stat().st_mtime).strftime('%Y-%m-%d %H:%M:%S')
                        items.append({
                            'name': p.name,
                            'url': f'/history/{p.name}',
                            'timestamp': ts
                        })
                self.send_response(200)
                self.send_header('Content-Type', 'application/json')
                self.end_headers()
                self.wfile.write(json.dumps(items, ensure_ascii=False).encode('utf-8'))
            except Exception as e:
                self.send_response(500)
                self.send_header('Content-Type', 'application/json')
                self.end_headers()
                self.wfile.write(json.dumps({'error': str(e)}, ensure_ascii=False).encode('utf-8'))
        else:
            return super().do_GET()

    def handle_file_upload(self):
        try:
            content_type = self.headers.get('Content-Type')
            if not content_type or not content_type.startswith('multipart/form-data'):
                self.send_error(400, "Invalid content type")
                return

            content_length = int(self.headers.get('Content-Length', 0))
            post_data = self.rfile.read(content_length)
            fs = cgi.FieldStorage(
                fp=io.BytesIO(post_data),
                headers=self.headers,
                environ={'REQUEST_METHOD': 'POST', 'CONTENT_TYPE': content_type}
            )

            if 'file' not in fs:
                self.send_error(400, "Missing file field")
                return

            file_item = fs['file']
            filename = file_item.filename or 'uploaded_file.xlsx'
            data = file_item.file.read()

            # 保存到当前工作目录
            save_dir = Path.cwd()
            save_path = save_dir / filename
            with open(save_path, 'wb') as f:
                f.write(data)

            # 调用分析脚本，生成 analysis_results.json
            cmd = [sys.executable, '处理数据.py', str(save_path)]
            proc = subprocess.run(cmd, capture_output=True, text=True)

            if proc.returncode != 0:
                resp = {
                    'success': False,
                    'message': '分析失败',
                    'stderr': proc.stderr,
                    'stdout': proc.stdout
                }
                self.send_response(500)
                self.send_header('Content-Type', 'application/json')
                self.end_headers()
                self.wfile.write(json.dumps(resp).encode('utf-8'))
                return

            # 读取生成的结果摘要（可选）
            result_path = save_dir / 'analysis_results.json'
            summary = {}
            if result_path.exists():
                try:
                    with open(result_path, 'r', encoding='utf-8') as f:
                        data_json = json.load(f)
                    summary = data_json.get('data_summary', {})
                except Exception:
                    summary = {}

            # 保存历史记录文件：上传表格名(去扩展名) + 分析时间
            history_dir = save_dir / 'history'
            history_dir.mkdir(exist_ok=True)
            base_name = Path(filename).stem
            timestamp = datetime.now().strftime('%Y%m%d-%H%M%S')
            history_filename = f"{base_name}_{timestamp}.json"
            history_path = history_dir / history_filename
            try:
                if result_path.exists():
                    with open(result_path, 'rb') as rf, open(history_path, 'wb') as hf:
                        hf.write(rf.read())
            except Exception:
                pass

            response = {
                'success': True,
                'message': '文件上传并分析完成',
                'filename': filename,
                'size': len(data),
                'summary': summary,
                'history_file': f'/history/{history_filename}',
                'history_timestamp': timestamp
            }

            self.send_response(200)
            self.send_header('Content-Type', 'application/json')
            self.end_headers()
            self.wfile.write(json.dumps(response, ensure_ascii=False).encode('utf-8'))
        except Exception as e:
            response = {
                'success': False,
                'message': f'上传或分析失败: {str(e)}'
            }
            self.send_response(500)
            self.send_header('Content-Type', 'application/json')
            self.end_headers()
            self.wfile.write(json.dumps(response, ensure_ascii=False).encode('utf-8'))

if __name__ == '__main__':
    PORT = 8000
    # 允许通过命令行指定端口，如: python file_handler.py 8001
    if len(sys.argv) > 1:
        try:
            PORT = int(sys.argv[1])
        except Exception:
            pass
    with socketserver.TCPServer(('', PORT), FileUploadHandler) as httpd:
        print(f"Serving at http://localhost:{PORT}/")
        httpd.serve_forever()