"""Local web app for formatting Word documents. No third-party packages needed."""
from __future__ import annotations

import json
import mimetypes
import os
import secrets
import threading
import time
from email.parser import BytesParser
from email import policy
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import parse_qs, urlparse

from formatter import FormatError, MAX_FILE, format_docx, inspect, load_templates


ROOT = Path(__file__).parent
STATIC = ROOT / "static"
JOBS = {}
JOBS_LOCK = threading.Lock()
JOB_TTL = 15 * 60
MAX_JOBS = 20


def prune_jobs() -> None:
    now = time.time()
    with JOBS_LOCK:
        for key in list(JOBS):
            if JOBS[key][0] < now:
                del JOBS[key]


def multipart_docx(body: bytes, content_type: str) -> bytes:
    if "multipart/form-data" not in content_type or "boundary=" not in content_type:
        raise FormatError("请使用文件上传接口")
    raw = (f"Content-Type: {content_type}\r\nMIME-Version: 1.0\r\n\r\n").encode() + body
    message = BytesParser(policy=policy.default).parsebytes(raw)
    if not message.is_multipart():
        raise FormatError("上传数据格式错误")
    for part in message.iter_parts():
        if part.get_param("name", header="content-disposition") == "file":
            filename = part.get_filename() or ""
            if not filename.lower().endswith(".docx"):
                raise FormatError("仅支持 .docx 文件")
            data = part.get_payload(decode=True)
            if not data or len(data) > 9 * 1024 * 1024:
                raise FormatError("小程序文件须小于 9 MB")
            return data
    raise FormatError("未找到上传的文件")


class Handler(BaseHTTPRequestHandler):
    def send_data(self, status: int, body: bytes, content_type: str, extra=None) -> None:
        self.send_response(status)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        self.send_header("X-Content-Type-Options", "nosniff")
        for key, value in (extra or {}).items():
            self.send_header(key, value)
        self.end_headers()
        self.wfile.write(body)

    def json(self, status: int, obj) -> None:
        self.send_data(status, json.dumps(obj, ensure_ascii=False).encode("utf-8"), "application/json; charset=utf-8")

    def do_GET(self) -> None:
        path = urlparse(self.path).path
        if path.startswith("/api/jobs/") and path.endswith("/download"):
            job_id = path.split("/")[3]
            prune_jobs()
            with JOBS_LOCK:
                job = JOBS.get(job_id)
            if not job:
                self.json(404, {"error": "下载链接已过期"})
                return
            self.send_data(200, job[1], "application/vnd.openxmlformats-officedocument.wordprocessingml.document", {
                "Content-Disposition": 'attachment; filename="formatted.docx"'
            })
            return
        if path == "/api/templates":
            self.json(200, [{k: t[k] for k in ("id", "name", "version", "basis", "source_url")} for t in load_templates()])
            return
        if path == "/":
            path = "/index.html"
        allowed = {"/index.html", "/app.css", "/app.js"}
        if path not in allowed:
            self.json(404, {"error": "页面不存在"})
            return
        file = STATIC / path.lstrip("/")
        body = file.read_bytes()
        mime = mimetypes.guess_type(str(file))[0] or "application/octet-stream"
        self.send_data(200, body, f"{mime}; charset=utf-8")

    def do_POST(self) -> None:
        parsed = urlparse(self.path)
        if parsed.path not in ("/api/analyze", "/api/format", "/api/jobs"):
            self.json(404, {"error": "接口不存在"})
            return
        try:
            length = int(self.headers.get("Content-Length", "0"))
            if length <= 0 or length > MAX_FILE:
                raise FormatError("请选择不超过 15 MB 的 .docx 文件")
            data = self.rfile.read(length)
            template = parse_qs(parsed.query).get("template", ["gov_official"])[0]
            if parsed.path == "/api/jobs":
                data = multipart_docx(data, self.headers.get("Content-Type", ""))
                result, report = format_docx(data, template)
                prune_jobs()
                job_id = secrets.token_urlsafe(32)
                with JOBS_LOCK:
                    if len(JOBS) >= MAX_JOBS:
                        oldest = min(JOBS, key=lambda key: JOBS[key][0])
                        del JOBS[oldest]
                    JOBS[job_id] = (time.time() + JOB_TTL, result)
                self.json(200, {"job_id": job_id, "download_path": f"/api/jobs/{job_id}/download", "report": report, "expires_in": JOB_TTL})
            elif parsed.path == "/api/analyze":
                self.json(200, inspect(data, template))
            else:
                result, _ = format_docx(data, template)
                self.send_data(200, result, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", {
                    "Content-Disposition": 'attachment; filename="formatted.docx"'
                })
        except FormatError as exc:
            self.json(400, {"error": str(exc)})
        except Exception:
            self.json(500, {"error": "处理失败，请检查文档是否损坏"})


def main() -> None:
    host = os.environ.get("HOST", "127.0.0.1")
    port = int(os.environ.get("PORT", "8765"))
    server = ThreadingHTTPServer((host, port), Handler)
    print(f"公文格式助手已启动：http://{host}:{port}", flush=True)
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        pass
    finally:
        server.server_close()


if __name__ == "__main__":
    main()
