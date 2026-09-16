#!/usr/bin/env python3
"""Serve the project-report download page on 0.0.0.0:8080."""
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
import os

ROOT = Path("/workspace/public")
PORT = 8080
DOCX_NAME = "Shorya_Agarwal_MDPS_Project_Report.docx"
DOWNLOAD_AS = "Shorya_Agarwal_Multiple_Disease_Prediction_Report.docx"


class Handler(SimpleHTTPRequestHandler):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=str(ROOT), **kwargs)

    def log_message(self, fmt, *args):
        sys_stderr = __import__("sys").stderr
        sys_stderr.write("%s - %s\n" % (self.address_string(), fmt % args))

    def end_headers(self):
        self.send_header("Cache-Control", "no-store")
        super().end_headers()

    def do_GET(self):
        path = self.path.split("?", 1)[0]
        if path in ("/download", "/download.docx", "/report.docx"):
            self.path = "/" + DOCX_NAME
        return super().do_GET()

    def guess_type(self, path):
        if str(path).endswith(".docx"):
            return "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        return super().guess_type(path)

    def send_head(self):
        # Force a download filename for the Word file
        rel = self.path.split("?", 1)[0].lstrip("/")
        if rel == DOCX_NAME or self.path.startswith("/download"):
            file_path = ROOT / DOCX_NAME
            if file_path.is_file():
                try:
                    f = open(file_path, "rb")
                except OSError:
                    self.send_error(404, "File not found")
                    return None
                self.send_response(200)
                self.send_header(
                    "Content-Type",
                    "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                )
                self.send_header(
                    "Content-Disposition",
                    f'attachment; filename="{DOWNLOAD_AS}"',
                )
                self.send_header("Content-Length", str(file_path.stat().st_size))
                self.end_headers()
                return f
        return super().send_head()


if __name__ == "__main__":
    os.chdir(ROOT)
    httpd = ThreadingHTTPServer(("0.0.0.0", PORT), Handler)
    print(f"Download server listening on 0.0.0.0:{PORT}", flush=True)
    httpd.serve_forever()
