from __future__ import annotations

import contextlib
import io
import json
import os
import subprocess
import sys
import threading
import webbrowser
from http.server import BaseHTTPRequestHandler, HTTPServer
from pathlib import Path

from main import build_reimbursement_pdf

HTML_PAGE = """<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>费用报销单自动生成工具</title>
<style>
  * { box-sizing: border-box; }
  body {
    font-family: -apple-system, BlinkMacSystemFont, "PingFang SC", "Helvetica Neue", sans-serif;
    background: #f4f6f9;
    color: #1f2937;
    margin: 0;
    padding: 32px 16px;
  }
  .card {
    max-width: 720px;
    margin: 0 auto;
    background: #fff;
    border-radius: 14px;
    box-shadow: 0 4px 20px rgba(0, 0, 0, 0.06);
    padding: 28px 32px;
  }
  h1 { margin: 0 0 6px; font-size: 22px; }
  .sub { color: #6b7280; font-size: 13px; margin-bottom: 24px; }
  .field { margin-bottom: 18px; }
  label { display: block; font-size: 14px; font-weight: 600; margin-bottom: 6px; }
  .row { display: flex; gap: 10px; }
  input[type="text"] {
    flex: 1;
    padding: 10px 12px;
    border: 1px solid #d1d5db;
    border-radius: 8px;
    font-size: 14px;
  }
  button {
    padding: 10px 16px;
    border: none;
    border-radius: 8px;
    font-size: 14px;
    font-weight: 600;
    cursor: pointer;
    transition: opacity .15s;
  }
  button:hover { opacity: 0.9; }
  button:disabled { opacity: 0.5; cursor: not-allowed; }
  .btn-secondary { background: #e5e7eb; color: #111827; }
  .btn-primary { background: #2563eb; color: #fff; width: 100%; padding: 12px; font-size: 15px; }
  .btn-success { background: #059669; color: #fff; }
  .actions { display: none; gap: 10px; margin-top: 14px; }
  .log-box {
    margin-top: 20px;
    background: #111827;
    color: #e5e7eb;
    border-radius: 10px;
    padding: 14px 16px;
    font-family: ui-monospace, SFMono-Regular, Menlo, monospace;
    font-size: 12.5px;
    line-height: 1.6;
    white-space: pre-wrap;
    max-height: 340px;
    overflow-y: auto;
    display: none;
  }
  .ok { color: #34d399; }
  .warn { color: #fbbf24; }
  .err { color: #f87171; }
  .tips {
    margin-top: 18px;
    padding: 12px 14px;
    background: #f9fafb;
    border-radius: 8px;
    font-size: 12.5px;
    color: #4b5563;
    line-height: 1.6;
  }
</style>
</head>
<body>
<div class="card">
  <h1>费用报销单自动生成工具</h1>
  <div class="sub">选择报销材料文件夹，自动识别订单/发票/行程单、核对合规并生成 PDF</div>

  <div class="field">
    <label>报销材料文件夹</label>
    <div class="row">
      <input type="text" id="dirPath" placeholder="点击右侧按钮选择文件夹，或粘贴路径">
      <button class="btn-secondary" type="button" onclick="pickFolder()">选择文件夹</button>
    </div>
  </div>

  <div class="field">
    <label>报销部门</label>
    <input type="text" id="dept" value="Bida" style="width: 100%;">
  </div>

  <button class="btn-primary" id="runBtn" onclick="generate()">开始生成报销单 PDF</button>

  <div class="actions" id="actions">
    <button class="btn-success" onclick="openPdf()">打开生成的 PDF</button>
    <button class="btn-secondary" onclick="revealPdf()">在访达中显示</button>
  </div>

  <div class="log-box" id="logBox"></div>

  <div class="tips">
    <b>文件夹结构要求：</b><br>
    • <code>signatrue.jpg</code>（签名图片）<br>
    • <code>orderImage/</code>（餐饮订单截图，需 17:30 后下单且 14 天内）<br>
    • <code>发票/</code>（广州餐饮电子发票，支持 PDF 或图片格式，自动扣除配送费）<br>
    • <code>traffic/</code>（可选：滴滴行程单与交通发票，发票支持 PDF 或图片格式，需 23:00 后上车）
  </div>
</div>

<script>
let lastPdfPath = "";

async function pickFolder() {
  const res = await fetch("/api/pick-folder", { method: "POST" });
  const data = await res.json();
  if (data.path) {
    document.getElementById("dirPath").value = data.path;
  }
}

function escapeHtml(s) {
  return s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

function renderLog(text) {
  const lines = text.split("\\n").map(line => {
    const esc = escapeHtml(line);
    if (line.includes("[通过]") || line.includes("全部合规校验通过")) return `<span class="ok">${esc}</span>`;
    if (line.includes("[警告]") || line.startsWith("  !")) return `<span class="warn">${esc}</span>`;
    if (line.includes("[不合规]") || line.includes("错误")) return `<span class="err">${esc}</span>`;
    return esc;
  });
  const box = document.getElementById("logBox");
  box.style.display = "block";
  box.innerHTML = lines.join("\\n");
}

async function generate() {
  const dirPath = document.getElementById("dirPath").value.trim();
  const dept = document.getElementById("dept").value.trim() || "Bida";
  if (!dirPath) {
    alert("请先选择或输入报销材料文件夹路径");
    return;
  }
  const btn = document.getElementById("runBtn");
  btn.disabled = true;
  btn.textContent = "正在识别与核对中，请稍候...";
  document.getElementById("actions").style.display = "none";

  try {
    const res = await fetch("/api/generate", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ dir: dirPath, dept })
    });
    const data = await res.json();
    renderLog(data.log || data.error || "");
    if (data.ok && data.output_pdf) {
      lastPdfPath = data.output_pdf;
      document.getElementById("actions").style.display = "flex";
    }
  } catch (e) {
    renderLog("错误: " + e);
  } finally {
    btn.disabled = false;
    btn.textContent = "开始生成报销单 PDF";
  }
}

async function openPdf() {
  if (!lastPdfPath) return;
  await fetch("/api/open", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ path: lastPdfPath, reveal: false })
  });
}

async function revealPdf() {
  if (!lastPdfPath) return;
  await fetch("/api/open", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ path: lastPdfPath, reveal: true })
  });
}
</script>
</body>
</html>
"""


class ReimburseHandler(BaseHTTPRequestHandler):
    def log_message(self, format, *args):
        pass

    def _send_json(self, data: dict, status: int = 200) -> None:
        body = json.dumps(data, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def do_GET(self):
        body = HTML_PAGE.encode("utf-8")
        self.send_response(200)
        self.send_header("Content-Type", "text/html; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def do_POST(self):
        length = int(self.headers.get("Content-Length", "0"))
        raw = self.rfile.read(length) if length else b"{}"
        payload = json.loads(raw.decode("utf-8"))

        if self.path == "/api/pick-folder":
            if sys.platform == "win32":
                ps_cmd = (
                    "Add-Type -AssemblyName System.Windows.Forms;"
                    "$f = New-Object System.Windows.Forms.FolderBrowserDialog;"
                    "if ($f.ShowDialog() -eq 'OK') { $f.SelectedPath }"
                )
                res = subprocess.run(
                    ["powershell", "-NoProfile", "-Command", ps_cmd],
                    capture_output=True,
                    text=True,
                    check=False,
                )
                folder = res.stdout.strip()
            else:
                res = subprocess.run(
                    ["osascript", "-e", 'POSIX path of (choose folder with prompt "请选择报销材料文件夹")'],
                    capture_output=True,
                    text=True,
                    check=False,
                )
                folder = res.stdout.strip().rstrip("/")
            self._send_json({"path": folder})
            return

        if self.path == "/api/generate":
            target_dir = Path(payload.get("dir", "")).expanduser()
            dept = payload.get("dept", "Bida")
            output_pdf = target_dir.parent / f"{target_dir.name}_生成.pdf"
            buf = io.StringIO()
            try:
                with contextlib.redirect_stdout(buf):
                    build_reimbursement_pdf(target_dir, output_pdf, dept=dept)
                self._send_json({"ok": True, "output_pdf": str(output_pdf), "log": buf.getvalue()})
            except Exception as exc:
                self._send_json({"ok": False, "error": str(exc), "log": f"{buf.getvalue()}\n错误: {exc}"}, status=500)
            return

        if self.path == "/api/open":
            pdf_path = payload.get("path", "")
            reveal = bool(payload.get("reveal", False))
            if sys.platform == "win32":
                if reveal:
                    subprocess.Popen(["explorer", "/select,", str(Path(pdf_path))])
                else:
                    os.startfile(pdf_path)
            else:
                cmd = ["open", "-R", pdf_path] if reveal else ["open", pdf_path]
                subprocess.Popen(cmd)
            self._send_json({"ok": True})
            return

        self._send_json({"error": "Not Found"}, status=404)


def main() -> None:
    server = HTTPServer(("127.0.0.1", 0), ReimburseHandler)
    port = server.server_address[1]
    url = f"http://127.0.0.1:{port}/"
    print(f"报销单生成工具已启动: {url}")
    print("请在浏览器中操作，关闭此窗口即可退出。")
    threading.Timer(0.5, lambda: webbrowser.open(url)).start()
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        pass
    finally:
        server.server_close()


if __name__ == "__main__":
    main()
