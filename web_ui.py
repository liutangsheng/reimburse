from __future__ import annotations

import contextlib
import io
import json
import os
import subprocess
import sys
import threading
import webbrowser
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
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
      <button class="btn-secondary" id="pickBtn" type="button" onclick="pickFolder()">选择文件夹</button>
    </div>
  </div>

  <div class="field">
    <label>报销部门</label>
    <input type="text" id="dept" value="Bida" style="width: 100%;">
  </div>

  <div class="field">
    <label style="display:flex;align-items:center;gap:8px;font-weight:500;cursor:pointer;">
      <input type="checkbox" id="exportDocx" style="width:16px;height:16px;">
      同步输出 Word（.docx，方便人工二次修改）
    </label>
  </div>

  <button class="btn-primary" id="runBtn" onclick="generate()">开始生成报销单 PDF</button>

  <div class="actions" id="actions">
    <button class="btn-success" onclick="openPdf()">打开生成的 PDF</button>
    <button class="btn-secondary" id="openDocxBtn" style="display:none;" onclick="openDocx()">打开生成的 Word</button>
    <button class="btn-secondary" onclick="revealPdf()">在文件夹中显示</button>
  </div>

  <div class="log-box" id="logBox"></div>

  <div class="tips">
    <b>怎么用（3 步搞定）：</b><br>
    1. <b>放好材料</b>：把要报销的订单和发票（PDF 或截图都可以）分别放进对应类目（如 <code>餐费</code>、<code>交通费</code>）的 <code>订单</code> 和 <code>发票</code> 文件夹里，根目录放好签名图片。<br>
    2. <b>自动核对填表</b>：在上方选中你的报销文件夹，工具会自动识别订单和发票内容、检查时间/金额是否合规（比如餐饮 17:30 后、14 天内、自动扣配送费，取订单和发票的较小金额），并自动填好报销单和金额汇总。<br>
    3. <b>一键出 PDF</b>：自动把「报销单 + 所有订单 + 所有发票」按顺序合并成一个完整的电子报销 PDF，直接就能提交。<br><br>
    <b>文件夹结构要求：</b><br>
    • <code>signatrue.jpg</code>（报销人签名图片）<br>
    • <code>&lt;报销类目&gt;/</code>（如 <code>餐费/</code>、<code>交通费/</code>、<code>住宿费/</code>，每个子目录对应一个报销选项）：<br>
    &nbsp;&nbsp;├─ <code>订单/</code>（订单截图或订单/行程单 PDF）<br>
    &nbsp;&nbsp;└─ <code>发票/</code>（电子发票，支持 PDF 或图片格式）
  </div>
</div>

<script>
let lastPdfPath = "";
let lastDocxPath = "";

async function pickFolder() {
  const btn = document.getElementById("pickBtn");
  btn.disabled = true;
  btn.textContent = "选择中...";
  try {
    const res = await fetch("/api/pick-folder", { method: "POST" });
    const data = await res.json();
    if (data.error) {
      alert("打开文件夹选择器失败: " + data.error);
    } else if (data.path) {
      document.getElementById("dirPath").value = data.path;
    }
  } catch (e) {
    alert("打开文件夹选择器失败: " + e);
  } finally {
    btn.disabled = false;
    btn.textContent = "选择文件夹";
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
  const exportDocx = document.getElementById("exportDocx").checked;
  if (!dirPath) {
    alert("请先选择或输入报销材料文件夹路径");
    return;
  }
  const btn = document.getElementById("runBtn");
  btn.disabled = true;
  btn.textContent = "正在识别与核对中，请稍候...";
  document.getElementById("actions").style.display = "none";
  document.getElementById("openDocxBtn").style.display = "none";
  lastDocxPath = "";

  try {
    const res = await fetch("/api/generate", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ dir: dirPath, dept, docx: exportDocx })
    });
    const data = await res.json();
    renderLog(data.log || data.error || "");
    if (data.ok && data.output_pdf) {
      lastPdfPath = data.output_pdf;
      if (data.output_docx) {
        lastDocxPath = data.output_docx;
        document.getElementById("openDocxBtn").style.display = "inline-block";
      }
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

async function openDocx() {
  if (!lastDocxPath) return;
  await fetch("/api/open", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ path: lastDocxPath, reveal: false })
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


def _pick_folder_windows() -> str:
    """Open native Windows folder browser dialog in the foreground without PowerShell startup delay."""
    import ctypes
    from ctypes import wintypes

    class BROWSEINFOW(ctypes.Structure):
        _fields_ = [
            ("hwndOwner", wintypes.HWND),
            ("pidlRoot", ctypes.c_void_p),
            ("pszDisplayName", wintypes.LPWSTR),
            ("lpszTitle", wintypes.LPCWSTR),
            ("ulFlags", wintypes.UINT),
            ("lpfn", ctypes.c_void_p),
            ("lParam", wintypes.LPARAM),
            ("iImage", ctypes.c_int),
        ]

    ole32 = ctypes.windll.ole32
    shell32 = ctypes.windll.shell32
    user32 = ctypes.windll.user32

    ole32.CoInitializeEx.argtypes = [ctypes.c_void_p, wintypes.DWORD]
    ole32.CoInitializeEx.restype = ctypes.c_long
    ole32.CoUninitialize.argtypes = []
    ole32.CoUninitialize.restype = None
    ole32.CoTaskMemFree.argtypes = [ctypes.c_void_p]
    ole32.CoTaskMemFree.restype = None

    shell32.SHBrowseForFolderW.argtypes = [ctypes.POINTER(BROWSEINFOW)]
    shell32.SHBrowseForFolderW.restype = ctypes.c_void_p
    shell32.SHGetPathFromIDListW.argtypes = [ctypes.c_void_p, wintypes.LPWSTR]
    shell32.SHGetPathFromIDListW.restype = wintypes.BOOL

    user32.GetForegroundWindow.argtypes = []
    user32.GetForegroundWindow.restype = wintypes.HWND

    coinit_hr = ole32.CoInitializeEx(None, 0x2)  # COINIT_APARTMENTTHREADED
    try:
        display_buf = ctypes.create_unicode_buffer(260)
        bi = BROWSEINFOW()
        bi.hwndOwner = user32.GetForegroundWindow()
        bi.pszDisplayName = ctypes.cast(display_buf, wintypes.LPWSTR)
        bi.lpszTitle = "请选择报销材料文件夹"
        bi.ulFlags = 0x0001 | 0x0040 | 0x0010  # BIF_RETURNONLYFSDIRS | BIF_NEWDIALOGSTYLE | BIF_EDITBOX

        pidl = shell32.SHBrowseForFolderW(ctypes.byref(bi))
        if not pidl:
            return ""
        try:
            path_buf = ctypes.create_unicode_buffer(4096)
            if shell32.SHGetPathFromIDListW(pidl, path_buf):
                return path_buf.value
            return ""
        finally:
            ole32.CoTaskMemFree(pidl)
    finally:
        if coinit_hr in (0, 1):  # S_OK or S_FALSE
            ole32.CoUninitialize()


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
            try:
                if sys.platform == "win32":
                    folder = _pick_folder_windows()
                else:
                    res = subprocess.run(
                        ["osascript", "-e", 'POSIX path of (choose folder with prompt "请选择报销材料文件夹")'],
                        capture_output=True,
                        text=True,
                        check=False,
                    )
                    folder = res.stdout.strip().rstrip("/")
                self._send_json({"path": folder})
            except Exception as exc:
                self._send_json({"path": "", "error": str(exc)}, status=500)
            return

        if self.path == "/api/generate":
            target_dir = Path(payload.get("dir", "")).expanduser()
            dept = payload.get("dept", "Bida")
            want_docx = bool(payload.get("docx", False))
            output_pdf = target_dir.parent / f"{target_dir.name}_生成.pdf"
            output_docx = (target_dir.parent / f"{target_dir.name}_生成.docx") if want_docx else None
            buf = io.StringIO()
            try:
                with contextlib.redirect_stdout(buf):
                    build_reimbursement_pdf(target_dir, output_pdf, dept=dept, output_docx=output_docx)
                self._send_json({
                    "ok": True,
                    "output_pdf": str(output_pdf),
                    "output_docx": str(output_docx) if output_docx else "",
                    "log": buf.getvalue(),
                })
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
    server = ThreadingHTTPServer(("127.0.0.1", 0), ReimburseHandler)
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
