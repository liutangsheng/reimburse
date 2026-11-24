from __future__ import annotations

import shutil
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
APP_NAME = "费用报销单生成工具"
VENDOR_TESS = ROOT / "vendor" / "tesseract"
DIST_DIR = ROOT / "dist"
APP_DIR = DIST_DIR / APP_NAME


def main() -> None:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    if hasattr(sys.stderr, "reconfigure"):
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")

    tess_exe = VENDOR_TESS / "tesseract.exe"
    tessdata_dir = VENDOR_TESS / "tessdata"
    required_files = [
        tess_exe,
        tessdata_dir / "eng.traineddata",
        tessdata_dir / "chi_sim.traineddata",
    ]
    missing = [str(p) for p in required_files if not p.exists()]
    if missing:
        raise FileNotFoundError(
            "缺少便携版 Tesseract 依赖文件，请先将其放入 vendor/tesseract/:\n  - "
            + "\n  - ".join(missing)
        )

    cmd = [
        sys.executable,
        "-m",
        "PyInstaller",
        "--noconfirm",
        "--clean",
        "--onedir",
        "--name",
        APP_NAME,
        "--add-data",
        f"{VENDOR_TESS};tesseract",
        "--exclude-module",
        "tkinter",
        "--exclude-module",
        "matplotlib",
        "--exclude-module",
        "numpy",
        "--exclude-module",
        "pandas",
        "--exclude-module",
        "cv2",
        "--exclude-module",
        "paddle",
        "--exclude-module",
        "onnxruntime",
        str(ROOT / "app.py"),
    ]
    print("开始打包 Windows 绿色版...")
    subprocess.run(cmd, cwd=ROOT, check=True)

    zip_base = DIST_DIR / f"{APP_NAME}_Windows绿色版"
    zip_path = Path(shutil.make_archive(str(zip_base), "zip", root_dir=DIST_DIR, base_dir=APP_NAME))
    print(f"\n打包完成！")
    print(f"绿色版目录: {APP_DIR}")
    print(f"分发压缩包: {zip_path}")


if __name__ == "__main__":
    main()
