from __future__ import annotations

import contextlib
import io
import os
import shutil
import subprocess
import sys
import uuid
from pathlib import Path

from PIL import Image

_TESS_HAS_CHI_SIM: bool | None = None


def _subprocess_kwargs() -> dict:
    """Suppress extra console windows when invoking subprocesses on Windows."""
    if sys.platform == "win32":
        return {"creationflags": subprocess.CREATE_NO_WINDOW}
    return {}


def _find_tesseract() -> str:
    """Locate tesseract binary from bundled package, local vendor dir, PATH, or system install."""
    exe_name = "tesseract.exe" if sys.platform == "win32" else "tesseract"
    bundled_roots: list[Path] = []
    meipass = getattr(sys, "_MEIPASS", None)
    if meipass:
        bundled_roots.append(Path(meipass))
    if getattr(sys, "frozen", False):
        exe_dir = Path(sys.executable).resolve().parent
        bundled_roots.extend([exe_dir, exe_dir / "_internal"])
    bundled_roots.append(Path(__file__).resolve().parent.parent / "vendor")

    for root in bundled_roots:
        candidate = root / "tesseract" / exe_name
        if candidate.exists():
            return str(candidate)

    found = shutil.which("tesseract")
    if found:
        return found
    for candidate in (
        "/opt/homebrew/bin/tesseract",
        "/usr/local/bin/tesseract",
        r"C:\Program Files\Tesseract-OCR\tesseract.exe",
        r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
    ):
        if Path(candidate).exists():
            return candidate
    return "tesseract"


def _find_tessdata_dir(tess_exe: str) -> str | None:
    """Locate bundled tessdata directory next to tesseract binary if present."""
    tess_path = Path(tess_exe)
    if tess_path.is_file():
        sibling = tess_path.parent / "tessdata"
        if sibling.is_dir():
            return str(sibling)
    meipass = getattr(sys, "_MEIPASS", None)
    if meipass:
        bundled = Path(meipass) / "tesseract" / "tessdata"
        if bundled.is_dir():
            return str(bundled)
    return None


def _has_tesseract_chi_sim() -> bool:
    """Check whether tesseract has the chi_sim language pack installed."""
    global _TESS_HAS_CHI_SIM
    if _TESS_HAS_CHI_SIM is None:
        try:
            tess_exe = _find_tesseract()
            cmd = [tess_exe, "--list-langs"]
            tessdata_dir = _find_tessdata_dir(tess_exe)
            if tessdata_dir:
                cmd.extend(["--tessdata-dir", tessdata_dir])
            res = subprocess.run(
                cmd,
                capture_output=True,
                text=True,
                encoding="utf-8",
                errors="replace",
                check=False,
                **_subprocess_kwargs(),
            )
            langs = {line.strip() for line in (res.stdout or "").splitlines()}
            _TESS_HAS_CHI_SIM = "chi_sim" in langs
        except OSError:
            _TESS_HAS_CHI_SIM = False
    return _TESS_HAS_CHI_SIM


def _to_grayscale_on_white(im: Image.Image) -> Image.Image:
    """Convert a PIL image to grayscale, compositing transparent pixels over white first."""
    if im.mode in ("RGBA", "LA") or (im.mode == "P" and "transparency" in im.info):
        rgba = im.convert("RGBA")
        bg = Image.new("RGBA", rgba.size, (255, 255, 255, 255))
        return Image.alpha_composite(bg, rgba).convert("L")
    return im.convert("L")


def _ocr_image(
    img: Image.Image,
    psm: int | None = None,
    work_dir: Path | None = None,
    lang: str = "eng",
) -> str:
    """Run tesseract OCR on a PIL Image using a temporary file in the current directory."""
    base_dir = Path.cwd()
    if not os.access(base_dir, os.W_OK) and work_dir is not None:
        base_dir = work_dir
    tmp_path = base_dir / f"_ocr_tmp_{uuid.uuid4().hex}.png"
    tess_exe = _find_tesseract()
    cmd = [tess_exe, str(tmp_path), "stdout", "-l", lang]
    tessdata_dir = _find_tessdata_dir(tess_exe)
    if tessdata_dir:
        cmd.extend(["--tessdata-dir", tessdata_dir])
    if psm is not None:
        cmd.extend(["--psm", str(psm)])
    try:
        img.save(tmp_path)
        res = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            check=False,
            **_subprocess_kwargs(),
        )
        return res.stdout or ""
    finally:
        with contextlib.suppress(OSError):
            tmp_path.unlink(missing_ok=True)


def _image_to_rgb_png(img_path: Path) -> tuple[io.BytesIO, int, int]:
    """Load an image file (compositing alpha over white) into an RGB PNG buffer and return (buf, w, h)."""
    with Image.open(img_path) as raw:
        if raw.mode in ("RGBA", "LA") or (raw.mode == "P" and "transparency" in raw.info):
            rgba = raw.convert("RGBA")
            bg = Image.new("RGBA", rgba.size, (255, 255, 255, 255))
            im = Image.alpha_composite(bg, rgba).convert("RGB")
        else:
            im = raw.convert("RGB")
        iw, ih = im.size
        buf = io.BytesIO()
        im.save(buf, format="PNG")
        buf.seek(0)
        return buf, iw, ih


def _render_pdf_pages_to_pngs(pdf_path: Path, scale: float = 2.5) -> list[io.BytesIO]:
    """Render each page of a PDF file into an in-memory PNG buffer via pypdfium2."""
    import pypdfium2 as pdfium

    buffers: list[io.BytesIO] = []
    pdf = pdfium.PdfDocument(str(pdf_path))
    try:
        for page in pdf:
            bitmap = page.render(scale=scale)
            pil_img = bitmap.to_pil().convert("RGB")
            buf = io.BytesIO()
            pil_img.save(buf, format="PNG")
            buf.seek(0)
            buffers.append(buf)
    finally:
        pdf.close()
    return buffers
