"""Lapisan OCR yang bisa diganti (pluggable). Semua backend mengembalikan list[Token]
dengan koordinat piksel, sehingga parser bekerja berbasis POSISI (layout-aware), bukan teks lurus."""
from __future__ import annotations
import csv, io, subprocess, tempfile
from dataclasses import dataclass


@dataclass
class Token:
    text: str
    x1: float
    y1: float
    x2: float
    y2: float
    conf: float = 1.0

    @property
    def cx(self): return (self.x1 + self.x2) / 2
    @property
    def cy(self): return (self.y1 + self.y2) / 2


_rapid = None
TARGET_WIDTH = 1920

def ocr_rapidocr(path: str) -> list[Token]:
    """PP-OCR (PaddleOCR) via ONNX Runtime. Ringan, CPU-only, model sudah terbundel di paket pip."""
    global _rapid
    if _rapid is None:
        from rapidocr_onnxruntime import RapidOCR
        _rapid = RapidOCR(intra_op_num_threads=1, inter_op_num_threads=1)
    # Normalisasi resolusi: gambar kecil/diperkecil diperbesar dulu (akurasi OCR turun tajam pada teks < ~12 px)
    from PIL import Image
    import numpy as np
    im = Image.open(path).convert("RGB")
    f = min(2.0, max(1.0, TARGET_WIDTH / im.width))
    if f > 1.01:
        im = im.resize((round(im.width * f), round(im.height * f)), Image.LANCZOS)
    res, _ = _rapid(np.ascontiguousarray(np.array(im)), use_cls=False)   # RapidOCR menganggap ndarray = RGB
    out = []
    for box, txt, conf in (res or []):
        xs = [p[0] / f for p in box]; ys = [p[1] / f for p in box]
        out.append(Token(txt.strip(), min(xs), min(ys), max(xs), max(ys), float(conf)))
    return out


def ocr_tesseract(path: str, lang: str = "eng") -> list[Token]:
    """Tesseract (CLI). Output level-kata digabung menjadi frasa agar setara dengan RapidOCR."""
    r = subprocess.run(["tesseract", path, "stdout", "-l", lang, "--psm", "11", "tsv"],
                       capture_output=True, text=True, check=True)
    words = []
    for row in csv.DictReader(io.StringIO(r.stdout), delimiter="\t", quoting=csv.QUOTE_NONE):
        t = (row.get("text") or "").strip()
        if t and float(row["conf"]) >= 0:
            x, y, w, h = (float(row[k]) for k in ("left", "top", "width", "height"))
            words.append(Token(t, x, y, x + w, y + h, float(row["conf"]) / 100))
    words.sort(key=lambda t: (round(t.cy / 8), t.x1))
    merged: list[Token] = []
    for w in words:
        if merged and abs(merged[-1].cy - w.cy) < 8 and 0 <= w.x1 - merged[-1].x2 < 14:
            m = merged[-1]
            merged[-1] = Token(m.text + " " + w.text, m.x1, min(m.y1, w.y1), w.x2, max(m.y2, w.y2), min(m.conf, w.conf))
        else:
            merged.append(w)
    return merged


BACKENDS = {"rapidocr": ocr_rapidocr, "tesseract": ocr_tesseract}

def run_ocr(path: str, backend: str = "rapidocr") -> list[Token]:
    toks = BACKENDS[backend](path)
    for t in toks:      # OCR kadang menghasilkan tanda kurung/titik dua lebar-penuh (CJK)
        t.text = t.text.translate(str.maketrans({"（": "(", "）": ")", "：": ":", "［": "[", "］": "]"}))
    return toks
