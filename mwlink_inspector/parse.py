"""Parser screenshot NEC iPASOLINK LCT. Dua jenis halaman:
 (1) Provisioning > MODEM Port Setting  -> slot -> nama port ("TO ...") = STASIUN LAWAN
 (2) Equipment Setup > Radio Configuration -> slot -> Channel Spacing, TX RF, RX RF
Kedua halaman digabung per SLOT (lebih aman daripada urutan kolom: slot bisa loncat, mis. 01-08, 11, 12)."""
from __future__ import annotations
import os, re
from dataclasses import dataclass, field
from .ocr import Token, run_ocr

DIGIT_FIX = str.maketrans({"O": "0", "o": "0", "l": "1", "I": "1", "|": "1", "S": "5", "B": "8"})
SLOT_RE = re.compile(r"Slot\s*([0-9OolI]{1,2})", re.I)
STOP_RE = re.compile(r"^(Rad|Version|802|Disable|Low|Latency|Enable)", re.I)
CONF_WARN = 0.90


@dataclass
class PortPage:
    ne_name: str | None = None
    rows: list[tuple[int | None, str]] = field(default_factory=list)    # (label slot yang terbaca | None, stasiun lawan), urut atas->bawah
    warnings: list[str] = field(default_factory=list)


@dataclass
class RadioPage:
    ne_name: str | None = None
    labels: list[int | None] = field(default_factory=list)              # label slot per kolom, urut kiri->kanan
    bw: dict[int, float] = field(default_factory=dict)                  # kunci = indeks kolom (0..n-1)
    tx: dict[int, float] = field(default_factory=dict)
    rx: dict[int, float] = field(default_factory=dict)
    shift: dict[int, float] = field(default_factory=dict)
    warnings: list[str] = field(default_factory=list)


STRICT_SLOT = re.compile(r"Slot\s*(\d{1,2})")                           # label slot yang terbaca bersih
# Dua varian tampilan NEC: "MODEM (Slot01)" (IDU banyak slot) dan "MODEM1" (varian satu modem)
PORT_ROW = re.compile(r"MODEM\s*\(|^MODEM\s*\d{1,2}\.?$")                  # baris port
HEAD_COL = re.compile(r"\(\s*S\w{1,3}t\s*\w{1,3}\s*\)|^MODEM\s*\d{1,2}\.?$")   # header kolom radio
BARE_MODEM = re.compile(r"^MODEM\s*(\d{1,2})\.?$")


def _label(text: str) -> int | None:
    m = STRICT_SLOT.search(text) or BARE_MODEM.search(text.strip())
    return int(m.group(1)) if m else None


def _unit(tokens: list[Token]) -> float:
    """Skala gambar relatif terhadap tinggi teks referensi (~17 px). Semua toleransi piksel dikalikan nilai ini."""
    hs = sorted(t.y2 - t.y1 for t in tokens if t.y2 > t.y1)
    return max(0.4, (hs[len(hs) // 2] / 17.0)) if hs else 1.0


def find_ne_name(tokens: list[Token]) -> str | None:
    """Ambil teks setelah 'NE Name:'. Jika tidak ada (mis. hanya 'Login User:Admin') -> None."""
    for t in tokens:
        m = re.search(r"NE\s*Name\s*[:;]?\s*(\S+)", t.text, re.I)
        if m:
            return m.group(1).strip()
    return None


def _num(text: str) -> float | None:
    s = text.strip().replace(",", "")
    if re.search(r"\d", s):
        s = re.sub(r"(?<=\d)[Oo]", "0", s)
    m = re.search(r"\d+(?:\.\d+)?", s)
    return float(m.group()) if m else None


def _bw(text: str) -> float | None:
    return _num(text) if re.search(r"M\s*H\s*[zZ2]", text, re.I) else None


def classify(tokens: list[Token]) -> str:
    blob = " ".join(t.text for t in tokens)
    if re.search(r"TX\s*RF\s*Freq", blob, re.I):
        return "radio"
    if re.search(r"Port\s*Name", blob, re.I) or any(PORT_ROW.search(t.text) for t in tokens):
        return "port"
    return "unknown"


def parse_port_page(tokens: list[Token]) -> PortPage:
    page = PortPage(ne_name=find_ne_name(tokens))
    k = _unit(tokens)
    for t in sorted((t for t in tokens if PORT_ROW.search(t.text)), key=lambda t: t.cy):
        row = sorted((u for u in tokens if u is not t and abs(u.cy - t.cy) <= 10 * k and u.x1 >= t.x2 - 2 * k),
                     key=lambda u: u.x1)
        parts = []
        for u in row:
            if STOP_RE.match(u.text):
                break
            parts.append(u)
        inline = re.sub(r"^[\s\.\,:;]+", "", t.text.split(")", 1)[1]) if ")" in t.text else ""      # "MODEM (Slot02)TO KLAMONO" dalam satu token
        name = " ".join([inline] + [u.text for u in parts]).strip()
        name = re.sub(r"^(?:(?i:TO)\s+|TO(?=\d{0,3}[A-Z]{2,4}\d{2,4}))", "", name)
        lab = _label(t.text)
        tag = f"Slot{lab:02d}" if lab is not None else f"baris {len(page.rows) + 1}"
        if not name:
            page.warnings.append(f"{tag}: nama port kosong/tidak terbaca")
        elif min((u.conf for u in parts), default=1) < CONF_WARN:
            page.warnings.append(f"{tag}: confidence OCR rendah pada nama port ('{name}')")
        page.rows.append((lab, name))
    return page


def parse_radio_page(tokens: list[Token]) -> RadioPage:
    page = RadioPage(ne_name=find_ne_name(tokens))
    k = _unit(tokens)
    heads = [t for t in tokens if HEAD_COL.search(t.text)]
    if not heads:
        page.warnings.append("Header kolom (SlotNN) tidak ditemukan")
        return page
    top = min(t.cy for t in heads)
    heads = sorted((t for t in heads if t.cy - top < 40 * k), key=lambda t: t.cx)
    cols = {i: t.cx for i, t in enumerate(heads)}
    page.labels = [_label(t.text) for t in heads]

    def label_row(pattern: str):
        for t in tokens:
            if re.search(pattern, t.text, re.I) and t.x1 < min(cols.values()):
                return t
        return None

    FRIENDLY = {r"Channel\s*Spacing": "Channel Spacing", r"TX\s*RF\s*Freq": "TX RF Frequency", r"RX\s*RF\s*Freq": "RX RF Frequency", r"Shift\s*Freq": "Shift Frequency"}

    def read_row(pattern: str, parse):
        lab = label_row(pattern)
        out: dict[int, float] = {}
        if lab is None:
            return out
        for t in tokens:
            if t is lab or t.x1 <= lab.x2 or abs(t.cy - lab.cy) > 12 * k:
                continue
            i = min(cols, key=lambda c: abs(cols[c] - t.cx))
            if abs(cols[i] - t.cx) > 60 * k:
                continue
            v = parse(t.text)
            if v is not None:
                out[i] = v
                if t.conf < CONF_WARN:
                    page.warnings.append(f"Kolom {i + 1}: confidence OCR rendah pada {FRIENDLY.get(pattern, pattern)} ({t.text}, {t.conf:.2f})")
        return out

    page.bw = read_row(r"Channel\s*Spacing", _bw)
    if len(page.bw) < len(cols):    # cadangan: label baris tidak terbaca -> ambil token "NNMHz" yang sejajar kolom
        for t in tokens:
            if re.fullmatch(r"\d+(?:\.\d+)?\s*MHz", t.text.strip(), re.I):
                i = min(cols, key=lambda c: abs(cols[c] - t.cx))
                if abs(cols[i] - t.cx) <= 60 * k and i not in page.bw:
                    page.bw[i] = _num(t.text)
    page.tx = read_row(r"TX\s*RF\s*Freq", _num)
    page.rx = read_row(r"RX\s*RF\s*Freq", _num)
    page.shift = read_row(r"Shift\s*Freq", _num)       # opsional: tabel ODU di bagian bawah
    for name, d in (("Channel Spacing", page.bw), ("TX RF Frequency", page.tx), ("RX RF Frequency", page.rx)):
        miss = [str(i + 1) for i in cols if i not in d]
        if miss:
            page.warnings.append(f"{name}: tidak terbaca pada kolom ke-{', '.join(miss)}")
    for name, d in (("TX", page.tx), ("RX", page.rx)):         # frekuensi MW wajar: 1 - 100 GHz
        for i, v in d.items():
            if not 1000 <= v <= 100000:
                page.warnings.append(f"Kolom {i + 1}: nilai {name} {v:g} MHz tidak wajar; kemungkinan salah baca OCR")
    for i, sh in page.shift.items():                     # |Rx - Tx| harus = Shift Frequency
        if i in page.tx and i in page.rx and abs(abs(page.rx[i] - page.tx[i]) - sh) > 0.5:
            page.warnings.append(f"Kolom {i + 1}: selisih Rx-Tx ({abs(page.rx[i]-page.tx[i]):g}) != Shift Freq ({sh:g}); cek hasil OCR")
    return page


# ---------------------------------------------------------------- gabungkan per situs
@dataclass
class Observed:
    site: str                 # kunci grup (dari nama file)
    slot: int
    stn_name: str | None      # dari NE Name (None bila tidak ada di gambar)
    stasiun_lawan: str | None
    tx: float | None
    rx: float | None
    bw: float | None
    source_files: list[str] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)


def site_key(path: str) -> str:
    """Nama situs dari nama file: buang ekstensi, salinan "(2)", lalu nomor halaman di akhir yang DIPISAH
    pemisah (_ - spasi). Angka yang menempel pada nama (INPRES2) tidak dibuang."""
    stem = os.path.splitext(os.path.basename(path))[0].strip()
    stem = re.sub(r"\s*\(\d+\)$", "", stem)                 # "Site_A_1 (2)" -> "Site_A_1"
    key = re.sub(r"[\s_\-]+\d{1,2}$", "", stem)              # "Site_A_1" -> "Site_A"
    return key or stem


def _natural(path: str):
    return [int(t) if t.isdigit() else t.lower() for t in re.split(r"(\d+)", os.path.basename(path))]


def _join(ports, cols):
    """Jodohkan baris Port Setting dengan kolom Radio Configuration. Return list (label, nama, bw, tx, rx, peringatan)."""
    pairs = []
    if len(ports) == len(cols):      # jumlah sama -> per URUTAN (kebal terhadap nomor slot yang salah baca)
        for (pl, name), (cl, bw, tx, rx) in zip(ports, cols):
            w = []
            if pl is not None and cl is not None and pl != cl:
                w.append(f"Slot port ({pl:02d}) berbeda dengan slot kolom radio ({cl:02d}); dijodohkan per urutan")
            pairs.append((pl if pl is not None else cl, name, bw, tx, rx, w))
    else:                             # jumlah beda -> per nomor slot yang terbaca bersih
        pm = {l: nm for l, nm in ports if l is not None}
        cm = {c[0]: c for c in cols if c[0] is not None}
        for lab in sorted(set(pm) | set(cm)):
            c = cm.get(lab, (lab, None, None, None))
            w = ["Jumlah baris Port Setting (%d) != kolom Radio Configuration (%d); dijodohkan per nomor slot" % (len(ports), len(cols))]
            if lab not in pm: w.append(f"Slot{lab:02d}: tidak ada di halaman Port Setting")
            if lab not in cm: w.append(f"Slot{lab:02d}: tidak ada di halaman Radio Configuration")
            pairs.append((lab, pm.get(lab), c[1], c[2], c[3], w))
    return pairs


def parse_images(paths: list[str], backend: str = "rapidocr") -> list[Observed]:
    """Satu situs = semua gambar dengan awalan nama file yang sama. Situs dengan >1 kabinet/IDU boleh memakai
    beberapa pasang gambar (Port Setting + Radio Configuration); pasangan dibentuk per urutan nomor file."""
    groups: dict[str, dict] = {}
    for p in sorted(paths, key=_natural):
        g = groups.setdefault(site_key(p), {"ne": None, "port_pages": [], "radio_pages": [], "files": [], "warn": []})
        g["files"].append(os.path.basename(p))
        try:
            toks = run_ocr(p, backend)
        except Exception as e:                       # file rusak / OCR gagal -> tetap lanjut, tandai
            g["warn"].append(f"{os.path.basename(p)}: OCR gagal ({type(e).__name__})")
            continue
        kind = classify(toks)
        if kind == "port":
            pg = parse_port_page(toks); g["port_pages"].append(pg.rows); g["warn"] += pg.warnings
            g["ne"] = g["ne"] or pg.ne_name
        elif kind == "radio":
            pg = parse_radio_page(toks); g["radio_pages"].append(pg); g["warn"] += pg.warnings
            g["ne"] = g["ne"] or pg.ne_name
        else:
            g["warn"].append(f"{os.path.basename(p)}: jenis halaman tidak dikenali (bukan Port Setting / Radio Configuration)")
    out: list[Observed] = []
    for site, g in groups.items():
        radio_cols = [[(lab, pg.bw.get(i), pg.tx.get(i), pg.rx.get(i)) for i, lab in enumerate(pg.labels)]
                      for pg in g["radio_pages"]]
        port_pages = g["port_pages"]
        if not any(port_pages) and not any(radio_cols):   # tidak ada data terbaca -> baris kosong untuk verifikasi manual
            out.append(Observed(site, 0, g["ne"], None, None, None, None, g["files"],
                                g["warn"] + ["Tidak ada data yang berhasil dibaca dari gambar"]))
            continue
        pairs = []
        if len(port_pages) == len(radio_cols):            # satu pasang per kabinet, sesuai urutan nomor file
            for i, (pr, cr) in enumerate(zip(port_pages, radio_cols)):
                extra = [f"Pasangan gambar ke-{i + 1} dari {len(port_pages)} (dijodohkan per urutan nomor file)"] if len(port_pages) > 1 else []
                pairs += [(*x[:5], x[5] + extra) for x in _join(pr, cr)]
        else:
            pairs = _join([r for pr in port_pages for r in pr], [c for cr in radio_cols for c in cr])
            pairs = [(*x[:5], x[5] + [f"Jumlah halaman Port Setting ({len(port_pages)}) != Radio Configuration ({len(radio_cols)}); dijodohkan gabungan"]) for x in pairs]
        for i, (lab, name, bw, tx, rx, w) in enumerate(pairs):
            ws_ = (list(g["warn"]) if i == 0 else []) + w
            out.append(Observed(site, lab if lab is not None else i + 1, g["ne"], name or None, tx, rx, bw, g["files"], ws_))
    return out

def _clean_text(v) -> str | None:
    if v is None or (isinstance(v, float) and v != v):       # None / NaN
        return None
    v = str(v).strip()
    return v or None


def _clean_num(v) -> float | None:
    try:
        x = float(v)
    except (TypeError, ValueError):
        return None
    return None if x != x else x


def df_to_observed(df) -> list[Observed]:
    """Ubah tabel (hasil OCR yang mungkin sudah diedit di UI) menjadi daftar Observed.
    Sel kosong dari tabel berupa NaN -> None; baris tambahan tanpa slot/site tetap diterima."""
    out = []
    for i, r in enumerate(df.to_dict("records")):
        slot = _clean_num(r.get("slot"))
        out.append(Observed(
            site=_clean_text(r.get("site")) or "tanpa_nama",
            slot=int(slot) if slot is not None else i + 1,
            stn_name=_clean_text(r.get("stn_name")),
            stasiun_lawan=_clean_text(r.get("stasiun_lawan")),
            tx=_clean_num(r.get("tx")), rx=_clean_num(r.get("rx")), bw=_clean_num(r.get("bw")),
            warnings=([w] if (w := _clean_text(r.get("peringatan"))) else [])))
    return out
