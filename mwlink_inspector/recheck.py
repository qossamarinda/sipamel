"""Cek ulang status setelah form diedit manual (mis. koordinat hasil ukur Open Shelter diubah).
Aturan: untuk baris yang punya DATA ISR dan DATA HASIL PEMERIKSAAN,
  - BW hasil > BW ISR, atau selisih koordinat hasil vs ISR > toleransi (default 30 m)  -> 'Tidak Sesuai Parameter Teknis'
  - selain itu -> 'Sesuai ISR'
Baris 'Tidak Berizin' / 'Tidak Aktif' / status kosong tidak diubah."""
from __future__ import annotations
import sys, openpyxl
from openpyxl.styles import PatternFill
from .match import haversine_m, ST_OK, ST_TECH

FIRST = 7
AUTO = ("Koordinat bergeser", "BW terpasang")


def _f(v):
    try: return float(v)
    except (TypeError, ValueError): return None


def recheck(path: str, out: str, tol_m: float = 30.0) -> list[str]:
    wb = openpyxl.load_workbook(path); ws = wb.active
    log = []
    for r in range(FIRST, ws.max_row + 1):
        st = ws.cell(r, 26).value
        if st not in (ST_OK, ST_TECH):
            continue
        k_lon, k_lat, o_bw = (_f(ws.cell(r, c).value) for c in (11, 12, 15))
        s_lon, s_lat, w_bw = (_f(ws.cell(r, c).value) for c in (19, 20, 23))
        reasons = []
        if w_bw is not None and o_bw is not None and w_bw > o_bw + 1e-9:
            reasons.append(f"BW terpasang {w_bw:g} MHz > BW ISR {o_bw:g} MHz")
        if None not in (k_lon, k_lat, s_lon, s_lat):
            d = haversine_m(s_lon, s_lat, k_lon, k_lat)
            if d > tol_m:
                reasons.append(f"Koordinat bergeser {d:.0f} m (> {tol_m:g} m)")
        new = ST_TECH if reasons else ST_OK
        notes = [n.strip() for n in (ws.cell(r, 28).value or "").split(";") if n.strip() and not n.strip().startswith(AUTO)]
        notes += reasons
        ws.cell(r, 28).value = "; ".join(notes) or None
        if new != st:
            log.append(f"Baris {r}: {st} -> {new} ({'; '.join(reasons) or 'parameter kembali sesuai'})")
        ws.cell(r, 26).value = new
    wb.save(out)
    return log


if __name__ == "__main__":
    inp, outp = sys.argv[1], sys.argv[2]
    tol = float(sys.argv[3]) if len(sys.argv) > 3 else 30.0
    for line in recheck(inp, outp, tol) or ["Tidak ada perubahan status"]:
        print(line)
