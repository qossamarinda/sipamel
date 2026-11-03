"""Muat database ISR (SIMS) dan normalisasi satuan."""
from __future__ import annotations
import math
import pandas as pd

NEED = ["CLNT_ID", "CLNT_NAME", "STN_NAME", "STASIUN_LAWAN", "FREQ", "FREQ_PAIR", "BWIDTH",
        "SID_LONG", "SID_LAT", "EQ_MDL", "CURR_LIC_NUM", "LINK_ID", "TCS_ID", "CITY"]


def _f(v):
    try:
        x = float(v)
        return None if math.isnan(x) else x
    except (TypeError, ValueError):
        return None


def _mhz(v):
    """FREQ di SIMS pada contoh sudah MHz (6660), BWIDTH dalam kHz (40000).
    Jika nilainya > 100000 diasumsikan kHz dan dikonversi ke MHz."""
    x = _f(v)
    if x is None:
        return None
    return x / 1000 if x > 100000 else x


def load_isr(path: str, sheet=0) -> list[dict]:
    df = pd.read_excel(path, sheet_name=sheet, dtype=str)
    missing = [c for c in NEED if c not in df.columns]
    if missing:
        raise ValueError(f"Kolom ISR tidak ditemukan: {missing}")
    rows = []
    for i, r in df.iterrows():
        g = lambda k: (r[k].strip() if isinstance(r[k], str) and r[k].strip() not in ("", ".") else None)
        rows.append(dict(
            idx=i, tcs_id=g("TCS_ID"), lic=g("CURR_LIC_NUM"), link_id=g("LINK_ID"),
            stn=g("STN_NAME"), lawan=g("STASIUN_LAWAN"),
            long=_f(r["SID_LONG"]), lat=_f(r["SID_LAT"]), model=g("EQ_MDL"),
            tx=_mhz(r["FREQ"]), rx=_mhz(r["FREQ_PAIR"]),
            bw=(_f(r["BWIDTH"]) / 1000 if _f(r["BWIDTH"]) is not None else None),   # kHz -> MHz
            clnt_id=g("CLNT_ID"), clnt_name=g("CLNT_NAME"), city=g("CITY")))
    return rows
