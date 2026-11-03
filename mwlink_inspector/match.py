"""Logika pencocokan hasil pemeriksaan vs ISR + penentuan status (sesuai aturan poin 4-6)."""
from __future__ import annotations
import math
from dataclasses import dataclass, field
from .names import name_score
from .parse import Observed

ST_OK, ST_NO_LIC, ST_TECH, ST_INACTIVE = "Sesuai ISR", "Tidak Berizin", "Tidak Sesuai Parameter Teknis", "Tidak Aktif"


@dataclass
class Options:
    name_threshold: float = 80      # ambang kemiripan nama (0-100)
    freq_tol: float = 0.001         # MHz
    coord_tol_m: float = 30.0       # toleransi pergeseran koordinat
    strict_rx: bool = False         # True: FREQ_PAIR kosong dianggap tidak cocok (aturan literal)


@dataclass
class Result:
    obs: Observed | None            # None -> baris "Tidak Aktif" (tidak ada di hasil pemeriksaan)
    isr: dict | None
    status: str
    notes: list[str] = field(default_factory=list)
    stn_hasil: str | None = None    # STN_NAME untuk blok DATA HASIL PEMERIKSAAN
    pair_ref: dict | None = None    # ISR pasangan stn+lawan (untuk CLIENT ID/NAME, kota)
    coord: tuple[float, float] | None = None
    verify: bool = False            # True -> perlu verifikasi ulang manual (disorot kuning di form)


def haversine_m(lon1, lat1, lon2, lat2):
    R = 6371000
    p1, p2 = math.radians(lat1), math.radians(lat2)
    a = math.sin((p2 - p1) / 2) ** 2 + math.cos(p1) * math.cos(p2) * math.sin(math.radians(lon2 - lon1) / 2) ** 2
    return 2 * R * math.asin(math.sqrt(a))


def _eq(a, b, tol): return a is not None and b is not None and abs(a - b) <= tol


def run_matching(observed: list[Observed], isr: list[dict], opt: Options = Options(),
                 stn_override: dict[str, str] | None = None,
                 site_coords: dict[str, tuple[float, float]] | None = None) -> list[Result]:
    stn_override, site_coords = stn_override or {}, site_coords or {}
    T = opt.name_threshold
    used: set[int] = set()
    results: list[Result] = []
    pairs_seen: set[tuple[str, str]] = set()   # (stn_isr, lawan_isr) yang terkait pemeriksaan

    def reverse_has_rx(stn_isr, lawan_isr, rx):
        """Rx stasiun ini = FREQ pada record arah sebaliknya (STN_NAME=lawan, STASIUN_LAWAN=stn ini)."""
        return any(_eq(c["tx"], rx, opt.freq_tol) and name_score(c["stn"], lawan_isr) >= T
                   and name_score(c["lawan"], stn_isr) >= T for c in isr)

    # --- tahap 1: kumpulkan kandidat tiap baris pemeriksaan
    cand_info = []
    for ob in observed:
        notes = list(ob.warnings)
        verify = bool(ob.warnings)
        # prioritas nama stasiun: override manual > NE Name di gambar > nama file
        if ob.site in stn_override:
            stn, src = stn_override[ob.site], "manual"
        elif ob.stn_name:
            stn, src = ob.stn_name, "NE Name"
        else:
            stn, src = ob.site, "nama file"
        lawan = ob.stasiun_lawan or ""
        known = any(name_score(stn, r["stn"]) >= T for r in isr if r["stn"])
        inferred = False
        if src == "nama file":
            notes.append(f"STN_NAME diambil dari nama file ('{ob.site}')")
        if known:
            pair = [r for r in isr if name_score(stn, r["stn"]) >= T and name_score(lawan, r["lawan"]) >= T]
        else:   # nama (dari file) tidak dikenal di ISR sama sekali -> coba lewat stasiun lawan + Tx persis
            pair = [r for r in isr if lawan and name_score(lawan, r["lawan"]) >= T and _eq(r["tx"], ob.tx, opt.freq_tol)]
            inferred = True
            verify = True
            notes.append(f"STN_NAME '{stn}' tidak ditemukan di ISR; pencocokan via STASIUN LAWAN + Tx (periksa nama file/NE Name)")
        cand_info.append(dict(ob=ob, stn=stn, src=src, pair=pair, inferred=inferred, notes=notes, verify=verify))

    # --- tahap 2: cocokkan frekuensi; skor 2 = Rx persis, 1 = Rx tidak terverifikasi (FREQ_PAIR kosong)
    scored = []
    for ci in cand_info:
        ob = ci["ob"]
        for r in ci["pair"]:
            if not _eq(r["tx"], ob.tx, opt.freq_tol):
                continue
            if r["rx"] is not None:
                if _eq(r["rx"], ob.rx, opt.freq_tol):
                    scored.append((2, id(ci), r, "")); continue
                continue
            if opt.strict_rx:
                continue
            if reverse_has_rx(r["stn"], r["lawan"], ob.rx):
                scored.append((1.5, id(ci), r, "Rx terverifikasi dari record arah sebaliknya (FREQ_PAIR kosong)"))
            else:
                scored.append((1, id(ci), r, "FREQ_PAIR kosong di ISR: hanya Tx yang cocok, Rx tidak terverifikasi"))
    assigned: dict[int, tuple[dict, str]] = {}
    for sc, cid, r, note in sorted(scored, key=lambda x: -x[0]):
        if cid in assigned or r["idx"] in used:
            continue
        assigned[cid] = (r, note); used.add(r["idx"])

    # --- tahap 3: status per baris
    for ci in cand_info:
        ob, notes = ci["ob"], ci["notes"]
        r, note = assigned.get(id(ci), (None, ""))
        res = Result(obs=ob, isr=r, status=ST_NO_LIC, notes=notes, verify=ci["verify"],
                     stn_hasil=ci["stn"], pair_ref=(ci["pair"][0] if ci["pair"] else None),
                     coord=site_coords.get(ob.site))
        missing = [n for n, v in (("STASIUN LAWAN", ob.stasiun_lawan), ("Tx", ob.tx), ("Rx", ob.rx)) if v in (None, "")]
        if ob.bw is None:
            notes.append("BW tidak terbaca; uji BW dilewati"); res.verify = True
        if missing:      # data OCR tidak lengkap -> tetap masuk form, status dikosongkan untuk verifikasi manual
            res.status = ""; res.verify = True; res.isr = None; r = None
            notes.append("Data OCR tidak lengkap (" + ", ".join(missing) + "); status belum dapat ditentukan")
        elif r is None:
            same_tx = [p for p in ci["pair"] if _eq(p["tx"], ob.tx, opt.freq_tol)]
            if same_tx:
                notes.append(f"Frekuensi sama dengan {len(same_tx)} ISR yang sudah terpakai oleh baris lain (duplikat di pemeriksaan, ISR tidak mencukupi)")
            elif ci["pair"]:
                notes.append("Pasangan stasiun terdaftar di ISR, tetapi kombinasi Tx/Rx tidak ada")
            else:
                notes.append("Stasiun/stasiun lawan tidak ditemukan di ISR")
        else:
            if note:
                notes.append(note); res.verify = True
            best = max(s[0] for s in scored if s[1] == id(ci) and s[2]["idx"] == r["idx"])
            if any(s[1] == id(ci) and s[2]["idx"] != r["idx"] and s[0] == best for s in scored):
                notes.append("Lebih dari satu ISR sama-sama cocok (mis. beda polarisasi); periksa manual"); res.verify = True
            res.status = ST_OK
            if ob.bw is not None and r["bw"] is not None and ob.bw > r["bw"] + 1e-9:
                res.status = ST_TECH; notes.append(f"BW terpasang {ob.bw:g} MHz > BW ISR {r['bw']:g} MHz")
            if res.coord and r["long"] is not None and r["lat"] is not None:
                d = haversine_m(res.coord[0], res.coord[1], r["long"], r["lat"])
                if d > opt.coord_tol_m:
                    res.status = ST_TECH; notes.append(f"Koordinat bergeser {d:.0f} m (> {opt.coord_tol_m:g} m)")
        if r is not None:
            pairs_seen.add((r["stn"], r["lawan"]))
        for p in ci["pair"]:
            pairs_seen.add((p["stn"], p["lawan"]))
        results.append(res)

    # --- tahap 4 (poin 6): ISR dengan pasangan sama, frekuensinya tidak ada di pemeriksaan -> Tidak Aktif
    for r in isr:
        if r["idx"] in used or (r["stn"], r["lawan"]) not in pairs_seen:
            continue
        results.append(Result(obs=None, isr=r, status=ST_INACTIVE, pair_ref=r,
                              notes=["Ada di ISR tetapi frekuensi tidak ditemukan pada perangkat"]))
    return results
