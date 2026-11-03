"""Contoh: python -m mwlink_inspector.cli -i screenshots/ -d isr.xlsx -t template.xlsx -o hasil.xlsx \
          --stn Site_B=BPP047_INPRES2 --coord Site_B=116.8399,-1.2327"""
import argparse, glob, os
from .pipeline import run
from .match import Options
from .export import OPEN_SHELTER


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("-i", "--images", required=True, help="folder atau daftar file gambar")
    ap.add_argument("-d", "--isr", required=True); ap.add_argument("-t", "--template", required=True)
    ap.add_argument("-o", "--out", default="hasil_pemeriksaan.xlsx")
    ap.add_argument("--backend", default="rapidocr", choices=["rapidocr", "tesseract"])
    ap.add_argument("--metode", default="remote", choices=["remote", "shelter"], help="remote=Remote Site, shelter=Open Shelter")
    ap.add_argument("--date"); ap.add_argument("--threshold", type=float, default=80)
    ap.add_argument("--strict-rx", action="store_true", help="FREQ_PAIR kosong = tidak cocok (aturan literal)")
    ap.add_argument("--stn", action="append", default=[], help="KUNCI_SITUS=STN_NAME (override manual)")
    ap.add_argument("--coord", action="append", default=[], help="KUNCI_SITUS=LONG,LAT hasil ukur lapangan")
    a = ap.parse_args()
    imgs = sorted(glob.glob(os.path.join(a.images, "*.png")) + glob.glob(os.path.join(a.images, "*.jp*g"))) \
        if os.path.isdir(a.images) else [a.images]
    stn = dict(s.split("=", 1) for s in a.stn)
    coords = {k: tuple(map(float, v.split(","))) for k, v in (c.split("=", 1) for c in a.coord)}
    obs, res = run(imgs, a.isr, a.template, a.out, a.backend, a.date,
                   Options(name_threshold=a.threshold, strict_rx=a.strict_rx), stn, coords,
                   metode=OPEN_SHELTER if a.metode == 'shelter' else None)
    from collections import Counter
    print(f"{len(obs)} baris terbaca dari {len(imgs)} gambar ->", dict(Counter(r.status for r in res)))
    print("Tersimpan:", a.out)


if __name__ == "__main__":
    main()
