from __future__ import annotations
import datetime as dt
from .parse import parse_images
from .isr import load_isr
from .match import run_matching, Options
from .export import write_form


def run(images, isr_path, template_path, out_path, backend="rapidocr", tanggal=None,
        options: Options = Options(), stn_override=None, site_coords=None, metode=None):
    from .export import METHOD
    tanggal = tanggal or dt.date.today().isoformat()
    observed = parse_images(images, backend)
    isr = load_isr(isr_path)
    results = run_matching(observed, isr, options, stn_override, site_coords)
    write_form(template_path, out_path, results, tanggal, metode or METHOD)
    return observed, results
