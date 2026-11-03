"""Isi 'Template Format Pemeriksaan (Remote Shelter)'. Gaya, lebar kolom, merge header, dan dropdown
validasi template dipertahankan; baris contoh (baris 7 ke bawah) diganti dengan hasil."""
from __future__ import annotations
import copy, re
import openpyxl
from openpyxl.styles import PatternFill
from .match import Result

FIRST = 7          # baris data pertama pada template
METHOD = "Pemeriksaan melalui Remote Site"
OPEN_SHELTER = "Inspeksi melalui Open Shelter"
YELLOW = PatternFill("solid", fgColor="FFFF00")
DIST = ('=IF(AND(ISNUMBER(K{r}),ISNUMBER(L{r}),ISNUMBER(S{r}),ISNUMBER(T{r})),'
        'ROUND(12742000*ASIN(SQRT(SIN(RADIANS(T{r}-L{r})/2)^2+COS(RADIANS(L{r}))*COS(RADIANS(T{r}))'
        '*SIN(RADIANS(S{r}-K{r})/2)^2)),1),"")')


def _clean_name(s):  # nama stasiun lawan dari OCR: rapikan spasi ganda
    return re.sub(r"\s+", " ", s).strip() if s else None


def write_form(template_path: str, out_path: str, results: list[Result], tanggal: str,
               metode: str = METHOD, device_present: str | None = None) -> str:
    wb = openpyxl.load_workbook(template_path)
    ws = wb.active
    last_tpl = ws.max_row
    style_src = {c: copy.copy(ws.cell(FIRST, c)._style) for c in range(1, 30)}
    style_src[30] = style_src[29]
    ws.merge_cells('AD3:AD5'); ws['AD3'] = 'SELISIH KOORDINAT (m)'; ws['AD3']._style = copy.copy(ws['AC3']._style)
    ws.column_dimensions['AD'].width = 16
    height = ws.row_dimensions[FIRST].height
    for r in range(FIRST, last_tpl + 1):               # kosongkan contoh
        for c in range(1, 31):
            ws.cell(r, c).value = None

    for i, res in enumerate(results):
        r = FIRST + i
        for c in range(1, 31):
            ws.cell(r, c)._style = copy.copy(style_src[c])
        ws.row_dimensions[r].height = height
        isr, ob = res.isr, res.obs
        ref = isr or res.pair_ref or {}
        notes = list(res.notes)
        verify = res.verify
        matched = bool(ob and isr)
        # koordinat hasil: = data ISR (Remote Site & Open Shelter); bila ada nilai ukur manual, pakai itu
        coord = (res.coord if res.coord else ((isr["long"], isr["lat"]) if matched else None))
        if matched and metode == OPEN_SHELTER and not res.coord:
            notes.append("Koordinat sementara = data ISR; konfirmasi dengan hasil ukur lapangan"); verify = True
        row = {
            1: i + 1, 2: tanggal, 3: metode,
            4: ref.get("clnt_id"), 5: ref.get("clnt_name"),
            # --- DATA ISR (SIMS) : hanya terisi bila cocok / Tidak Aktif
            6: isr and isr["tcs_id"], 7: isr and isr["lic"], 8: isr and isr["link_id"],
            9: isr and isr["stn"], 10: isr and isr["lawan"],
            11: isr and isr["long"], 12: isr and isr["lat"],
            13: isr and isr["tx"], 14: isr and isr["rx"], 15: isr and isr["bw"], 16: isr and isr["model"],
            # --- DATA HASIL PEMERIKSAAN : kosong untuk baris 'Tidak Aktif'
            17: ob and res.stn_hasil, 18: ob and _clean_name(ob.stasiun_lawan),
            19: coord and coord[0], 20: coord and coord[1],
            21: ob and ob.tx, 22: ob and ob.rx, 23: ob and ob.bw,
            25: device_present if ob else None,
            26: res.status or None, 28: (("[VERIFIKASI] " if verify else "") + "; ".join(notes)) or None, 29: ref.get("city"),
            30: DIST.format(r=r),
        }
        for c, v in row.items():
            ws.cell(r, c).value = v if v not in ("", False) else None
        if verify:
            for c in (26, 28):
                ws.cell(r, c).fill = YELLOW
        for c in (11, 12, 13, 14, 15, 19, 20, 21, 22, 23):
            ws.cell(r, c).number_format = "General"
    end = FIRST + len(results) - 1
    if end < last_tpl:                                  # buang baris contoh berlebih
        ws.delete_rows(end + 1, last_tpl - end)
    # perpanjang dropdown & area cetak mengikuti jumlah baris
    for dv in ws.data_validations.dataValidation:
        col = str(dv.sqref)[0:1] if len(str(dv.sqref)) else ""
        dv.sqref = openpyxl.worksheet.cell_range.MultiCellRange(f"{re.match(r'[A-Z]+', str(dv.sqref).split(':')[0]).group()}6:{re.match(r'[A-Z]+', str(dv.sqref).split(':')[0]).group()}{max(end, 24)}")
    ws.print_area = f"A6:AD{max(end, 22)}"
    for c in (28,):                                      # KETERANGAN: bungkus teks
        for r in range(FIRST, end + 1):
            al = copy.copy(ws.cell(r, c).alignment); al.wrap_text = True; ws.cell(r, c).alignment = al
    wb.save(out_path)
    return out_path
