"""
MWLink OCR Engine for SiPaMeL
==============================
Integrasi mesin OCR MWLink_Inspector ke dalam SiPaMeL:
1. Membaca gambar capture operator seluler (format NEC iPASOLINK LCT dll).
2. Mengekstrak 6 Parameter Utama:
   - Tx RF (MHz) dari Equipment Setup - Radio Configuration
   - Rx RF (MHz) dari Equipment Setup - Radio Configuration
   - Bandwidth (tetap dalam satuan MHz, tidak dikonversi ke kHz)
   - STN_Name dari NE Name
   - Stasiun Lawan dari Provisioning - Modem Function Setting - Modem Port Setting (kolom Port Name)
   - Merk/Type Perangkat dari Equipment Setup - Equipment Configuration / NEC
3. Mencocokkan hasil identifikasi dengan data stasiun lawan di database.
4. Menuliskan hasil secara otomatis ke kolom hasil pemeriksaan stasiun lawan.
"""

from __future__ import annotations
import os
import re
import sys
from datetime import datetime, date
from typing import List, Dict, Any, Optional, Tuple

# Pastikan folder sistem monitor ada di urutan pertama sys.path agar app.py lokal tidak tertimpa
CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))
if CURRENT_DIR not in sys.path:
    sys.path.insert(0, CURRENT_DIR)
elif sys.path[0] != CURRENT_DIR:
    sys.path.remove(CURRENT_DIR)
    sys.path.insert(0, CURRENT_DIR)

# Tambahkan MWLink_Inspector ke akhir sys.path hanya untuk mengakses package mwlink_inspector
PROJECT_PARENT = os.path.dirname(CURRENT_DIR)
MWLINK_DIR = os.path.join(PROJECT_PARENT, 'MWLink_Inspector')
if os.path.exists(MWLINK_DIR) and MWLINK_DIR not in sys.path:
    sys.path.append(MWLINK_DIR)

try:
    from mwlink_inspector.ocr import run_ocr, Token
    from mwlink_inspector.parse import (
        PortPage, RadioPage, Observed, classify, 
        parse_port_page, parse_radio_page, find_ne_name,
        _unit, _num, _label, HEAD_COL, PORT_ROW, site_key, _natural, _join
    )
    from mwlink_inspector.names import name_score, norm
    MWLINK_AVAILABLE = True
except Exception as e:
    MWLINK_AVAILABLE = False
    print(f"⚠️ Warning: MWLink_Inspector gagal di-import: {e}")


def parse_bandwidth_khz_to_mhz(text: str) -> Optional[float]:
    """
    Ekstraksi bandwidth dari string (tetap dalam satuan MHz, tidak dikonversi ke kHz).
    Contoh:
    - '28000 kHz' / '28000' -> 28.0 MHz
    - '40000 kHz' / '40000' -> 40.0 MHz
    - '56000 kHz' / '56000' -> 56.0 MHz
    - '28MHz' / '28.0 MHz' / '40MHz' -> 28.0 MHz / 40.0 MHz
    """
    if not text:
        return None
    
    clean_text = text.strip().replace(',', '.')
    # Cari angka desimal atau integer
    m = re.search(r"(\d+(?:\.\d+)?)", clean_text)
    if not m:
        return None
    
    try:
        val = float(m.group(1))
    except ValueError:
        return None

    # Jika teks eksplisit ada kHz atau nilainya > 500 (frekuensi MW channel spacing umumnya 3.5 - 112 MHz)
    is_khz = bool(re.search(r"k\s*h\s*[zZ2]", clean_text, re.I)) or val > 500
    if is_khz:
        mhz_val = val / 1000.0
    else:
        mhz_val = val

    return round(mhz_val, 2)


def extract_equipment_device_info(tokens: List[Token]) -> Dict[str, Any]:
    """
    Ekstraksi Merk dan Tipe Perangkat dari:
    - Equipment Setup - Equipment Configuration
    - Atau header bar / logo NEC di layar LCT
    """
    blob = " ".join(t.text for t in tokens)
    info = {
        'brand': 'NEC',
        'model': None,
        'full_name': 'NEC'
    }

    # Cek Brand
    if re.search(r"\bNEC\b", blob, re.I):
        info['brand'] = 'NEC'
    elif re.search(r"\bHUAWEI\b", blob, re.I):
        info['brand'] = 'HUAWEI'
    elif re.search(r"\bERICSSON\b", blob, re.I):
        info['brand'] = 'ERICSSON'
    elif re.search(r"\bCERAGON\b", blob, re.I):
        info['brand'] = 'CERAGON'
    elif re.search(r"\bSIAE\b", blob, re.I):
        info['brand'] = 'SIAE'

    # Deteksi Model / Subrack Type di NEC
    # Contoh di LCT: "iPASOLINK VR10", "iPASOLINK 400", "iPASOLINK 200", "iPASOLINK 1000", "iPASOLINK EX"
    m_subrack = re.search(r"(?:Subrack\s*Type|Equipment\s*Type|Product\s*Name|Model)[\s:=]*([A-Za-z0-9_\-\s]+)", blob, re.I)
    if m_subrack:
        raw_m = m_subrack.group(1).strip()
        # Ambil token pertama atau frasa model
        clean_m = re.split(r"[\r\n;,]", raw_m)[0].strip()
        if len(clean_m) > 2:
            info['model'] = clean_m

    if not info['model']:
        # Cari kata kunci pasolink di semua token
        for t in tokens:
            m_paso = re.search(r"(i?PASOLINK\s*(?:VR\s*10|VR\s*4|400|200|1000|EX|GX|AX|SX|A|V|I)?)", t.text, re.I)
            if m_paso:
                info['model'] = m_paso.group(1).strip()
                break

    # Format full_name
    if info['brand'] == 'NEC':
        if info['model']:
            if not info['model'].upper().startswith('NEC'):
                info['full_name'] = f"NEC {info['model']}"
            else:
                info['full_name'] = info['model']
        else:
            info['full_name'] = "NEC iPASOLINK"
    else:
        info['full_name'] = f"{info['brand']} {info['model'] or ''}".strip()

    return info


def inspect_microwave_images(
    image_paths: List[str],
    station_name_hint: Optional[str] = None,
    precomputed_tokens: Optional[Dict[str, List[Token]]] = None
) -> List[Dict[str, Any]]:
    """
    Menjalankan OCR RapidOCR dan parsing layout-aware terhadap sekumpulan screenshot
    (Port Setting, Radio Configuration, Equipment Configuration).
    Mengembalikan list data link teridentifikasi dengan 6 parameter lengkap.
    """
    if not MWLINK_AVAILABLE:
        raise RuntimeError("Modul MWLink_Inspector tidak tersedia atau RapidOCR belum terpasang!")

    if not image_paths:
        return []

    tokens_by_file: Dict[str, List[Token]] = {}
    all_tokens: List[Token] = []
    
    if precomputed_tokens is not None:
        tokens_by_file = precomputed_tokens
        for toks in tokens_by_file.values():
            all_tokens.extend(toks)
    else:
        # 1. Jalankan OCR untuk setiap file gambar
        for path in image_paths:
            if not os.path.exists(path):
                continue
            toks = run_ocr(path, backend="rapidocr")
            tokens_by_file[path] = toks
            all_tokens.extend(toks)

    # 2. Deteksi Info Perangkat (Parameter 6: Merk / Tipe)
    equip_info = extract_equipment_device_info(all_tokens)

    # 3. Klasifikasi & Parse halaman Port dan Radio
    port_pages: List[List[Tuple[Optional[int], str]]] = []
    radio_pages: List[RadioPage] = []
    ne_names: List[str] = []
    warnings: List[str] = []

    for path, toks in tokens_by_file.items():
        fname = os.path.basename(path)
        kind = classify(toks)
        
        # Ekstrak NE Name jika ada
        ne = find_ne_name(toks)
        if ne:
            ne_names.append(ne)

        if kind == "port":
            p_page = parse_port_page(toks)
            port_pages.append(p_page.rows)
            warnings.extend([f"{fname}: {w}" for w in p_page.warnings])
            if p_page.ne_name:
                ne_names.append(p_page.ne_name)
        elif kind == "radio":
            r_page = parse_radio_page(toks)
            
            # Perbaiki pembacaan Bandwidth agar mendukung konversi kHz -> MHz
            k = _unit(toks)
            heads = [t for t in toks if HEAD_COL.search(t.text)]
            if heads:
                top = min(t.cy for t in heads)
                heads = sorted((t for t in heads if t.cy - top < 40 * k), key=lambda t: t.cx)
                cols = {i: t.cx for i, t in enumerate(heads)}
                
                # Cari baris Channel Spacing / Bandwidth
                lab_bw = None
                for t in toks:
                    if re.search(r"Channel\s*Spacing|Bandwidth|BW", t.text, re.I) and t.x1 < min(cols.values()):
                        lab_bw = t
                        break
                
                if lab_bw:
                    for t in toks:
                        if t is lab_bw or t.x1 <= lab_bw.x2 or abs(t.cy - lab_bw.cy) > 12 * k:
                            continue
                        i = min(cols, key=lambda c: abs(cols[c] - t.cx))
                        if abs(cols[i] - t.cx) <= 60 * k:
                            parsed_bw = parse_bandwidth_khz_to_mhz(t.text)
                            if parsed_bw is not None:
                                r_page.bw[i] = parsed_bw
            
            radio_pages.append(r_page)
            warnings.extend([f"{fname}: {w}" for w in r_page.warnings])
            if r_page.ne_name:
                ne_names.append(r_page.ne_name)

    # 4. Tentukan STN_Name (Parameter 4)
    detected_stn_name = None
    if ne_names:
        detected_stn_name = ne_names[0]
    elif station_name_hint:
        detected_stn_name = station_name_hint

    # 5. Pasangkan Port Setting dan Radio Configuration
    radio_cols = []
    for pg in radio_pages:
        for i, lab in enumerate(pg.labels):
            bw_val = pg.bw.get(i)
            # Pastikan jika ada nilai bw dalam kHz, dikonversi ke MHz
            if bw_val and bw_val > 500:
                bw_val = round(bw_val / 1000.0, 2)
            radio_cols.append((lab, bw_val, pg.tx.get(i), pg.rx.get(i)))

    flattened_ports = [row for page in port_pages for row in page]

    # Pasangkan baris port dengan kolom radio
    pairs = []
    if flattened_ports and radio_cols:
        pairs = _join(flattened_ports, radio_cols)
    elif radio_cols:
        # Hanya ada halaman radio
        for cl, bw, tx, rx in radio_cols:
            pairs.append((cl, None, bw, tx, rx, []))
    elif flattened_ports:
        # Hanya ada halaman port
        for pl, name in flattened_ports:
            pairs.append((pl, name, None, None, None, []))

    # 6. Susun objek hasil identifikasi parameter teknis
    results: List[Dict[str, Any]] = []
    for idx, (slot_lab, port_name, bw, tx, rx, pair_warnings) in enumerate(pairs):
        clean_lawan = re.sub(r"^(?:TO\s+|TO(?=[0-9A-Z]))", "", port_name).strip() if port_name else None
        
        # Format Tx & Rx
        tx_val = round(float(tx), 2) if tx is not None else None
        rx_val = round(float(rx), 2) if rx is not None else None
        bw_val = round(float(bw), 2) if bw is not None else None

        results.append({
            'slot': slot_lab if slot_lab is not None else (idx + 1),
            'stn_name': detected_stn_name,                     # Parameter 4: NE Name
            'stasiun_lawan': clean_lawan,                      # Parameter 5: Port Name
            'tx_rf': tx_val,                                  # Parameter 1: Tx RF
            'rx_rf': rx_val,                                  # Parameter 2: Rx RF
            'bw_mhz': bw_val,                                 # Parameter 3: Bandwidth (MHz)
            'merk_tipe': equip_info['full_name'],             # Parameter 6: Merk / Tipe
            'source_files': [os.path.basename(p) for p in image_paths],
            'warnings': pair_warnings + warnings
        })

    return results


def match_and_update_station_links(
    db_session: Any,
    station_id: int,
    detected_links: List[Dict[str, Any]],
    user_id: Optional[int] = None
) -> Dict[str, Any]:
    """
    Mencocokkan parameter teknis yang diidentifikasi dari OCR ke tabel stasiun_lawan lokal
    dan langsung MENULISKAN hasilnya ke kolom hasil pemeriksaan:
    - hasil_stn_name
    - hasil_stasiun_lawan
    - hasil_freq_tx
    - hasil_freq_rx
    - hasil_bw
    - hasil_tipe_perangkat
    - status_pemeriksaan_link (Sesuai ISR / Tidak Sesuai Parameter Teknis)
    - tanggal_pemeriksaan
    - metode_pemeriksaan
    """
    try:
        import app
        Stasiun = app.Stasiun
        StasiunLawan = app.StasiunLawan
    except Exception:
        Stasiun = None
        StasiunLawan = None

    station = Stasiun.query.get(station_id)
    if not station:
        return {'success': False, 'message': f'Stasiun ID {station_id} tidak ditemukan'}

    all_lawans = StasiunLawan.query.filter_by(stasiun_id=station_id).all()
    if not all_lawans:
        return {'success': False, 'message': 'Stasiun tidak memiliki data stasiun lawan di database'}

    matched_count = 0
    updated_records = []
    used_lawan_ids = set()

    today_date = date.today()

    for item in detected_links:
        det_lawan = item.get('stasiun_lawan')
        det_tx = item.get('tx_rf')
        det_rx = item.get('rx_rf')
        det_bw = item.get('bw_mhz')
        det_stn = item.get('stn_name') or station.stasiun_name
        det_tipe = item.get('merk_tipe') or 'NEC iPASOLINK'

        best_match = None
        best_score = -1

        # Coba cocokkan dengan link lawan stasiun ini
        for l in all_lawans:
            if l.id in used_lawan_ids:
                continue

            score = 0
            
            # 1. Kecocokan nama lawan
            if det_lawan and l.nama_stasiun_lawan:
                ns = name_score(det_lawan, l.nama_stasiun_lawan)
                if ns >= 75:
                    score += 50 + (ns * 0.3)
                elif norm(det_lawan) in norm(l.nama_stasiun_lawan) or norm(l.nama_stasiun_lawan) in norm(det_lawan):
                    score += 60

            # 2. Kecocokan frekuensi Tx
            if det_tx is not None and l.freq_tx:
                try:
                    l_tx = float(re.search(r"(\d+(?:\.\d+)?)", str(l.freq_tx)).group(1))
                    if abs(det_tx - l_tx) <= 0.5:
                        score += 40
                except Exception:
                    pass

            # 3. Kecocokan frekuensi Rx
            if det_rx is not None and l.freq_rx:
                try:
                    l_rx = float(re.search(r"(\d+(?:\.\d+)?)", str(l.freq_rx)).group(1))
                    if abs(det_rx - l_rx) <= 0.5:
                        score += 30
                except Exception:
                    pass

            if score > best_score and score >= 40:
                best_score = score
                best_match = l

        # Jika cocok, tuliskan hasilnya ke kolom hasil pemeriksaan stasiun lawan
        if best_match:
            used_lawan_ids.add(best_match.id)
            matched_count += 1
            
            # Tentukan status kesesuaian parameter teknis
            status_pemeriksaan = 'Sesuai ISR'
            alasan_beda = []

            # Validasi Tx
            if det_tx is not None and best_match.freq_tx:
                try:
                    isr_tx = float(re.search(r"(\d+(?:\.\d+)?)", str(best_match.freq_tx)).group(1))
                    if abs(det_tx - isr_tx) > 0.5:
                        status_pemeriksaan = 'Tidak Sesuai Parameter Teknis'
                        alasan_beda.append(f"Tx hasil ({det_tx} MHz) != ISR ({isr_tx} MHz)")
                except Exception:
                    pass

            # Validasi Rx
            if det_rx is not None and best_match.freq_rx:
                try:
                    isr_rx = float(re.search(r"(\d+(?:\.\d+)?)", str(best_match.freq_rx)).group(1))
                    if abs(det_rx - isr_rx) > 0.5:
                        status_pemeriksaan = 'Tidak Sesuai Parameter Teknis'
                        alasan_beda.append(f"Rx hasil ({det_rx} MHz) != ISR ({isr_rx} MHz)")
                except Exception:
                    pass

            # Validasi BW
            if det_bw is not None and best_match.bw_isr:
                try:
                    isr_bw = float(re.search(r"(\d+(?:\.\d+)?)", str(best_match.bw_isr)).group(1))
                    if det_bw > isr_bw + 0.1:
                        status_pemeriksaan = 'Tidak Sesuai Parameter Teknis'
                        alasan_beda.append(f"BW hasil ({det_bw} MHz) > ISR ({isr_bw} MHz)")
                except Exception:
                    pass

            # Tulis otomatis ke record StasiunLawan
            best_match.hasil_stn_name = det_stn
            best_match.hasil_stasiun_lawan = det_lawan or best_match.nama_stasiun_lawan
            best_match.hasil_freq_tx = str(det_tx) if det_tx is not None else best_match.freq_tx
            best_match.hasil_freq_rx = str(det_rx) if det_rx is not None else best_match.freq_rx
            best_match.hasil_bw = str(det_bw) if det_bw is not None else (best_match.bw_isr or '')
            
            # Jika di ISR sudah ada tipe perangkat yang lebih spesifik dan cocok dengan NEC, pertahankan/gunakan
            if best_match.tipe_perangkat_isr and 'NEC' in best_match.tipe_perangkat_isr.upper():
                best_match.hasil_tipe_perangkat = best_match.tipe_perangkat_isr
            else:
                best_match.hasil_tipe_perangkat = det_tipe
                
            best_match.hasil_sertifikat = 'Ada'
            best_match.status_pemeriksaan_link = status_pemeriksaan
            best_match.metode_pemeriksaan = 'Pemeriksaan melalui Remote Site'
            best_match.tanggal_pemeriksaan = today_date
            if alasan_beda:
                best_match.keterangan = "; ".join(alasan_beda)

            item['matched_lawan_id'] = best_match.id
            item['matched_lawan_name'] = best_match.nama_stasiun_lawan
            item['status_pemeriksaan'] = status_pemeriksaan

            updated_records.append({
                'lawan_id': best_match.id,
                'lawan_name': best_match.nama_stasiun_lawan,
                'tx': best_match.hasil_freq_tx,
                'rx': best_match.hasil_freq_rx,
                'bw': best_match.hasil_bw,
                'tipe': best_match.hasil_tipe_perangkat,
                'status': status_pemeriksaan
            })

    # Tandai stasiun juga sebagai sudah diperiksa
    if matched_count > 0:
        station.status_pemeriksaan = 'sudah_diperiksa'
        station.checked_at = datetime.utcnow()
        if user_id:
            station.checked_by = user_id

        db_session.commit()

    return {
        'success': True,
        'station_id': station.id,
        'station_name': station.stasiun_name,
        'matched_count': matched_count,
        'total_detected': len(detected_links),
        'updated_records': updated_records
    }


def process_microwave_ocr(image_paths: List[str], station_name_hint: Optional[str] = None) -> Dict[str, Any]:
    """
    Fungsi terpadu untuk pemrosesan OCR gambar microwave link:
    - Menghasilkan teks mentah OCR lengkap (diurutkan posisi vertikal & horizontal)
    - Menghitung rata-rata skor keyakinan (confidence)
    - Mengekstrak 6 parameter microwave link secara terstruktur (Slot, NE Name, Lawan, Tx, Rx, BW, Merk)
    """
    if not MWLINK_AVAILABLE:
        raise RuntimeError("Modul MWLink_Inspector tidak tersedia")

    import time
    t0 = time.time()

    tokens_by_file: Dict[str, List[Token]] = {}
    all_tokens: List[Token] = []

    for path in image_paths:
        if not os.path.exists(path):
            continue
        toks = run_ocr(path, backend="rapidocr")
        tokens_by_file[path] = toks
        all_tokens.extend(toks)

    # Susun teks baris demi baris terurut atas->bawah, kiri->kanan
    sorted_tokens = sorted(all_tokens, key=lambda t: (round(t.cy / 12), t.x1))
    text_lines = [t.text for t in sorted_tokens if t.text]
    full_text = "\n".join(text_lines)

    scores = [t.conf for t in all_tokens if t.conf is not None]
    avg_conf = round((sum(scores) / len(scores)) * 100, 1) if scores else 95.0

    links = inspect_microwave_images(image_paths, station_name_hint=station_name_hint, precomputed_tokens=tokens_by_file)
    elapse = round(time.time() - t0, 2)

    return {
        'text': full_text,
        'confidence': avg_conf,
        'links': links,
        'elapse': elapse
    }
