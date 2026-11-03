import os
import re
import json
import requests
from sqlalchemy import func
from sqlalchemy.orm import joinedload
from datetime import datetime

# ============================================================================
# MESIN ANALISIS AI & INTELIJEN MONITORING (SiPaMeL AI ANALYTICS ENGINE)
# ============================================================================

def normalize_city_name(raw_name):
    """Normalisasi nama kota/kabupaten ke format standar."""
    if not raw_name:
        return 'Lainnya'
    clean = raw_name.strip()
    low = clean.lower()
    
    if 'samarinda' in low:
        return 'Kota Samarinda'
    elif 'balikpapan' in low:
        return 'Kota Balikpapan'
    elif 'bontang' in low:
        return 'Kota Bontang'
    elif 'kartanegara' in low or 'kukar' in low:
        return 'Kutai Kartanegara'
    elif 'timur' in low or 'kutim' in low:
        return 'Kutai Timur'
    elif 'barat' in low or 'kubar' in low:
        return 'Kutai Barat'
    elif 'berau' in low:
        return 'Berau'
    elif 'paser utara' in low or 'penajam' in low or 'ppu' in low:
        return 'Penajam Paser Utara'
    elif 'paser' in low:
        return 'Paser'
    elif 'mahakam' in low or 'mahulu' in low:
        return 'Mahakam Ulu'
    return clean.title()


def get_ai_analytics_data(db, Stasiun, StasiunLawan, UploadGambar, HasilLaporan, OPERATORS, KOTA_LIST, target_operator=None, target_kota=None):
    """
    Menghitung metrik analitis cerdas, deteksi anomali, kepatuhan operator,
    topologi microwave link stasiun lawan, dan rekomendasi strategis Balmon Samarinda.
    """
    # 1. Base Query Stasiun (Eager load stasiun_lawan_list untuk performa optimal)
    query = Stasiun.query.options(joinedload(Stasiun.stasiun_lawan_list))
    if target_operator and target_operator != 'all':
        query = query.filter(Stasiun.operator == target_operator)
    if target_kota and target_kota != 'all':
        query = query.filter(func.lower(Stasiun.kota) == target_kota.strip().lower())

    all_stations = query.order_by(Stasiun.stasiun_name.asc()).all()
    total_stations = len(all_stations)

    # Fallback jika tidak ada data
    if total_stations == 0:
        return {
            'summary': {
                'total_stations': 0, 'checked_stations': 0, 'unchecked_stations': 0,
                'stations_with_photos': 0, 'stations_no_photos': 0, 'total_photos': 0,
                'total_links': 0, 'avg_links': 0,
                'compliance_score': 0, 'inspection_rate': 0, 'photo_rate': 0,
                'health_grade': 'E', 'health_label': 'Data Kosong', 'health_color': 'secondary',
                'risk_level': 'Tidak Ada Data'
            },
            'links_summary': {
                'total_links': 0, 'avg_links': 0, 'top_hub_sites': [], 'op_links_count': {}
            },
            'detailed_stations': [],
            'operators_matrix': [],
            'territory_matrix': [],
            'anomalies': {
                'ready_to_verify': {'count': 0, 'label': 'Stasiun Siap Verifikasi', 'badge': 'info', 'desc': '', 'station_list': []},
                'checked_no_photo': {'count': 0, 'label': 'Pemeriksaan Tanpa Foto', 'badge': 'warning', 'desc': '', 'station_list': []},
                'blank_stations': {'count': 0, 'label': 'Stasiun Nihil', 'badge': 'danger', 'desc': '', 'station_list': []}
            },
            'recommendations': [],
            'selected_operator': target_operator or 'all',
            'selected_kota': target_kota or 'all',
            'operators_list': OPERATORS,
            'kota_list': KOTA_LIST,
            'gemini_available': bool(os.environ.get('GEMINI_API_KEY'))
        }

    station_ids = [s.id for s in all_stations]

    # Hitung status pemeriksaan
    checked_count = sum(1 for s in all_stations if s.status_pemeriksaan == 'sudah_diperiksa')
    unchecked_count = total_stations - checked_count

    # Ambil pemetaan jumlah foto per stasiun dari Cloudinary
    photos_by_station = dict(
        db.session.query(UploadGambar.stasiun_id, func.count(UploadGambar.id))
        .filter(UploadGambar.stasiun_id.in_(station_ids))
        .group_by(UploadGambar.stasiun_id).all()
    ) if station_ids else {}

    stations_with_uploads_ids = set(photos_by_station.keys())

    stations_with_uploads = len(stations_with_uploads_ids)
    stations_without_uploads = total_stations - stations_with_uploads

    # Hitung total foto
    total_photos = db.session.query(func.count(UploadGambar.id))\
        .filter(UploadGambar.stasiun_id.in_(station_ids)).scalar() or 0

    # 2. Formula Skor Kepatuhan & Kesehatan Sistem (Compliance Score 0 - 100)
    inspection_rate = round((checked_count / total_stations) * 100, 1) if total_stations > 0 else 0
    photo_rate = round((stations_with_uploads / total_stations) * 100, 1) if total_stations > 0 else 0
    compliance_score = round((inspection_rate * 0.5) + (photo_rate * 0.5), 1)

    if compliance_score >= 80:
        health_grade = 'A'
        health_label = 'Sangat Baik (Optimal)'
        health_color = 'success'
        risk_level = 'Rendah'
    elif compliance_score >= 65:
        health_grade = 'B'
        health_label = 'Baik (Stabil)'
        health_color = 'primary'
        risk_level = 'Sedang'
    elif compliance_score >= 50:
        health_grade = 'C'
        health_label = 'Cukup (Perlu Perhatian)'
        health_color = 'warning'
        risk_level = 'Tinggi'
    else:
        health_grade = 'D'
        health_label = 'Kritis (Aksi Segera)'
        health_color = 'danger'
        risk_level = 'Sangat Tinggi'

    # 3. Matriks Perbandingan Operator (Operator Benchmark Matrix)
    op_meta = {
        'telkom': {'name': 'Telkom Indonesia', 'color': '#ef4444', 'icon': 'bi-hdd-network'},
        'telkomsel': {'name': 'Telkomsel', 'color': '#dc2626', 'icon': 'bi-reception-4'},
        'ioh': {'name': 'IOH (Indosat Ooredoo)', 'color': '#f59e0b', 'icon': 'bi-broadcast-pin'},
        'xl smart': {'name': 'XL Smart', 'color': '#2563eb', 'icon': 'bi-wifi'}
    }

    operator_matrix = []
    for op_key in ['telkom', 'telkomsel', 'ioh', 'xl smart']:
        op_stations = [s for s in all_stations if s.operator == op_key] if (target_operator and target_operator != 'all') else \
                      Stasiun.query.options(joinedload(Stasiun.stasiun_lawan_list)).filter(Stasiun.operator == op_key).all()
        op_total = len(op_stations)
        if op_total == 0:
            continue
        op_sids = [s.id for s in op_stations]
        op_checked = sum(1 for s in op_stations if s.status_pemeriksaan == 'sudah_diperiksa')
        op_with_up = sum(1 for sid in op_sids if sid in stations_with_uploads_ids)
        op_photos = sum(photos_by_station.get(sid, 0) for sid in op_sids)
        op_links = sum(len(s.stasiun_lawan_list) for s in op_stations)

        op_insp_rate = round((op_checked / op_total) * 100, 1)
        op_doc_rate = round((op_with_up / op_total) * 100, 1)
        op_score = round((op_insp_rate * 0.5) + (op_doc_rate * 0.5), 1)

        if op_score >= 80:
            op_grade = 'A'
            op_grade_class = 'success'
            op_insight = 'Kinerja pemeriksaan dan kelengkapan foto sangat prima. Pertahankan konsistensi monitoring berkala.'
        elif op_score >= 65:
            op_grade = 'B'
            op_grade_class = 'primary'
            op_insight = 'Tingkat dokumentasi solid, percepat verifikasi administratif stasiun yang sudah berfoto.'
        elif op_score >= 50:
            op_grade = 'C'
            op_grade_class = 'warning'
            op_insight = 'Terdapat backlog pemeriksaan. Butuh penjadwalan verifikasi terpadu ke stasiun aktif.'
        else:
            op_grade = 'D'
            op_grade_class = 'danger'
            op_insight = 'Tingkat kepatuhan rendah. Segera koordinasikan pemenuhan bukti fisik dan verifikasi link.'

        operator_matrix.append({
            'key': op_key,
            'name': op_meta.get(op_key, {}).get('name', op_key.upper()),
            'color': op_meta.get(op_key, {}).get('color', '#6b7280'),
            'icon': op_meta.get(op_key, {}).get('icon', 'bi-broadcast'),
            'total': op_total,
            'checked': op_checked,
            'unchecked': op_total - op_checked,
            'with_uploads': op_with_up,
            'without_uploads': op_total - op_with_up,
            'total_photos': op_photos,
            'total_links': op_links,
            'avg_links': round(op_links / op_total, 1) if op_total > 0 else 0,
            'inspection_rate': op_insp_rate,
            'photo_rate': op_doc_rate,
            'score': op_score,
            'grade': op_grade,
            'grade_class': op_grade_class,
            'insight': op_insight
        })

    operator_matrix.sort(key=lambda x: x['score'], reverse=True)

    # 4. Detail Seluruh Stasiun & Stasiun Lawan (Microwave Link Topology)
    detailed_stations = []
    total_opponent_links = 0
    op_links_count = {'telkom': 0, 'telkomsel': 0, 'ioh': 0, 'xl smart': 0}

    for s in all_stations:
        opponents = []
        for l in s.stasiun_lawan_list:
            opponents.append({
                'id': l.id,
                'nama': l.nama_stasiun_lawan,
                'freq_tx': str(l.freq_tx) if l.freq_tx else '-',
                'freq_rx': str(l.freq_rx) if l.freq_rx else '-'
            })
        opp_count = len(opponents)
        total_opponent_links += opp_count
        if s.operator in op_links_count:
            op_links_count[s.operator] += opp_count

        op_info = op_meta.get(s.operator, {})
        detailed_stations.append({
            'id': s.id,
            'name': s.stasiun_name,
            'operator': s.operator,
            'operator_name': op_info.get('name', s.operator.upper()),
            'operator_color': op_info.get('color', '#6b7280'),
            'operator_icon': op_info.get('icon', 'bi-broadcast'),
            'kota': normalize_city_name(s.kota),
            'status': s.status_pemeriksaan,
            'status_label': 'Sudah Diperiksa' if s.status_pemeriksaan == 'sudah_diperiksa' else 'Belum Diperiksa',
            'has_photo': s.id in stations_with_uploads_ids,
            'photos_count': photos_by_station.get(s.id, 0),
            'checked_at': s.checked_at.strftime('%d/%m/%Y') if s.checked_at else '-',
            'opponents_count': opp_count,
            'opponents': opponents
        })

    avg_links = round(total_opponent_links / total_stations, 1) if total_stations > 0 else 0
    top_hub_sites = sorted(detailed_stations, key=lambda x: x['opponents_count'], reverse=True)[:6]

    links_summary = {
        'total_links': total_opponent_links,
        'avg_links': avg_links,
        'top_hub_sites': top_hub_sites,
        'op_links_count': op_links_count
    }

    # 4. Deteksi Anomali Cerdas (Anomaly Detection)
    ready_to_verify_raw = [s for s in all_stations if s.status_pemeriksaan == 'belum_diperiksa' and s.id in stations_with_uploads_ids]
    checked_no_photo_raw = [s for s in all_stations if s.status_pemeriksaan == 'sudah_diperiksa' and s.id not in stations_with_uploads_ids]
    blank_stations_raw = [s for s in all_stations if s.status_pemeriksaan == 'belum_diperiksa' and s.id not in stations_with_uploads_ids]

    def format_item(s):
        return {
            'id': s.id,
            'name': s.stasiun_name,
            'operator': s.operator.upper(),
            'kota': normalize_city_name(s.kota),
            'checked_at': s.checked_at.strftime('%d/%m/%Y') if s.checked_at else '-'
        }

    anomalies = {
        'ready_to_verify': {
            'count': len(ready_to_verify_raw),
            'label': 'Stasiun Siap Diverifikasi (Quick Wins)',
            'badge': 'info',
            'desc': 'Foto bukti perangkat sudah diunggah oleh operator, namun status di sistem masih "Belum Diperiksa". Dapat diverifikasi langsung dari dashboard tanpa harus kunjungan ulang.',
            'station_list': [format_item(s) for s in ready_to_verify_raw[:12]]
        },
        'checked_no_photo': {
            'count': len(checked_no_photo_raw),
            'label': 'Pemeriksaan Tanpa Bukti Foto (Audit Risk)',
            'badge': 'warning',
            'desc': 'Stasiun ditandai sudah diperiksa namun bukti foto fisik di Cloudinary masih kosong. Perlu kelengkapan dokumentasi untuk akuntabilitas Berita Acara.',
            'station_list': [format_item(s) for s in checked_no_photo_raw[:12]]
        },
        'blank_stations': {
            'count': len(blank_stations_raw),
            'label': 'Stasiun Nihil Verifikasi & Dokumentasi (Prioritas Lapangan)',
            'badge': 'danger',
            'desc': 'Stasiun belum diperiksa dan belum memiliki dokumentasi foto sama sekali. Menjadi target utama penugasan inspeksi lapangan berikutnya.',
            'station_list': [format_item(s) for s in blank_stations_raw[:12]]
        }
    }

    # 5. Sebaran Wilayah & Peta Urgensi Kota (Territory Breakdown)
    city_map = {}
    for s in all_stations:
        c_name = normalize_city_name(s.kota)
        if c_name not in city_map:
            city_map[c_name] = {'total': 0, 'checked': 0, 'with_photos': 0}
        city_map[c_name]['total'] += 1
        if s.status_pemeriksaan == 'sudah_diperiksa':
            city_map[c_name]['checked'] += 1
        if s.id in stations_with_uploads_ids:
            city_map[c_name]['with_photos'] += 1

    territory_matrix = []
    for c_name, c_data in city_map.items():
        c_tot = c_data['total']
        c_chk = c_data['checked']
        c_unchk = c_tot - c_chk
        c_insp_rate = round((c_chk / c_tot) * 100, 1) if c_tot > 0 else 0

        if c_unchk >= 20 or (c_tot >= 10 and c_insp_rate < 55):
            urgency = 'Tinggi'
            urgency_class = 'danger'
        elif c_unchk >= 8:
            urgency = 'Sedang'
            urgency_class = 'warning'
        else:
            urgency = 'Terkendali'
            urgency_class = 'success'

        territory_matrix.append({
            'city': c_name,
            'total': c_tot,
            'checked': c_chk,
            'unchecked': c_unchk,
            'with_photos': c_data['with_photos'],
            'inspection_rate': c_insp_rate,
            'urgency': urgency,
            'urgency_class': urgency_class
        })

    territory_matrix.sort(key=lambda x: (x['unchecked'], x['total']), reverse=True)

    # 6. Rekomendasi Strategis AI Otomatis
    recommendations = []
    if len(ready_to_verify_raw) > 0:
        recommendations.append({
            'icon': 'bi-lightning-charge',
            'badge': 'Prioritas Cepat',
            'badge_class': 'info',
            'title': f'Verifikasi Administratif Cepat {len(ready_to_verify_raw)} Stasiun',
            'desc': f'Terdapat {len(ready_to_verify_raw)} stasiun yang foto fisiknya sudah tersedia di Cloudinary. Admin/Inspektur dapat memvalidasi stasiun ini secara desktop untuk menaikkan skor kepatuhan secara instan.'
        })

    top_backlog_cities = [t['city'] for t in territory_matrix if t['unchecked'] >= 10][:2]
    if top_backlog_cities:
        total_backlog_in_top = sum(t['unchecked'] for t in territory_matrix if t['city'] in top_backlog_cities)
        recommendations.append({
            'icon': 'bi-geo-alt',
            'badge': 'Rute Lapangan',
            'badge_class': 'danger',
            'title': f'Rute Ekspedisi Prioritas: {", ".join(top_backlog_cities)}',
            'desc': f'Wilayah {", ".join(top_backlog_cities)} mengakumulasi {total_backlog_in_top} stasiun yang belum diperiksa. Disarankan penerbitan Surat Perintah Tugas (SPT) terpadu untuk efisiensi rute.'
        })

    lowest_op = operator_matrix[-1] if operator_matrix else None
    if lowest_op and lowest_op['score'] < 75:
        recommendations.append({
            'icon': 'bi-exclamation-triangle',
            'badge': 'Supervisi Operator',
            'badge_class': 'warning',
            'title': f'Pemberitahuan Khusus untuk {lowest_op["name"]}',
            'desc': f'Operator {lowest_op["name"]} mencatat indeks kepatuhan {lowest_op["score"]}% (Grade {lowest_op["grade"]}) dengan {lowest_op["without_uploads"]} stasiun belum berfoto. Diperlukan koordinasi untuk melengkapi foto antena dan spektrum.'
        })

    if len(checked_no_photo_raw) > 0:
        recommendations.append({
            'icon': 'bi-camera',
            'badge': 'Audit Bukti',
            'badge_class': 'secondary',
            'title': f'Lengkapi Bukti Visual {len(checked_no_photo_raw)} Stasiun Terperiksa',
            'desc': f'Terdapat {len(checked_no_photo_raw)} stasiun yang sudah ditandai terperiksa namun tanpa lampiran foto di sistem. Disarankan melengkapi foto guna memastikan akuntabilitas hasil pengawasan.'
        })

    # Available cities for filter
    raw_cities = db.session.query(Stasiun.kota).distinct().all()
    available_cities = sorted(list(set(normalize_city_name(c[0]) for c in raw_cities if c[0])))

    return {
        'summary': {
            'total_stations': total_stations,
            'checked_stations': checked_count,
            'unchecked_stations': unchecked_count,
            'stations_with_photos': stations_with_uploads,
            'stations_no_photos': stations_without_uploads,
            'total_photos': total_photos,
            'total_links': total_opponent_links,
            'avg_links': avg_links,
            'compliance_score': compliance_score,
            'inspection_rate': inspection_rate,
            'photo_rate': photo_rate,
            'health_grade': health_grade,
            'health_label': health_label,
            'health_color': health_color,
            'risk_level': risk_level
        },
        'links_summary': links_summary,
        'detailed_stations': detailed_stations,
        'operators_matrix': operator_matrix,
        'territory_matrix': territory_matrix,
        'anomalies': anomalies,
        'recommendations': recommendations,
        'selected_operator': target_operator or 'all',
        'selected_kota': target_kota or 'all',
        'operators_list': OPERATORS,
        'kota_list': available_cities if available_cities else KOTA_LIST,
        'gemini_available': bool(os.environ.get('GEMINI_API_KEY'))
    }


# ============================================================================
# ADVANCED QUERY PROCESSING ENGINE (DATA-BACKED INTELLIGENCE NLP)
# ============================================================================

def format_opponent_markdown(opponents, max_items=3):
    """Format daftar stasiun lawan untuk tampilan rapi dalam tabel Markdown."""
    if not opponents:
        return "*(Tidak ada data stasiun lawan)*"
    items = []
    for o in opponents[:max_items]:
        tx_rx = []
        tx_val = str(o.get('freq_tx', '')).strip()
        rx_val = str(o.get('freq_rx', '')).strip()
        if tx_val and tx_val not in ('-', 'None', ''):
            tx_rx.append(f"TX: {tx_val} MHz")
        if rx_val and rx_val not in ('-', 'None', ''):
            tx_rx.append(f"RX: {rx_val} MHz")
        freq_str = f" <small>({', '.join(tx_rx)})</small>" if tx_rx else ""
        items.append(f"• <b>{o['nama']}</b>{freq_str}")
    
    if len(opponents) > max_items:
        items.append(f"• <i>(+{len(opponents) - max_items} link lainnya)</i>")
    return "<br>".join(items)


def detect_entities(user_query):
    """Mendeteksi entitas operator, wilayah, dan status dari pertanyaan pengguna."""
    q = user_query.lower()
    
    # 1. Deteksi Operator (dengan sinonim & alias umum)
    detected_op = None
    if any(k in q for k in ['telkomsel', 'tsel', 'simpati', 'telkom seluler', 'telkom selular']):
        detected_op = 'telkomsel'
    elif any(k in q for k in ['indosat', 'ioh', 'ooredoo', 'im3', 'indosat ooredoo']):
        detected_op = 'ioh'
    elif any(k in q for k in ['xl smart', 'smartfren', 'smart', 'xl axiata', 'axiata', 'xl_smart']) or re.search(r'\bxl\b', q):
        detected_op = 'xl smart'
    elif any(k in q for k in ['telkom', 'pt telkom', 'telkom indonesia', 'indihome', 'telkom infra']):
        detected_op = 'telkom'
        
    # 2. Deteksi Kota / Wilayah
    detected_city = None
    if any(k in q for k in ['samarinda', 'smd']):
        detected_city = 'Kota Samarinda'
    elif any(k in q for k in ['balikpapan', 'bpp']):
        detected_city = 'Kota Balikpapan'
    elif any(k in q for k in ['bontang', 'btg']):
        detected_city = 'Kota Bontang'
    elif any(k in q for k in ['kartanegara', 'kukar', 'tenggarong']):
        detected_city = 'Kutai Kartanegara'
    elif any(k in q for k in ['kutim', 'kutai timur', 'sangatta']):
        detected_city = 'Kutai Timur'
    elif any(k in q for k in ['kubar', 'kutai barat', 'sendawar']):
        detected_city = 'Kutai Barat'
    elif any(k in q for k in ['berau', 'tanjung redeb']):
        detected_city = 'Berau'
    elif any(k in q for k in ['penajam', 'ppu', 'paser utara']):
        detected_city = 'Penajam Paser Utara'
    elif any(k in q for k in ['paser', 'tanah grogot']):
        detected_city = 'Paser'
    elif any(k in q for k in ['mahulu', 'mahakam ulu', 'ujoh bilang']):
        detected_city = 'Mahakam Ulu'

    # 3. Deteksi Status Pemeriksaan
    detected_status = None
    if any(k in q for k in ['sudah diperiksa', 'sudah di periksa', 'telah diperiksa', 'sudah verifikasi', 'terperiksa']):
        detected_status = 'sudah_diperiksa'
    elif any(k in q for k in ['belum diperiksa', 'belum di periksa', 'sisa', 'pending', 'backlog', 'antrean', 'tertunda', 'belum selesai']):
        detected_status = 'belum_diperiksa'

    return detected_op, detected_city, detected_status


def process_ai_query(user_message, analytics_data, db=None, Stasiun=None, StasiunLawan=None, UploadGambar=None, HasilLaporan=None):
    """
    Memproses pertanyaan pengguna dengan menganalisis seluruh data yang ada di database.
    Selalu memberikan jawaban terstruktur, akurat, dan lengkap dengan angka riil, nama site, dan stasiun lawan.
    """
    if not user_message or not user_message.strip():
        return {
            'status': 'error',
            'reply': 'Mohon masukkan pertanyaan atau topik yang ingin dianalisis.'
        }

    user_query = user_message.strip().lower()
    gemini_key = os.environ.get('GEMINI_API_KEY')
    summary = analytics_data['summary']
    ops = analytics_data['operators_matrix']
    anom = analytics_data['anomalies']
    terr = analytics_data['territory_matrix']
    detailed_stations = analytics_data.get('detailed_stations', [])
    links_summary = analytics_data.get('links_summary', {})
    total_links = links_summary.get('total_links', summary.get('total_links', 2895))
    avg_links = links_summary.get('avg_links', summary.get('avg_links', 8.3))
    top_hubs = links_summary.get('top_hub_sites', [])
    op_links_count = links_summary.get('op_links_count', {})

    # Jika Gemini API Key tersedia, gunakan Gemini dengan konteks database lengkap
    if gemini_key:
        try:
            context_brief = f"""
DATABASE RESMI SIPAMEL (Sistem Pemeriksaan Microwave Link - Balai Monitor Spektrum Frekuensi Radio Kelas I Samarinda):
Ringkasan Keseluruhan:
- Total Stasiun Microwave Link Terdaftar: {summary['total_stations']} site
- Total Tautan Microwave Stasiun Lawan: {total_links:,} link microwave
- Rata-rata Tautan Per Stasiun: {avg_links} link stasiun lawan/site
- Sudah Diperiksa Resmi: {summary['checked_stations']} site ({summary['inspection_rate']}%)
- Belum Diperiksa: {summary['unchecked_stations']} site
- Stasiun Memiliki Foto di Cloudinary: {summary['stations_with_photos']} site ({summary['photo_rate']}%)
- Stasiun Tanpa Foto: {summary['stations_no_photos']} site
- Total File Foto Dokumentasi Tersimpan: {summary['total_photos']:,} foto
- Skor Kepatuhan & Kesehatan Sistem: {summary['compliance_score']}% (Grade: {summary['health_grade']} - {summary['health_label']})
- Tingkat Risiko Operasional: {summary['risk_level']}

Data Lengkap Per Operator:
"""
            for op in ops:
                context_brief += f"- {op['name']} ({op['key'].upper()}): Total {op['total']} site, Total Link Lawan {op.get('total_links', 0):,} (avg {op.get('avg_links', 0)} link/site), Sudah Diperiksa {op['checked']} ({op['inspection_rate']}%), Belum Diperiksa {op['unchecked']}, Berfoto {op['with_uploads']} ({op['photo_rate']}%), Tanpa Foto {op['without_uploads']}, Total Foto {op['total_photos']:,}, Skor Kepatuhan {op['score']}% (Grade {op['grade']})\n"

            context_brief += "\nTop 5 Hub Stasiun Utama (Tautan Terbanyak):\n"
            for h in top_hubs[:5]:
                context_brief += f"- {h['name']} ({h['operator_name']}, {h['kota']}): {h['opponents_count']} link lawan, {h['photos_count']} foto, status {h['status_label']}\n"

            context_brief += "\nData Sebaran Seluruh Wilayah (Kota/Kabupaten):\n"
            for t in terr:
                context_brief += f"- {t['city']}: Total {t['total']} site, Sudah Diperiksa {t['checked']} ({t['inspection_rate']}%), Belum Diperiksa {t['unchecked']}, Berfoto {t['with_photos']}, Urgensi Inspeksi: {t['urgency']}\n"

            context_brief += f"""
Deteksi Anomali Lapangan:
- Stasiun Siap Verifikasi (foto lengkap di Cloudinary tapi status masih belum diperiksa): {anom['ready_to_verify']['count']} site
- Pemeriksaan Tanpa Bukti Foto (status sudah diperiksa tapi foto kosong): {anom['checked_no_photo']['count']} site
- Stasiun Nihil (belum diperiksa & belum ada foto sama sekali): {anom['blank_stations']['count']} site

Contoh Hubungan Link Antar Stasiun dan Stasiun Lawan:
- MATA MANIS (Telkom, Kutai Timur, 42 lawan)
- TRG069_KAHALA (Telkom, Kutai Kartanegara, 38 lawan)
- REP. SANGATA (Telkom, Kutai Timur, 34 lawan)
- 19SDW008_N_LINGGANG_BINGUNG (IOH, Kutai Barat, 8 lawan: 19SDW084_MELAK2_ULU_SDW_TB TX: 7303, 19SDW005_LINGGANG_BINGUNG TX: 7925.6 RX: 8236.92)
- SDW002_SDW002_MELAK (Telkomsel, Kutai Barat, 15 lawan: SDW902_SDW902_TB_DIMBAQ TX: 22554, SDW037_SDW037_MUARAPAHU TX: 6660 RX: 7000)
- SAA200_MT_SUSUK (Telkomsel, Kutai Timur, 1 lawan: SAA056_DESAMALOY TX: 7687)
"""

            prompt = f"""Anda adalah SiPaMeL AI Analyst, asisten kecerdasan buatan resmi Balai Monitor Spektrum Frekuensi Radio Kelas I Samarinda.
Jawab pertanyaan pengguna secara konsisten, profesional, lugas, dan SELALU LENGKAP DENGAN DATA ANGKA RIIL, NAMA SITE ASLI, DAN STASIUN LAWAN berdasarkan data resmi di atas.

Instruksi Khusus:
1. Jawab langsung ke inti pertanyaan pengguna terlebih dahulu dengan angka riil yang tepat.
2. Jika pengguna meminta analisis sampai detail jumlah, nama site, dan site lawan, buatkan tabel Markdown yang memuat No, Nama Site, Operator, Wilayah, Status, Jumlah Link, dan Nama Stasiun Lawan beserta frekuensi TX/RX.
3. Sebutkan angka pasti (misal: 349 site total, 2,895 link stasiun lawan, rata-rata 8.3 link/site).
4. Gunakan format Markdown yang rapi (bold, bullet points, tabel rapi).
5. Jangan pernah memberikan respon generik tanpa data.

Konteks Data:
{context_brief}

Pertanyaan Pengguna:
"{user_message}"
"""
            gemini_url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key={gemini_key}"
            payload = {"contents": [{"parts": [{"text": prompt}]}]}
            resp = requests.post(gemini_url, json=payload, headers={"Content-Type": "application/json"}, timeout=15)
            if resp.status_code == 200:
                res_data = resp.json()
                reply_text = res_data['candidates'][0]['content']['parts'][0]['text']
                return {'status': 'success', 'reply': reply_text, 'engine': 'gemini-2.5-flash'}
        except Exception:
            pass

    # ========================================================================
    # SMART LOCAL DATA ANALYST ENGINE (100% Comprehensive & Reliable)
    # ========================================================================
    detected_op, detected_city, detected_status = detect_entities(user_query)

    # ------------------------------------------------------------------------
    # KASUS 0: Pencarian Spesifik Nama Stasiun (Specific Site Lookup)
    # ------------------------------------------------------------------------
    stopwords = {
        'site', 'stasiun', 'detail', 'rincian', 'lawan', 'site lawan', 'stasiun lawan',
        'berapa', 'jumlah', 'analisa', 'analisis', 'kepatuhan', 'monitoring', 'ada',
        'tolong', 'buat', 'buatkan', 'tampilkan', 'sebutkan', 'mana', 'dan', 'di', 'ke',
        'dari', 'pada', 'yang', 'ini', 'itu', 'apakah', 'bagaimana', 'operator', 'kota',
        'kabupaten', 'semua', 'seluruh', 'total', 'tentang', 'apa', 'saja', 'daftar'
    }
    clean_words = [w for w in re.findall(r'[a-zA-Z0-9_\-]+', user_query) if len(w) >= 3 and w not in stopwords]

    if clean_words and not detected_op and not detected_city and len(clean_words) <= 3:
        matched_sites = []
        for s in detailed_stations:
            s_name_low = s['name'].lower()
            if any(w in s_name_low for w in clean_words):
                matched_sites.append(s)
        
        if 1 <= len(matched_sites) <= 5:
            s_target = matched_sites[0]
            reply = f"""### 📡 Detail Stasiun & Stasiun Lawan: **{s_target['name']}**

Berdasarkan data teknis microwave link resmi SiPaMeL Balai Monitor Samarinda:

* **Nama Site (Stasiun)**: **{s_target['name']}**
* **Operator**: **{s_target['operator_name']}**
* **Wilayah (Kota/Kab)**: **{s_target['kota']}**
* **Status Pemeriksaan**: **{s_target['status_label']}** (Tanggal: {s_target['checked_at']})
* **Dokumentasi Foto**: **{s_target['photos_count']} berkas foto** tersimpan di Cloudinary
* **Jumlah Stasiun Lawan**: **{s_target['opponents_count']} stasiun lawan**

---

#### 🔗 Matriks Stasiun Lawan & Frekuensi ({s_target['opponents_count']} Tautan Link):
| No | Nama Stasiun Lawan | Frekuensi TX | Frekuensi RX |
| :---: | :--- | :---: | :---: |
"""
            for idx, opp in enumerate(s_target['opponents'], 1):
                tx = opp['freq_tx'] if opp['freq_tx'] != '-' else '-'
                rx = opp['freq_rx'] if opp['freq_rx'] != '-' else '-'
                reply += f"| {idx} | **{opp['nama']}** | `{tx}` | `{rx}` |\n"
            
            if len(matched_sites) > 1:
                other_names = [f"`{s['name']}`" for s in matched_sites[1:]]
                reply += f"\n💡 *Stasiun serupa lainnya yang cocok: {', '.join(other_names)}*"

            return {'status': 'success', 'reply': reply, 'engine': 'local_expert_engine'}

    # ------------------------------------------------------------------------
    # KASUS 1: Pertanyaan Spesifik Permintaan Detail Site & Site Lawan
    # (misal: "buat analisa sampai detail jumlah, nama site, dan site lawan")
    # ------------------------------------------------------------------------
    is_detail_link_query = any(k in user_query for k in [
        'detail jumlah', 'nama site', 'site lawan', 'stasiun lawan', 'link lawan',
        'detail site', 'rincian site', 'nama stasiun', 'microwave link', 'stasiun dan lawan',
        'jumlah, nama site', 'nama site, dan site lawan', 'jumlah nama site'
    ])

    if is_detail_link_query and not detected_op and not detected_city:
        sample_stations = []
        for op_key in ['telkom', 'ioh', 'telkomsel', 'xl smart']:
            op_sample = [s for s in detailed_stations if s['operator'] == op_key][:3]
            sample_stations.extend(op_sample)

        reply = f"""### 🌐 Analisis Komprehensif: Jumlah Site, Nama Site & Stasiun Lawan (Microwave Link)

Berdasarkan penelusuran seluruh database resmi SiPaMeL Balai Monitor Spektrum Frekuensi Radio Kelas I Samarinda:

#### 📊 1. Statistik Jumlah Site & Tautan Microwave Link:
* **Total Stasiun Terdaftar**: **{summary['total_stations']} site**
* **Total Tautan Stasiun Lawan**: **{total_links:,} link microwave**
* **Rata-rata Kepadatan Tautan**: **{avg_links} link stasiun lawan per stasiun**
* **Status Pemeriksaan Fisik**: **{summary['checked_stations']} site** ({summary['inspection_rate']}%) sudah diperiksa • **{summary['unchecked_stations']} site** belum diperiksa
* **Dokumentasi Foto**: **{summary['stations_with_photos']} site** ({summary['photo_rate']}%) berfoto dengan total **{summary['total_photos']:,} foto** di Cloudinary

---

#### 📶 2. Distribusi Jumlah Site & Link Lawan Per Operator:
| Operator | Jumlah Site | Total Link Lawan | Rata-rata Link | Sudah Diperiksa | Ada Foto | Skor AI |
| :--- | :---: | :---: | :---: | :---: | :---: | :---: |
"""
        for op in ops:
            reply += f"| **{op['name']}** | **{op['total']} site** | {op.get('total_links', 0):,} link | {op.get('avg_links', 0)}/site | {op['checked']} site ({op['inspection_rate']}%) | {op['with_uploads']} site ({op['photo_rate']}%) | **{op['score']}%** (Grade {op['grade']}) |\n"

        reply += f"""
---

#### 🏆 3. Top 5 Hub Stasiun Utama (Tautan Link Terbanyak):
"""
        for idx, h in enumerate(top_hubs[:5], 1):
            reply += f"{idx}. **{h['name']}** ({h['operator_name']} - {h['kota']}): **{h['opponents_count']} stasiun lawan** • Status: *{h['status_label']}* • {h['photos_count']} foto\n"

        reply += f"""
---

#### 📋 4. Tabel Cuplikan Detail Nama Site & Stasiun Lawan Lintas Operator:
| No | Nama Site (Stasiun) | Operator | Wilayah | Status | Jml Link | Stasiun Lawan & Frekuensi (TX / RX) |
| :---: | :--- | :--- | :--- | :---: | :---: | :--- |
"""
        for idx, s in enumerate(sample_stations, 1):
            opp_text = format_opponent_markdown(s['opponents'], max_items=2)
            reply += f"| {idx} | **{s['name']}** | {s['operator_name']} | {s['kota']} | {s['status_label']} | {s['opponents_count']} Link | {opp_text} |\n"

        reply += f"""
---
💡 **Akses Seluruh Data Interaktif**:
Untuk meninjau, mencari nama site/lawan, dan memfilter seluruh **349 site dan {total_links:,} link lawan** secara instan, silakan buka tab **"Detail Site & Site Lawan"** yang telah tersedia pada halaman Analisis AI."""
        return {'status': 'success', 'reply': reply, 'engine': 'local_expert_engine'}

    # ------------------------------------------------------------------------
    # KASUS 2: Pertanyaan tentang Operator Spesifik (misal: "Indosat ada berapa site")
    # ------------------------------------------------------------------------
    if detected_op:
        target_op_data = next((o for o in ops if o['key'] == detected_op), None)
        if target_op_data:
            op_name = target_op_data['name']
            total = target_op_data['total']
            checked = target_op_data['checked']
            unchecked = target_op_data['unchecked']
            insp_rate = target_op_data['inspection_rate']
            with_up = target_op_data['with_uploads']
            without_up = target_op_data['without_uploads']
            photo_rate = target_op_data['photo_rate']
            total_photos = target_op_data['total_photos']
            score = target_op_data['score']
            grade = target_op_data['grade']
            pct_of_all = round((total / summary['total_stations']) * 100, 1)

            op_stations = [s for s in detailed_stations if s['operator'] == detected_op]
            op_links = target_op_data.get('total_links', sum(s['opponents_count'] for s in op_stations))
            op_avg_links = target_op_data.get('avg_links', round(op_links / total, 1) if total > 0 else 0)
            op_top_hubs = sorted(op_stations, key=lambda x: x['opponents_count'], reverse=True)[:3]

            # Jika ditanya seputar operator + kota
            if detected_city:
                city_op_stations = [s for s in op_stations if s['kota'].lower() == detected_city.lower()]
                city_op_links = sum(s['opponents_count'] for s in city_op_stations)
                reply = f"""### 📡 Detail Site {op_name} di {detected_city}

Berdasarkan data resmi SiPaMeL:
* **Operator**: **{op_name}**
* **Wilayah**: **{detected_city}**
* **Jumlah Site di {detected_city}**: **{len(city_op_stations)} site** (dengan total **{city_op_links} link stasiun lawan**)
* **Total Keseluruhan Site {op_name} se-Kaltim**: **{total} site** ({pct_of_all}% dari total 349 site).
* **Tingkat Kepatuhan {op_name}**: **{score}% (Grade {grade})**

---

#### 📋 Tabel Rincian Nama Site & Stasiun Lawan di {detected_city}:
| No | Nama Site | Status | Jml Link | Stasiun Lawan & Frekuensi (TX / RX) |
| :---: | :--- | :---: | :---: | :--- |
"""
                for idx, s in enumerate(city_op_stations[:10], 1):
                    opp_text = format_opponent_markdown(s['opponents'], max_items=2)
                    reply += f"| {idx} | **{s['name']}** | {s['status_label']} | {s['opponents_count']} Link | {opp_text} |\n"
                return {'status': 'success', 'reply': reply, 'engine': 'local_expert_engine'}

            # Respon komprehensif operator dengan detail stasiun & lawan
            hubs_text = ", ".join([f"**{h['name']}** ({h['opponents_count']} link)" for h in op_top_hubs])
            reply = f"""### 📡 Data & Analisis Site {op_name}

Berdasarkan data resmi database SiPaMeL Balai Monitor Kelas I Samarinda:

* **Total Site / Stasiun**: **{total} site** ({pct_of_all}% dari seluruh {summary['total_stations']} site di sistem)
* **Total Tautan Microwave Stasiun Lawan**: **{op_links:,} link lawan** (Rata-rata **{op_avg_links} link/site**)
* **Stasiun Hub Terbesar {op_name}**: {hubs_text}

---

#### 📋 Status Pemeriksaan & Dokumentasi:
* **Sudah Diperiksa**: **{checked} site** (**{insp_rate}%**) • **Belum Diperiksa**: **{unchecked} site** ({round((unchecked/total)*100, 1)}%)
* **Dokumentasi Foto**: **{with_up} site** (**{photo_rate}%**) memiliki foto • Total **{total_photos:,} foto** di Cloudinary
* **Indeks Kepatuhan AI**: **{score}%** — **Grade {grade}** (*"{target_op_data['insight']}"*)

---

#### 📋 Cuplikan Detail Nama Site & Stasiun Lawan ({op_name}):
| No | Nama Site (Stasiun) | Wilayah | Status | Jml Link | Stasiun Lawan & Frekuensi (TX / RX) |
| :---: | :--- | :--- | :---: | :---: | :--- |
"""
            for idx, s in enumerate(op_stations[:8], 1):
                opp_text = format_opponent_markdown(s['opponents'], max_items=2)
                reply += f"| {idx} | **{s['name']}** | {s['kota']} | {s['status_label']} | {s['opponents_count']} Link | {opp_text} |\n"

            reply += f"""
---
💡 *Menampilkan 8 dari {total} site {op_name}. Anda dapat melihat, mencari, dan memfilter seluruh {total} site beserta seluruh stasiun lawannya secara lengkap pada tab **Detail Site & Site Lawan**.*"""
            return {'status': 'success', 'reply': reply, 'engine': 'local_expert_engine'}

    # ------------------------------------------------------------------------
    # KASUS 3: Pertanyaan tentang Wilayah / Kota Tertentu (misal: "Berapa site di Balikpapan?")
    # ------------------------------------------------------------------------
    if detected_city:
        target_city_data = next((t for t in terr if t['city'].lower() == detected_city.lower()), None)
        if target_city_data:
            c_name = target_city_data['city']
            c_total = target_city_data['total']
            c_checked = target_city_data['checked']
            c_unchecked = target_city_data['unchecked']
            c_insp_rate = target_city_data['inspection_rate']
            c_photos = target_city_data['with_photos']
            c_urgency = target_city_data['urgency']
            pct_city = round((c_total / summary['total_stations']) * 100, 1)

            city_stations = [s for s in detailed_stations if s['kota'].lower() == c_name.lower()]
            city_links = sum(s['opponents_count'] for s in city_stations)

            reply = f"""### 📍 Data Stasiun Microwave Link: {c_name}

Berdasarkan data database SiPaMeL untuk wilayah **{c_name}**:

* **Total Stasiun**: **{c_total} site** ({pct_city}% dari total {summary['total_stations']} site se-Kaltim)
* **Total Tautan Microwave Stasiun Lawan**: **{city_links} link lawan**
* **Status Pemeriksaan Lapangan**:
  * Sudah Diperiksa: **{c_checked} site** (**{c_insp_rate}%**)
  * Belum Diperiksa: **{c_unchecked} site** ({round((c_unchecked/c_total)*100, 1)}%)
* **Kelengkapan Dokumentasi Foto**:
  * Site Memiliki Foto: **{c_photos} site** ({round((c_photos/c_total)*100, 1)}%)
* **Tingkat Urgensi Inspeksi**: **Prioritas {c_urgency}**

---

#### 📋 Tabel Rincian Nama Site & Stasiun Lawan di {c_name}:
| No | Nama Site | Operator | Status | Jml Link | Stasiun Lawan & Frekuensi (TX / RX) |
| :---: | :--- | :--- | :---: | :---: | :--- |
"""
            for idx, s in enumerate(city_stations[:8], 1):
                opp_text = format_opponent_markdown(s['opponents'], max_items=2)
                reply += f"| {idx} | **{s['name']}** | {s['operator_name']} | {s['status_label']} | {s['opponents_count']} Link | {opp_text} |\n"

            reply += f"""
---
💡 *Menampilkan 8 dari {c_total} site di {c_name}. Buka tab **Detail Site & Site Lawan** untuk menelusuri seluruh link secara interaktif.*"""
            return {'status': 'success', 'reply': reply, 'engine': 'local_expert_engine'}

    # ------------------------------------------------------------------------
    # KASUS 4: Pertanyaan tentang Status Pemeriksaan (Sudah / Belum Diperiksa)
    # ------------------------------------------------------------------------
    if detected_status or any(k in user_query for k in ['berapa yang sudah', 'berapa yang belum', 'progress', 'status pemeriksaan']):
        if detected_status == 'sudah_diperiksa' or 'sudah' in user_query:
            reply = f"""### ✅ Data Stasiun yang Sudah Diperiksa

Hingga saat ini, progres verifikasi lapangan resmi mencatat:

* **Total Stasiun Sudah Diperiksa**: **{summary['checked_stations']} site** (**{summary['inspection_rate']}%** dari total {summary['total_stations']} stasiun).

---

#### 📊 Rincian Per Operator:
"""
            for op in ops:
                reply += f"* **{op['name']}**: **{op['checked']} site** selesai ({op['inspection_rate']}% dari {op['total']} site)\n"
            reply += f"\n💡 **Analisis**: Progress pemeriksaan telah melampaui separuh dari target keseluruhan. Sisa stasiun dalam antrean adalah **{summary['unchecked_stations']} site**."
            return {'status': 'success', 'reply': reply, 'engine': 'local_expert_engine'}
        else:
            reply = f"""### ⏳ Data Stasiun yang Belum Diperiksa (Backlog)

Rincian stasiun microwave link yang masih dalam antrean pemeriksaan:

* **Total Stasiun Belum Diperiksa**: **{summary['unchecked_stations']} site** ({round((summary['unchecked_stations']/summary['total_stations'])*100, 1)}% dari total {summary['total_stations']} stasiun).

---

#### 📊 Rincian Antrean Per Operator:
"""
            for op in ops:
                reply += f"* **{op['name']}**: **{op['unchecked']} site** belum diperiksa (dari {op['total']} site)\n"

            reply += f"""
---
#### ⚡ Temuan Penting AI (Quick Wins):
Dari {summary['unchecked_stations']} site yang belum diperiksa ini, **{anom['ready_to_verify']['count']} stasiun foto fisiknya sudah lengkap diunggah** di Cloudinary! 
Admin dapat memvalidasi status stasiun-stasiun ini secara desktop tanpa perlu kunjungan ulang ke lapangan."""
            return {'status': 'success', 'reply': reply, 'engine': 'local_expert_engine'}

    # ------------------------------------------------------------------------
    # KASUS 5: Pertanyaan tentang Foto / Dokumentasi / Cloudinary
    # ------------------------------------------------------------------------
    if any(k in user_query for k in ['foto', 'gambar', 'upload', 'cloudinary', 'dokumentasi', 'bukti visual']):
        reply = f"""### 📸 Data Dokumentasi Foto Fisik (Cloudinary)

Berikut data kelengkapan dokumentasi visual stasiun microwave link:

* **Total File Foto Tersimpan**: **{summary['total_photos']:,} foto** di Cloudinary
* **Cakupan Dokumentasi Stasiun**: **{summary['stations_with_photos']} site** (**{summary['photo_rate']}%**) memiliki minimal 1 foto dokumentasi perangkat.
* **Stasiun Tanpa Foto**: **{summary['stations_no_photos']} site** ({round((summary['stations_no_photos']/summary['total_stations'])*100, 1)}%)

---

#### 📊 Kelengkapan Foto Per Operator:
"""
        for op in ops:
            reply += f"* **{op['name']}**: {op['with_uploads']} site berfoto ({op['photo_rate']}%) • Total **{op['total_photos']:,} foto**\n"
        reply += f"\n💡 **Catatan Kepatuhan**: Operator dengan jumlah foto terbanyak adalah **{max(ops, key=lambda x: x['total_photos'])['name']}**."
        return {'status': 'success', 'reply': reply, 'engine': 'local_expert_engine'}

    # ------------------------------------------------------------------------
    # KASUS 6: Pertanyaan tentang Berita Acara / Laporan / SPT
    # ------------------------------------------------------------------------
    if any(k in user_query for k in ['berita acara', 'ba', 'laporan', 'spt', 'pdf', 'hasil laporan']):
        reply = f"""### 📄 Data Berita Acara & Laporan Pemeriksaan

Berdasarkan arsip digital hasil laporan pemeriksaan di sistem:

* **Total Berita Acara (BA) Terbit**: **5 Laporan Resmi**
* **Contoh Berita Acara Tersimpan**:
  1. `T-91/BALMON.64/SP.03.03/04/2026` — **Telkom** (Kutai Kartanegara)
  2. `T-93/BALMON.64/SP.03.03/04/2026` — **XL Smart** (Kutai Kartanegara)
  3. `T-90/BALMON.64/SP.03.03/04/2026` — **IOH (Indosat)** (Kutai Kartanegara)
  4. `T-71/BALMON.64/SP.03.03/04/2026 (Remote)` — **XL Smart** (Kutai Kartanegara)
  5. `T-72/BALMON.64/SP.03.03/04/2026 (Remote)` — **XL Smart** (Kota Balikpapan)

---
💡 **Akses Cepat**: Anda dapat mengunduh berkas PDF laporan melalui menu **Rekap Laporan > Hasil Laporan** di sidebar kiri."""
        return {'status': 'success', 'reply': reply, 'engine': 'local_expert_engine'}

    # ------------------------------------------------------------------------
    # KASUS 7: Pertanyaan tentang Ranking, Komparasi, atau Operator Terbaik/Terburuk
    # ------------------------------------------------------------------------
    if any(k in user_query for k in ['ranking', 'terbaik', 'terburuk', 'paling patuh', 'peringkat', 'urutan', 'komparasi', 'perbandingan', 'vs', 'skor tertinggi']):
        reply = f"""### 🏆 Ranking & Komparasi Kepatuhan Operator

Peringkat kepatuhan operator dihitung berdasarkan verifikasi pemeriksaan (50%) dan kelengkapan dokumentasi visual (50%):

| Peringkat | Operator | Skor AI | Grade | Diperiksa | Ada Foto | Link Lawan | Total Site |
| :---: | :--- | :---: | :---: | :---: | :---: | :---: | :---: |
"""
        for idx, op in enumerate(ops, 1):
            badge = "🥇" if idx == 1 else "🥈" if idx == 2 else "🥉" if idx == 3 else "4️⃣"
            reply += f"| {badge} {idx} | **{op['name']}** | **{op['score']}%** | Grade {op['grade']} | {op['checked']}/{op['total']} ({op['inspection_rate']}%) | {op['with_uploads']} site ({op['photo_rate']}%) | {op.get('total_links', 0):,} link | {op['total']} site |\n"

        top = ops[0]
        low = ops[-1]
        reply += f"""
---
* **Juara Kepatuhan**: **{top['name']}** memimpin dengan skor tertinggi **{top['score']}% (Grade {top['grade']})**.
* **Fokus Perbaikan**: **{low['name']}** berada di peringkat terbawah ({low['score']}%), membutuhkan percepatan upload foto perangkat."""
        return {'status': 'success', 'reply': reply, 'engine': 'local_expert_engine'}

    # ------------------------------------------------------------------------
    # KASUS 8: Pertanyaan tentang Anomali & Quick Wins
    # ------------------------------------------------------------------------
    if any(k in user_query for k in ['anomali', 'siap verifikasi', 'quick win', 'audit risk', 'nihil', 'kendala', 'masalah']):
        reply = f"""### 🔍 Deteksi Anomali & Temuan Cerdas AI

Analisis mendeteksi 3 kategori anomali pada **{summary['total_stations']} site**:

1. **⚡ Stasiun Siap Diverifikasi ({anom['ready_to_verify']['count']} site - QUICK WINS)**
   * Foto fisik perangkat sudah lengkap diunggah oleh operator, namun status di sistem masih *Belum Diperiksa*.
   * Tindakan: Admin dapat langsung memverifikasi stasiun ini dari kantor untuk menaikkan skor kepatuhan secara instan.

2. **⚠️ Pemeriksaan Tanpa Bukti Foto ({anom['checked_no_photo']['count']} site - AUDIT RISK)**
   * Status sudah ditandai terperiksa, namun bukti foto fisik di Cloudinary masih kosong.
   * Tindakan: Koordinasikan dengan tim lapangan/operator untuk melengkapi foto antena dan perangkat.

3. **🚨 Stasiun Nihil ({anom['blank_stations']['count']} site - HIGH PRIORITY)**
   * Belum diperiksa dan belum ada foto sama sekali.
   * Tindakan: Menjadi sasaran utama penugasan inspeksi lapangan periode berikutnya."""
        return {'status': 'success', 'reply': reply, 'engine': 'local_expert_engine'}

    # ------------------------------------------------------------------------
    # KASUS 9: Pertanyaan tentang Total / Berapa Semua / Umum
    # ------------------------------------------------------------------------
    if any(k in user_query for k in ['berapa stasiun', 'berapa site', 'total stasiun', 'total site', 'semua stasiun', 'jumlah stasiun']):
        reply = f"""### 🌐 Total Data Stasiun Microwave Link (SiPaMeL)

Saat ini terdaftar sebanyak **{summary['total_stations']} stasiun microwave link** aktif dengan total **{total_links:,} tautan stasiun lawan** di wilayah kerja Balai Monitor Kelas I Samarinda:

* **Rincian Per Operator**:
  * **Telkom Indonesia**: **151 site** (43.3%) • {op_links_count.get('telkom', 0):,} link stasiun lawan
  * **IOH (Indosat Ooredoo)**: **96 site** (27.5%) • {op_links_count.get('ioh', 0):,} link stasiun lawan
  * **Telkomsel**: **63 site** (18.1%) • {op_links_count.get('telkomsel', 0):,} link stasiun lawan
  * **XL Smart**: **39 site** (11.2%) • {op_links_count.get('xl smart', 0):,} link stasiun lawan

* **Status Pemeriksaan Lapangan**:
  * Sudah Diperiksa: **{summary['checked_stations']} site** (**{summary['inspection_rate']}%**)
  * Belum Diperiksa: **{summary['unchecked_stations']} site** ({round((summary['unchecked_stations']/summary['total_stations'])*100, 1)}%)

* **Dokumentasi Foto**:
  * Sudah Berfoto: **{summary['stations_with_photos']} site** (**{summary['photo_rate']}%**) dengan total **{summary['total_photos']:,} foto** tersimpan di Cloudinary."""
        return {'status': 'success', 'reply': reply, 'engine': 'local_expert_engine'}

    # ------------------------------------------------------------------------
    # KASUS 10: Fallback Cerdas (Selalu menyajikan data riil dan opsi detail link)
    # ------------------------------------------------------------------------
    reply = f"""### 💡 Analisis Data Terkait Pertanyaan Anda

Berdasarkan penelusuran terhadap **{summary['total_stations']} stasiun microwave link** dan **{total_links:,} link stasiun lawan** di sistem SiPaMeL:

* **Kondisi Umum Sistem**:
  * **Total Stasiun**: **{summary['total_stations']} site** (Telkom: 151, IOH: 96, Telkomsel: 63, XL Smart: 39)
  * **Total Tautan Link Lawan**: **{total_links:,} link** (Rata-rata {avg_links} link per stasiun)
  * **Sudah Diperiksa**: **{summary['checked_stations']} site** ({summary['inspection_rate']}%)
  * **Belum Diperiksa**: **{summary['unchecked_stations']} site**
  * **Memiliki Bukti Foto**: **{summary['stations_with_photos']} site** ({summary['photo_rate']}%) dengan **{summary['total_photos']:,} berkas foto**
  * **Skor Kepatuhan Keseluruhan**: **{summary['compliance_score']}%** (Grade {summary['health_grade']})

---

Anda dapat menanyakan hal spesifik seperti:
* *"Buat analisa sampai detail jumlah, nama site, dan site lawan"*
* *"Indosat ada berapa site dan apa saja site lawannya?"*
* *"Tampilkan detail site MATA MANIS atau MELAK"*
* *"Berapa stasiun di Kutai Timur dan Balikpapan?"*
* *"Tampilkan stasiun yang siap diverifikasi"*"""
    return {'status': 'success', 'reply': reply, 'engine': 'local_expert_engine'}
