import sys
import os

# Limit CPU threads for ONNX runtime & math libraries to prevent memory exhaustion / cgroups throttling in cloud containers
os.environ["OMP_NUM_THREADS"] = "1"
os.environ["OPENBLAS_NUM_THREADS"] = "1"
os.environ["MKL_NUM_THREADS"] = "1"
os.environ["VECLIB_MAXIMUM_THREADS"] = "1"
os.environ["NUMEXPR_NUM_THREADS"] = "1"

if sys.platform == 'win32':
    if hasattr(sys.stdout, 'reconfigure'):
        sys.stdout.reconfigure(encoding='utf-8', errors='replace')
    if hasattr(sys.stderr, 'reconfigure'):
        sys.stderr.reconfigure(encoding='utf-8', errors='replace')

# Import RapidOCR early to initialize ONNX runtime & OpenMP threads
try:
    from rapidocr_onnxruntime import RapidOCR
    RAPID_OCR_AVAILABLE = True
except Exception as _e:
    RAPID_OCR_AVAILABLE = False
    print(f"RapidOCR not available: {_e}")

from flask import Flask, render_template, request, redirect, url_for, flash, jsonify, send_from_directory, session, current_app, send_file, make_response
from flask_login import LoginManager, login_user, logout_user, login_required, current_user
from werkzeug.security import generate_password_hash, check_password_hash
from werkzeug.utils import secure_filename
from flask_sqlalchemy import SQLAlchemy
from flask_login import UserMixin
from datetime import datetime
import os
import json
import uuid
from functools import wraps
import pandas as pd  
from openpyxl import Workbook, load_workbook  
from openpyxl.styles import Alignment, PatternFill, Font, Border, Side  
from io import BytesIO  
from sqlalchemy import extract
import cloudinary
import cloudinary.uploader
import cloudinary.api
from cloudinary.utils import cloudinary_url
from dotenv import load_dotenv
import urllib.parse
from sqlalchemy import create_engine
from sqlalchemy.pool import NullPool
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
import requests
import tempfile
import zipfile
from docx.enum.section import WD_ORIENT
import re
from ai_analytics import get_ai_analytics_data, process_ai_query

# Load environment variables
load_dotenv()

# ============================================================================
# KONFIGURASI APLIKASI
# ============================================================================

app = Flask(__name__)
app.secret_key = 'kunci-rahasia-stasiun-monitoring-2026'

# ============================================================================
# KONFIGURASI DATABASE (POSTGRESQL DI VPS)
# ============================================================================

basedir = os.path.abspath(os.path.dirname(__file__))

# Ambil DATABASE_URL dari environment variable
database_url = os.environ.get('DATABASE_URL')

if not database_url:
    # Fallback jika tidak ada di environment
    database_url = 'postgresql://flask_user:rahasia123@localhost:5432/station_monitoring'
    print("⚠️ DATABASE_URL tidak ditemukan, menggunakan default")

app.config['SQLALCHEMY_DATABASE_URI'] = database_url
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False
app.config['SQLALCHEMY_ENGINE_OPTIONS'] = {
    'pool_size': 10,
    'pool_recycle': 300,
    'pool_pre_ping': True,
}

print(f"✅ Using PostgreSQL database: {database_url.split('@')[1] if '@' in database_url else 'localhost'}")

app.config['UPLOAD_FOLDER'] = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'static', 'uploads')
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB
app.config['ALLOWED_EXTENSIONS'] = {'png', 'jpg', 'jpeg', 'gif', 'webp'}

app.config['SESSION_COOKIE_SECURE'] = False  # Set True jika pakai HTTPS
app.config['SESSION_COOKIE_HTTPONLY'] = True
app.config['SESSION_COOKIE_SAMESITE'] = 'Lax'
app.config['PERMANENT_SESSION_LIFETIME'] = 3600  # Session lifetime 1 jam

# Custom flash categories dengan CSS class
FLASH_CATEGORIES = {
    'success': 'alert-success',
    'error': 'alert-danger',
    'warning': 'alert-warning',
    'info': 'alert-info',
    'debug': 'alert-secondary'
}
# ============================================================================
# KONFIGURASI CLOUDINARY
# ============================================================================

# Konfigurasi Cloudinary dari environment variables
cloudinary.config(
    cloud_name=os.environ.get('CLOUDINARY_CLOUD_NAME'),
    api_key=os.environ.get('CLOUDINARY_API_KEY'),
    api_secret=os.environ.get('CLOUDINARY_API_SECRET'),
    secure=True  # Gunakan HTTPS
)

print("=" * 60)
print("✅ Cloudinary configured")
print(f"   Cloud Name: {os.environ.get('CLOUDINARY_CLOUD_NAME', 'NOT SET')}")
print("=" * 60)

# ============================================================================
# KONFIGURASI PDF UPLOAD (LOKAL) 
# ============================================================================

app.config['PDF_UPLOAD_FOLDER'] = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'static', 'pdf_uploads')
os.makedirs(app.config['PDF_UPLOAD_FOLDER'], exist_ok=True)


def upload_pdf_to_local(file, operator, nomor_ba):
    """Upload file PDF ke folder lokal - nomor_ba asli tetap di database, nama file dibersihkan"""
    if not file or not file.filename:
        return None, "File tidak ditemukan"
    
    if not file.filename.lower().endswith('.pdf'):
        return None, "Format file harus PDF"
    
    try:
         # Buat struktur folder: pdf_uploads/[operator]/[tahun]/[bulan]/
        pdf_folder = app.config['PDF_UPLOAD_FOLDER']
        operator_folder = os.path.join(pdf_folder, operator)
        year = datetime.now().strftime('%Y')
        month = datetime.now().strftime('%m')  # Tambah folder bulan untuk menghindari terlalu banyak file
        bulan_folder = os.path.join(operator_folder, year, month)
        
        os.makedirs(bulan_folder, exist_ok=True)
        
         # Buat nama file unik
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        safe_filename = secure_filename(file.filename)

         # Bersihkan nomor BA untuk nama file
        clean_nomor_ba = re.sub(r'[\\/*?:"<>|/]', '_', nomor_ba)
        clean_nomor_ba = clean_nomor_ba.replace('/', '_').replace('\\', '_')
        clean_nomor_ba = clean_nomor_ba.replace(' ', '_')
        # Batasi panjang (max 100 karakter)
        if len(clean_nomor_ba) > 100:
            clean_nomor_ba = clean_nomor_ba[:100]
        
        filename = f"{timestamp}_{clean_nomor_ba}_{safe_filename}"
        filepath = os.path.join(bulan_folder, filename)
        
        file.save(filepath)
        
        # Buat URL untuk mengakses file
        pdf_url = url_for('serve_pdf_file', operator=operator, year=year, month=month, filename=filename, _external=True)
        
        print(f"✅ PDF saved locally: {filepath}")
        print(f"   Original BA (in DB): {nomor_ba}")
        print(f"   Clean BA (in filename): {clean_nomor_ba}")
        print(f"   URL: {pdf_url}")
        
        return {
            'file_path': filepath,
            'url': pdf_url,
            'original_filename': safe_filename
        }, None
        
    except Exception as e:
        print(f"❌ Error saving PDF: {e}")
        import traceback
        traceback.print_exc()
        return None, str(e)

def delete_pdf_from_local(file_path):
    """Hapus file PDF dari lokal"""
    try:
        if file_path and os.path.exists(file_path):
            os.remove(file_path)
            print(f"✅ Deleted local PDF: {file_path}")
            return True, None
        return False, "File tidak ditemukan"
    except Exception as e:
        return False, str(e)


@app.route('/pdf_uploads/<operator>/<year>/<month>/<filename>')
def serve_pdf_file(operator, year, month, filename):
    """Serve PDF file from local storage"""
    filepath = os.path.join(app.config['PDF_UPLOAD_FOLDER'], operator, year, month, filename)
    if os.path.exists(filepath):
        return send_file(filepath, mimetype='application/pdf')
    return "File not found", 404

# ============================================================================
# HEADER CACHE CONTROL UNTUK MENCEGAH CACHE SETELAH LOGOUT
# ============================================================================

@app.after_request
def add_header(response):
    """
    Add headers to both force latest IE rendering engine or Chrome Frame,
    and also to cache the rendered page for 10 minutes.
    """
    response.headers['Cache-Control'] = 'no-store, no-cache, must-revalidate, post-check=0, pre-check=0, max-age=0'
    response.headers['Pragma'] = 'no-cache'
    response.headers['Expires'] = '-1'
    return response

# Daftar operator dan kota
OPERATORS = ['telkom', 'telkomsel', 'ioh', 'xl smart']
KOTA_LIST = [
    'Kota Samarinda', 'Kota Balikpapan', 'Kota Bontang', 'Kutai Barat',
    'Kutai Kartanegara', 'Kutai Timur', 'Berau', 'Paser', 'Penajam Paser Utara', 'Mahakam Ulu'
]
STATUS_OPTIONS = {
    'aktif': 'Aktif',
    'tidak_aktif': 'Tidak Aktif',
    'tidak_sesuai': 'Tidak Sesuai Parameter'
}

# ============================================================================
# JSON ENCODER
# ============================================================================
class CustomJSONEncoder(json.JSONEncoder):
    def default(self, obj):
        # Handle UploadGambar objects
        if isinstance(obj, UploadGambar):
            thumbnail_url = None
            if obj.public_id:
                thumbnail_url, _ = cloudinary_url(
                    obj.public_id, 
                    width=200, 
                    height=200, 
                    crop="fill", 
                    quality="auto"
                )
            
            return {
                'id': obj.id,
                'stasiun_id': obj.stasiun_id,
                'stasiun_lawan_id': obj.stasiun_lawan_id,
                'group_id': obj.group_id,
                'public_id': obj.public_id,
                'url': obj.cloudinary_url,
                'thumbnail_url': thumbnail_url,
                'original_filename': obj.original_filename,
                'status': obj.status,
                'uploaded_by': obj.uploaded_by,
                'uploaded_at': obj.uploaded_at.isoformat() if obj.uploaded_at else None,
                'is_checked': obj.is_checked,
                'width': obj.width,
                'height': obj.height,
                'format': obj.format,
                'bytes_size': obj.bytes_size
            }
        # Handle StasiunLawan objects
        elif isinstance(obj, StasiunLawan):
            return {
                'id': obj.id,
                'nama_stasiun_lawan': obj.nama_stasiun_lawan,
                'stasiun_id': obj.stasiun_id,
                'freq_tx': obj.freq_tx,
                'freq_rx': obj.freq_rx,
                'group_id': obj.group_id,
                'urutan': obj.urutan,
                'created_at': obj.created_at.isoformat() if obj.created_at else None
            }
        # Handle datetime
        elif isinstance(obj, datetime):
            return obj.isoformat()
        # Handle SQLAlchemy models
        elif hasattr(obj, '__table__'):
            # For SQLAlchemy models
            data = {}
            for column in obj.__table__.columns:
                value = getattr(obj, column.name)
                if isinstance(value, datetime):
                    data[column.name] = value.isoformat()
                else:
                    data[column.name] = value
            return data
        # Handle other objects with __dict__
        elif hasattr(obj, '__dict__'):
            try:
                # Remove internal attributes
                data = {}
                for key, value in obj.__dict__.items():
                    if not key.startswith('_'):
                        if isinstance(value, datetime):
                            data[key] = value.isoformat()
                        else:
                            data[key] = value
                return data
            except:
                pass
        return super().default(obj)

app.json_encoder = CustomJSONEncoder

# ============================================================================
# DATABASE MODELS
# ============================================================================

db = SQLAlchemy(app)
login_manager = LoginManager()
login_manager.init_app(app)
login_manager.login_view = 'login'
login_manager.login_message = 'Silakan login terlebih dahulu untuk mengakses halaman ini.'
login_manager.login_message_category = 'warning'
login_manager.refresh_view = 'login'
login_manager.needs_refresh_message = 'Sesi Anda telah berakhir, silakan login kembali.'
login_manager.needs_refresh_message_category = 'warning'

@login_manager.unauthorized_handler
def unauthorized_callback():
    if request.path.startswith('/api/') or request.is_json or request.headers.get('Accept') == 'application/json':
        return jsonify({
            'status': 'error',
            'message': 'Sesi login Anda telah berakhir. Silakan login kembali.'
        }), 401
    return redirect(url_for('login', next=request.url))

class User(UserMixin, db.Model):
    __tablename__ = 'users'
    id = db.Column(db.Integer, primary_key=True, autoincrement=True)
    username = db.Column(db.String(80), unique=True, nullable=False)
    password = db.Column(db.String(200), nullable=False)
    role = db.Column(db.String(50), nullable=False)  # admin_master, admin_operator, user_operator
    operator_type = db.Column(db.String(50), nullable=True)  # telkom, telkomsel, indosat, xl, NULL untuk admin_master
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    updated_at = db.Column(db.DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)
    
    def __repr__(self):
        return f'<User {self.username} - {self.role} - {self.operator_type}>'

class GrupStasiun(db.Model):
    __tablename__ = 'grup_stasiun'
    id = db.Column(db.Integer, primary_key=True, autoincrement=True)
    stasiun_id = db.Column(db.Integer, db.ForeignKey('stasiun.id'), nullable=False)
    nama_grup = db.Column(db.String(200), nullable=False)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    updated_at = db.Column(db.DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)
    
    # Relasi
    stasiun = db.relationship('Stasiun', backref='grup_list')
    lawan_list = db.relationship('StasiunLawan', backref='grup', lazy=True)

class Stasiun(db.Model):
    __tablename__ = 'stasiun'
    id = db.Column(db.Integer, primary_key=True, autoincrement=True)
    stasiun_name = db.Column(db.String(200), nullable=False)
    operator = db.Column(db.String(50), nullable=False)
    kota = db.Column(db.String(100), nullable=False)
    created_by = db.Column(db.Integer, db.ForeignKey('users.id'))
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    updated_at = db.Column(db.DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)
    
    # Kolom baru untuk status pemeriksaan stasiun
    status_pemeriksaan = db.Column(db.String(20), default='belum_diperiksa', nullable=False)
    checked_at = db.Column(db.DateTime, nullable=True)  # Waktu pemeriksaan
    checked_by = db.Column(db.Integer, db.ForeignKey('users.id'), nullable=True)  # Yang memeriksa
    
    stasiun_lawan_list = db.relationship('StasiunLawan', backref='stasiun', lazy=True, cascade='all, delete-orphan')
    uploads = db.relationship('UploadGambar', backref='stasiun', lazy=True, cascade='all, delete-orphan')
    
    def __repr__(self):
        return f'<Stasiun {self.stasiun_name} - {self.operator}>'

class StasiunLawan(db.Model):
    __tablename__ = 'stasiun_lawan'
    id = db.Column(db.Integer, primary_key=True, autoincrement=True)
    stasiun_id = db.Column(db.Integer, db.ForeignKey('stasiun.id'), nullable=False)
    nama_stasiun_lawan = db.Column(db.String(200), nullable=False)
    freq_tx = db.Column(db.String(50), nullable=True) 
    freq_rx = db.Column(db.String(50), nullable=True) 
    group_id = db.Column(db.Integer, nullable=True)
    grup_id = db.Column(db.Integer, db.ForeignKey('grup_stasiun.id'), nullable=True)
    urutan = db.Column(db.Integer, default=0)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    
    # Kolom parameter teknis dan status pemeriksaan stasiun lawan
    nomor_isr = db.Column(db.String(100), nullable=True)
    hasil_pemeriksaan = db.Column(db.String(100), default='Sesuai ISR')
    bw_mhz = db.Column(db.String(50), nullable=True)
    tipe_perangkat = db.Column(db.String(100), nullable=True)
    metode_pemeriksaan = db.Column(db.String(100), nullable=True)
    tanggal_pemeriksaan = db.Column(db.Date, nullable=True)
    catatan_pemeriksaan = db.Column(db.Text, nullable=True)
    link_id = db.Column(db.String(100), nullable=True)
    id_sims = db.Column(db.String(100), nullable=True)
    client_operator = db.Column(db.String(200), nullable=True)
    sertifikat_perangkat = db.Column(db.String(50), default='Ada')
    mulai_beroperasi = db.Column(db.String(100), nullable=True)
    longitude = db.Column(db.String(50), nullable=True)
    latitude = db.Column(db.String(50), nullable=True)
    hasil_freq_tx = db.Column(db.String(50), nullable=True)
    hasil_freq_rx = db.Column(db.String(50), nullable=True)
    hasil_bw_mhz = db.Column(db.String(50), nullable=True)
    hasil_tipe_perangkat = db.Column(db.String(100), nullable=True)
    hasil_longitude = db.Column(db.String(50), nullable=True)
    hasil_latitude = db.Column(db.String(50), nullable=True)
    updated_at = db.Column(db.DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)
    
    uploads = db.relationship('UploadGambar', backref='stasiun_lawan_obj', lazy=True)
    status_updates = db.relationship('StatusUpdate', backref='stasiun_lawan', lazy=True, cascade='all, delete-orphan')

# ============================================================================
# MODEL HASIL LAPORAN (Tambahkan di bagian models)
# ============================================================================
class HasilLaporan(db.Model):
    __tablename__ = 'hasil_laporan'
    id = db.Column(db.Integer, primary_key=True, autoincrement=True)
    operator = db.Column(db.String(50), nullable=False)
    nomor_ba = db.Column(db.String(100), nullable=False, unique=True)
    tanggal_ba = db.Column(db.Date, nullable=False)
    nomor_spt = db.Column(db.String(100), nullable=False)
    kota_kab = db.Column(db.String(100), nullable=False)
    file_pdf_url = db.Column(db.String(500), nullable=True)
    file_pdf_path = db.Column(db.String(500), nullable=True)  # ← PASTIKAN INI ADA
    created_by = db.Column(db.Integer, db.ForeignKey('users.id'))
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    updated_at = db.Column(db.DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)
    
    creator = db.relationship('User', backref='laporan_created')
    
    def __repr__(self):
        return f'<HasilLaporan {self.nomor_ba} - {self.operator}>'
    
    def to_dict(self):
        return {
            'id': self.id,
            'operator': self.operator,
            'nomor_ba': self.nomor_ba,
            'tanggal_ba': self.tanggal_ba.strftime('%Y-%m-%d') if self.tanggal_ba else None,
            'nomor_spt': self.nomor_spt,
            'kota_kab': self.kota_kab,
            'file_pdf_url': self.file_pdf_url,
            'created_by': self.created_by,
            'creator_name': self.creator.username if self.creator else None,
            'created_at': self.created_at.strftime('%d/%m/%Y %H:%M') if self.created_at else None
        }
    
class UploadGambar(db.Model):
    __tablename__ = 'upload_gambar'
    id = db.Column(db.Integer, primary_key=True, autoincrement=True)
    stasiun_id = db.Column(db.Integer, db.ForeignKey('stasiun.id'), nullable=False)
    stasiun_lawan_id = db.Column(db.Integer, db.ForeignKey('stasiun_lawan.id'), nullable=True)
    group_id = db.Column(db.Integer, nullable=True)
    
    # ===== CLOUDINARY FIELDS =====
    public_id = db.Column(db.String(300), nullable=False)  # Cloudinary public ID
    cloudinary_url = db.Column(db.String(500), nullable=False)  # URL dari Cloudinary
    original_filename = db.Column(db.String(300), nullable=True)  # Nama file asli
    
    # Metadata dari Cloudinary
    width = db.Column(db.Integer, nullable=True)
    height = db.Column(db.Integer, nullable=True)
    format = db.Column(db.String(10), nullable=True)
    bytes_size = db.Column(db.Integer, nullable=True)  
    # =================================
    
    status = db.Column(db.String(50), nullable=True)
    uploaded_by = db.Column(db.Integer, db.ForeignKey('users.id'))
    uploaded_at = db.Column(db.DateTime, default=datetime.utcnow)
    is_checked = db.Column(db.Boolean, default=False)
    
    def __repr__(self):
        return f'<UploadGambar {self.public_id}>'
    
    def to_dict(self):
        """Untuk serialisasi JSON"""
        # Dapatkan URL thumbnail dari Cloudinary
        thumbnail_url = None
        if self.public_id:
            thumbnail_url, _ = cloudinary_url(
                self.public_id, 
                width=200, 
                height=200, 
                crop="fill", 
                quality="auto"
            )
        
        return {
            'id': self.id,
            'stasiun_id': self.stasiun_id,
            'stasiun_lawan_id': self.stasiun_lawan_id,
            'group_id': self.group_id,
            'public_id': self.public_id,
            'url': self.cloudinary_url,
            'thumbnail_url': thumbnail_url,
            'original_filename': self.original_filename,
            'status': self.status,
            'uploaded_by': self.uploaded_by,
            'uploaded_at': self.uploaded_at.isoformat() if self.uploaded_at else None,
            'is_checked': self.is_checked,
            'width': self.width,
            'height': self.height,
            'format': self.format,
            'size_mb': round(self.bytes_size / (1024 * 1024), 2) if self.bytes_size else None
        }
    
class StatusUpdate(db.Model):
    __tablename__ = 'status_updates'
    id = db.Column(db.Integer, primary_key=True, autoincrement=True)
    stasiun_lawan_id = db.Column(db.Integer, db.ForeignKey('stasiun_lawan.id'), nullable=False)
    status = db.Column(db.String(50), nullable=False)
    updated_by = db.Column(db.Integer, db.ForeignKey('users.id'))
    updated_at = db.Column(db.DateTime, default=datetime.utcnow)
    catatan = db.Column(db.Text, nullable=True)
    
    def __repr__(self):
        return f'<StatusUpdate {self.status}>'

@login_manager.user_loader
def load_user(user_id):
    # FIX: Menggunakan SQLAlchemy 2.0 style untuk menghindari warning
    try:
        return db.session.get(User, int(user_id))
    except:
        return None

# ============================================================================
# CONTEXT PROCESSORS
# ============================================================================

@app.context_processor
def inject_models():
    """Inject models ke semua template"""
    return dict(
        StatusUpdate=StatusUpdate,
        db=db,
        Stasiun=Stasiun,
        StasiunLawan=StasiunLawan,
        UploadGambar=UploadGambar,
        User=User
    )

# ============================================================================
# CUSTOM DECORATORS
# ============================================================================

def admin_master_required(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if current_user.role != 'admin_master':
            flash('Akses ditolak! Hanya untuk Admin Master.', 'error')
            return redirect(url_for('index'))
        return f(*args, **kwargs)
    return decorated_function

def admin_operator_required(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if current_user.role != 'admin_operator':
            flash('Akses ditolak! Hanya untuk Admin Operator.', 'error')
            return redirect(url_for('index'))
        return f(*args, **kwargs)
    return decorated_function

def user_operator_required(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if current_user.role != 'user_operator':
            flash('Akses ditolak! Hanya untuk User Operator.', 'error')
            return redirect(url_for('index'))
        return f(*args, **kwargs)
    return decorated_function

def admin_master_or_admin_operator_required(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if current_user.role not in ['admin_master', 'admin_operator']:
            flash('Akses ditolak! Hanya untuk Admin.', 'error')
            return redirect(url_for('index'))
        return f(*args, **kwargs)
    return decorated_function

def get_operator_filter():
    """Get operator filter based on user role"""
    if current_user.role == 'admin_master':
        return OPERATORS
    elif current_user.role == 'admin_operator':
        return [current_user.operator_type]
    elif current_user.role == 'user_operator':
        return [current_user.operator_type]
    return []

def get_accessible_operators():
    """Get list of operators accessible by current user"""
    if current_user.role == 'admin_master':
        return OPERATORS
    elif current_user.role == 'admin_operator':
        return [current_user.operator_type]
    elif current_user.role == 'user_operator':
        return [current_user.operator_type]
    return []

# ============================================================================
# HELPER FUNCTIONS
# ============================================================================

def allowed_file(filename):
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in app.config['ALLOWED_EXTENSIONS']

def create_upload_folders():
    """Membuat folder upload jika belum ada"""
    upload_folder = app.config['UPLOAD_FOLDER']
    if not os.path.exists(upload_folder):
        os.makedirs(upload_folder)
        for operator in OPERATORS:
            operator_folder = os.path.join(upload_folder, operator)
            if not os.path.exists(operator_folder):
                os.makedirs(operator_folder)
                # Buat subfolder untuk stasiun-stasiun
                for i in range(1, 11):  # Contoh: buat 10 folder awal
                    station_folder = os.path.join(operator_folder, str(i))
                    os.makedirs(station_folder, exist_ok=True)

def sanitize_filename(filename):
    """Bersihkan nama file dari karakter tidak valid"""
    filename = secure_filename(filename)
    filename = filename.replace('@', '_').replace(' ', '_')
    return filename

def generate_unique_filename(original_filename, station_id, opponent_id=None, group_id=None):
    """
    Generate nama file yang UNIK untuk menghindari konflik antar user
    Format: YYYYMMDD_HHMMSS_FFF_RANDOM_NAMAAWAL.EXT
    Contoh: 20260226_103045_123_a1b2c3_screenshoot1.jpg
    """
    if not original_filename or '.' not in original_filename:
        return None
    
    # Bersihkan nama file asli
    safe_name = sanitize_filename(original_filename)
    
    # Ambil ekstensi file
    name_part, ext_part = safe_name.rsplit('.', 1)
    
    # Generate timestamp dengan microsecond
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S_%f')[:-3]
    
    # Generate random string (6 karakter) untuk memastikan keunikan
    random_str = uuid.uuid4().hex[:6]
    
    # Format nama file: TIMESTAMP_RANDOM_NAMAASLI.EXT
    # Contoh: 20260226_103045_123_a1b2c3_screenshoot1.jpg
    new_filename = f"{timestamp}_{random_str}_{safe_name}"
    
    # Potong jika terlalu panjang (maks 200 karakter)
    if len(new_filename) > 200:
        # Potong nama file asli, pertahankan timestamp dan random
        max_name_length = 200 - len(timestamp) - len(random_str) - 10
        safe_name = safe_name[:max_name_length]
        new_filename = f"{timestamp}_{random_str}_{safe_name}"
    
    return new_filename

def get_stations_with_detailed_groups_paginated(operator, kota='all', search_stasiun='', page=1, per_page=20):
    """OPTIMIZED VERSION - Hanya 3-4 query total!"""
    from sqlalchemy import func
    
    # ========== 1. QUERY STASIUN DENGAN PAGINATION ==========
    query = Stasiun.query.filter_by(operator=operator)
    
    if kota != 'all' and kota:
        query = query.filter(Stasiun.kota.ilike(kota))
    
    if search_stasiun:
        query = query.filter(Stasiun.stasiun_name.ilike(f'%{search_stasiun}%'))
    
    # Pagination di sini
    pagination = query.order_by(Stasiun.stasiun_name).paginate(page=page, per_page=per_page, error_out=False)
    stations = pagination.items
    
    if not stations:
        return stations, pagination
    
    # Ambil semua ID stasiun dalam 1 halaman
    station_ids = [s.id for s in stations]
    
    # ========== 2. AMBIL SEMUA LAWAN SEKALIGUS (1 QUERY) ==========
    all_opponents = StasiunLawan.query\
        .filter(StasiunLawan.stasiun_id.in_(station_ids))\
        .order_by(StasiunLawan.urutan)\
        .all()
    
    # ========== 3. AMBIL SEMUA STATUS TERBARU SEKALIGUS (1 QUERY) ==========
    opponent_ids = [o.id for o in all_opponents]
    latest_status_dict = {}
    
    if opponent_ids:
        # Subquery untuk status terbaru per lawan
        subquery = db.session.query(
            StatusUpdate.stasiun_lawan_id,
            func.max(StatusUpdate.updated_at).label('max_updated')
        ).filter(StatusUpdate.stasiun_lawan_id.in_(opponent_ids))\
         .group_by(StatusUpdate.stasiun_lawan_id)\
         .subquery()
        
        latest_statuses = db.session.query(StatusUpdate)\
            .join(subquery, 
                  (StatusUpdate.stasiun_lawan_id == subquery.c.stasiun_lawan_id) &
                  (StatusUpdate.updated_at == subquery.c.max_updated))\
            .all()
        
        for status in latest_statuses:
            latest_status_dict[status.stasiun_lawan_id] = status
    
    # ========== 4. HITUNG UPLOAD PER STASIUN (1 QUERY) ==========
    upload_counts = db.session.query(
        UploadGambar.stasiun_id,
        func.count(func.distinct(UploadGambar.public_id)).label('count')
    ).filter(UploadGambar.stasiun_id.in_(station_ids))\
     .group_by(UploadGambar.stasiun_id)\
     .all()
    
    upload_count_dict = {uc.stasiun_id: uc.count for uc in upload_counts}
    
    # ========== 5. AMBIL SEMUA GRUP PER STASIUN (1 QUERY) ==========
    all_groups = GrupStasiun.query.filter(GrupStasiun.stasiun_id.in_(station_ids)).all()
    groups_by_station = {}
    for group in all_groups:
        if group.stasiun_id not in groups_by_station:
            groups_by_station[group.stasiun_id] = []
        groups_by_station[group.stasiun_id].append(group)
    
    # ========== 6. GROUPING OPPONENTS PER STASIUN ==========
    opponents_by_station = {}
    for opponent in all_opponents:
        if opponent.stasiun_id not in opponents_by_station:
            opponents_by_station[opponent.stasiun_id] = []
        
        status = latest_status_dict.get(opponent.id)
        
        opponent_data = {
            'id': opponent.id,
            'nama': opponent.nama_stasiun_lawan,
            'freq_tx': opponent.freq_tx,
            'freq_rx': opponent.freq_rx,
            'group_id': opponent.group_id,
            'status': status.status if status else None,
            'status_display': STATUS_OPTIONS.get(status.status, 'Belum Ada') if status else 'Belum Ada',
            'catatan': status.catatan if status else ''
        }
        opponents_by_station[opponent.stasiun_id].append(opponent_data)
    
    # ========== 7. ATTACH DATA KE STASIUN ==========
    for station in stations:
        station.stasiun_lawan_list = opponents_by_station.get(station.id, [])
        station.upload_count = upload_count_dict.get(station.id, 0)
        station.grup_list = groups_by_station.get(station.id, [])
        
        # Hitung statistik status summary
        status_summary = {'aktif': 0, 'tidak_aktif': 0, 'tidak_sesuai': 0}
        has_lawan_tanpa_status = False
        
        for lawan in station.stasiun_lawan_list:
            if lawan['status']:
                if lawan['status'] in status_summary:
                    status_summary[lawan['status']] += 1
            else:
                has_lawan_tanpa_status = True
        
        station.status_summary = status_summary
        station.has_lawan_without_status = has_lawan_tanpa_status
        
        # ===== BUILD GROUPS DISPLAY (sama seperti kode asli Anda) =====
        groups_dict = {g.id: g.nama_grup for g in station.grup_list}
        
        lawan_by_group = {}
        group_stats = {}
        
        for lawan in station.stasiun_lawan_list:
            group_key = lawan['group_id'] if lawan['group_id'] is not None else 'ungrouped'
            if group_key not in lawan_by_group:
                lawan_by_group[group_key] = []
                group_stats[group_key] = {
                    'total': 0,
                    'aktif': 0,
                    'tidak_aktif': 0,
                    'tidak_sesuai': 0,
                    'belum_ada': 0
                }
            
            lawan_by_group[group_key].append(lawan)
            
            group_stats[group_key]['total'] += 1
            status = lawan['status']
            if status in group_stats[group_key]:
                group_stats[group_key][status] += 1
            else:
                group_stats[group_key]['belum_ada'] += 1
        
        groups_display = []
        
        if 'ungrouped' in lawan_by_group:
            groups_display.append({
                'group_id': None,
                'group_name': 'Tanpa Grup',
                'lawans': lawan_by_group['ungrouped'],
                'stats': group_stats['ungrouped'],
                'is_ungrouped': True,
                'badge_color': 'warning'
            })
        
        for group_key, lawans in lawan_by_group.items():
            if group_key != 'ungrouped':
                try:
                    group_id = int(group_key)
                    group_name = groups_dict.get(group_id, f'Grup {group_id}')
                    groups_display.append({
                        'group_id': group_id,
                        'group_name': group_name,
                        'lawans': lawans,
                        'stats': group_stats[group_key],
                        'is_ungrouped': False,
                        'badge_color': 'success'
                    })
                except:
                    continue
        
        groups_display.sort(key=lambda x: (x['is_ungrouped'], x['group_id'] if x['group_id'] is not None else 0))
        station.groups_display = groups_display
        
        station.total_opponents = len(station.stasiun_lawan_list)
        station.total_groups = len([g for g in groups_display if not g['is_ungrouped']])
    
    return stations, pagination

def get_grouped_opponents(station_id):
    """Get opponents grouped by group_id"""
    opponents = StasiunLawan.query.filter_by(stasiun_id=station_id).order_by(StasiunLawan.urutan).all()
    
    # Kelompokkan berdasarkan group_id
    groups = {}
    for opponent in opponents:
        group_key = opponent.group_id if opponent.group_id is not None else 'ungrouped'
        if group_key not in groups:
            groups[group_key] = []
        groups[group_key].append(opponent)
    
    return groups

def get_grouped_opponents_serializable(station_id):
    """Get opponents grouped by group_id in serializable format"""
    opponents = StasiunLawan.query.filter_by(stasiun_id=station_id).order_by(StasiunLawan.urutan).all()
    
    # Kelompokkan berdasarkan group_id
    groups = {}
    for opponent in opponents:
        group_key = opponent.group_id if opponent.group_id is not None else 'ungrouped'
        if group_key not in groups:
            groups[group_key] = []
        
        # Get latest status
        latest_status = StatusUpdate.query\
            .filter_by(stasiun_lawan_id=opponent.id)\
            .order_by(StatusUpdate.updated_at.desc())\
            .first()
        
        # Buat format serializable
        opponent_data = opponent.to_dict()
        
        if latest_status:
            opponent_data['latest_status'] = {
                'status': latest_status.status,
                'catatan': latest_status.catatan,
                'updated_at': latest_status.updated_at.isoformat() if latest_status.updated_at else None
            }
        else:
            opponent_data['latest_status'] = None
        
        groups[group_key].append(opponent_data)
    
    return groups

def get_upload_data_by_group(station_id):
    """Get upload data grouped by group_id - UPDATE untuk Cloudinary"""
    # Get all uploads for this station
    all_uploads = UploadGambar.query.filter_by(stasiun_id=station_id).all()
    
    # Get all opponents for this station
    opponents = StasiunLawan.query.filter_by(stasiun_id=station_id).all()
    
    # Group opponents by group_id
    opponents_by_group = {}
    for opponent in opponents:
        group_key = opponent.group_id if opponent.group_id is not None else 'ungrouped'
        if group_key not in opponents_by_group:
            opponents_by_group[group_key] = []
        opponents_by_group[group_key].append(opponent)
    
    # Group uploads by group_id
    uploads_by_group = {}
    for upload in all_uploads:
        if upload.group_id is not None:
            group_key = upload.group_id
        elif upload.stasiun_lawan_id:
            # Get opponent's group_id
            opponent = StasiunLawan.query.get(upload.stasiun_lawan_id)
            if opponent:
                group_key = opponent.group_id if opponent.group_id is not None else 'ungrouped'
            else:
                group_key = 'ungrouped'
        else:
            group_key = 'ungrouped'
        
        if group_key not in uploads_by_group:
            uploads_by_group[group_key] = []
        
        # Don't add duplicate public_ids
        existing_public_ids = [u.public_id for u in uploads_by_group[group_key]]
        if upload.public_id not in existing_public_ids:
            uploads_by_group[group_key].append(upload)
    
    # Combine both
    grouped_data = {}
    all_group_keys = set(list(opponents_by_group.keys()) + list(uploads_by_group.keys()))
    
    for group_key in all_group_keys:
        grouped_data[group_key] = {
            'uploads': uploads_by_group.get(group_key, []),
            'opponents': opponents_by_group.get(group_key, [])
        }
    
    return grouped_data

def get_card_view_data(station):
    """Helper function untuk mendapatkan data card view - UPDATE untuk Cloudinary"""
    uploads = UploadGambar.query.filter_by(stasiun_id=station.id).all()
    
    # Get images with opponent info
    images = []
    for upload in uploads:
        opponent_name = None
        if upload.stasiun_lawan_id:
            lawan = StasiunLawan.query.get(upload.stasiun_lawan_id)
            opponent_name = lawan.nama_stasiun_lawan if lawan else None
        
        thumbnail_url = None
        if upload.public_id:
            thumbnail_url, _ = cloudinary_url(
                upload.public_id, 
                width=200, 
                height=200, 
                crop="fill", 
                quality="auto"
            )
        
        images.append({
            'id': upload.id,
            'url': upload.cloudinary_url,
            'thumbnail_url': thumbnail_url,
            'filename': upload.original_filename,
            'status': upload.status,
            'opponent': opponent_name,
            'date': upload.uploaded_at.strftime('%d/%m/%Y'),
            'format': upload.format,
            'width': upload.width,
            'height': upload.height,
            'size_mb': round(upload.bytes_size / (1024 * 1024), 2) if upload.bytes_size else None
        })
    
    # Get opponents list
    opponents = []
    for lawan in station.stasiun_lawan_list:
        opponents.append(lawan.nama_stasiun_lawan)
    
    return {
        'id': station.id,
        'name': station.stasiun_name,
        'kota': station.kota,
        'image_count': len(uploads),
        'opponent_count': len(station.stasiun_lawan_list),
        'opponents': opponents,
        'images': images,
        'created_at': station.created_at,
        'updated_at': station.updated_at
    }

def get_station_detail_data(station_id):
    """Fungsi untuk mendapatkan data detail stasiun - UPDATE untuk Cloudinary"""
    station = Stasiun.query.get_or_404(station_id)
    
    all_uploads_raw = UploadGambar.query.filter_by(stasiun_id=station_id).all()
    
    # Kelompokkan upload berdasarkan group_id dan public_id
    group_file_mapping = {}
    
    for upload in all_uploads_raw:
        group_key = upload.group_id if upload.group_id is not None else 'ungrouped'
        
        if group_key not in group_file_mapping:
            group_file_mapping[group_key] = []
        
        # Cek apakah public_id sudah ada dalam grup ini
        existing_public_ids = [u['public_id'] for u in group_file_mapping[group_key]]
        if upload.public_id not in existing_public_ids:
            group_file_mapping[group_key].append({
                'id': upload.id,
                'public_id': upload.public_id,
                'url': upload.cloudinary_url,
                'original_filename': upload.original_filename,
                'status': upload.status,
                'uploaded_at': upload.uploaded_at,
                'stasiun_lawan_id': upload.stasiun_lawan_id,
                'format': upload.format,
                'width': upload.width,
                'height': upload.height,
                'bytes_size': upload.bytes_size
            })
    
    # Buat struktur grouped_uploads
    grouped_uploads = {}
    
    for group_key, uploads_list in group_file_mapping.items():
        # Sort uploads terbaru dulu
        uploads_sorted = sorted(uploads_list, key=lambda x: x['uploaded_at'], reverse=True)
        
        # Dapatkan semua lawan dalam grup ini
        if group_key == 'ungrouped':
            # Untuk ungrouped, ambil lawan dari upload individual
            opponent_ids = set(u['stasiun_lawan_id'] for u in uploads_sorted if u['stasiun_lawan_id'])
            lawans_in_group = StasiunLawan.query.filter(StasiunLawan.id.in_(opponent_ids)).all() if opponent_ids else []
        else:
            # Untuk grup, ambil semua lawan dengan group_id yang sama
            lawans_in_group = StasiunLawan.query.filter_by(
                stasiun_id=station_id, 
                group_id=group_key
            ).all()
        
        # Hitung statistik status
        status_summary = {'aktif': 0, 'tidak_aktif': 0, 'tidak_sesuai': 0}
        opponent_list = []
        
        for lawan in lawans_in_group:
            latest_status = StatusUpdate.query\
                .filter_by(stasiun_lawan_id=lawan.id)\
                .filter(StatusUpdate.status != 'riwayat')\
                .order_by(StatusUpdate.updated_at.desc())\
                .first()
            
            opponent_info = {
                'id': lawan.id,
                'nama': lawan.nama_stasiun_lawan,
                'freq_tx': lawan.freq_tx,
                'freq_rx': lawan.freq_rx,
                'latest_status': latest_status.status if latest_status else None,
                'catatan': latest_status.catatan if latest_status else None,
                'status_updated_at': latest_status.updated_at if latest_status else None
            }
            
            if opponent_info['latest_status'] in status_summary:
                status_summary[opponent_info['latest_status']] += 1
            
            opponent_list.append(opponent_info)
        
        # Urutkan lawan berdasarkan nama
        opponent_list.sort(key=lambda x: x['nama'])
        
        grouped_uploads[group_key] = {
            'group_id': group_key if group_key != 'ungrouped' else None,
            'uploads': uploads_sorted,
            'opponents': opponent_list,
            'opponent_count': len(opponent_list),
            'status_summary': status_summary
        }
    
    # Urutkan grup: ungrouped dulu, lalu grup numeric
    sorted_groups = {}
    
    if 'ungrouped' in grouped_uploads:
        sorted_groups['ungrouped'] = grouped_uploads.pop('ungrouped')
    
    # Urutkan grup numeric
    numeric_groups = {}
    for key, data in grouped_uploads.items():
        try:
            group_num = int(key)
            numeric_groups[group_num] = data
        except ValueError:
            continue
    
    for group_num in sorted(numeric_groups.keys()):
        sorted_groups[str(group_num)] = numeric_groups[group_num]
    
    # Hitung statistik total
    stasiun_lawan = StasiunLawan.query.filter_by(stasiun_id=station_id).all()
    status_counts = {'aktif': 0, 'tidak_aktif': 0, 'tidak_sesuai': 0}
    
    for lawan in stasiun_lawan:
        latest_status = StatusUpdate.query\
            .filter_by(stasiun_lawan_id=lawan.id)\
            .filter(StatusUpdate.status != 'riwayat')\
            .order_by(StatusUpdate.updated_at.desc())\
            .first()
        
        if latest_status and latest_status.status in status_counts:
            status_counts[latest_status.status] += 1
    
    return {
        'station': station,
        'all_uploads': all_uploads_raw,
        'grouped_uploads': sorted_groups,
        'status_counts': status_counts,
        'total_opponents': len(stasiun_lawan)
    }

def get_group_upload_status(station_id):
    """Cek status upload per grup untuk suatu stasiun"""
    all_uploads = UploadGambar.query.filter_by(stasiun_id=station_id).all()
    
    # Kelompokkan upload berdasarkan group_id dan public_id unik
    group_upload_count = {}
    
    for upload in all_uploads:
        group_key = upload.group_id if upload.group_id is not None else 'ungrouped'
        
        # Gunakan set untuk public_id unik
        if group_key not in group_upload_count:
            group_upload_count[group_key] = {
                'public_ids': set(),
                'upload_count': 0
            }
        
        if upload.public_id not in group_upload_count[group_key]['public_ids']:
            group_upload_count[group_key]['public_ids'].add(upload.public_id)
            group_upload_count[group_key]['upload_count'] += 1
    
    # Format hasil
    result = {}
    for group_key, data in group_upload_count.items():
        result[group_key] = {
            'has_upload': data['upload_count'] > 0,
            'upload_count': data['upload_count']
        }
    
    return result

def validate_excel_file(file_obj):
    """Validasi file Excel"""
    if not file_obj or not file_obj.filename:
        return False, "File tidak ditemukan"
    
    # Cek ekstensi
    if not file_obj.filename.lower().endswith(('.xlsx', '.xls')):
        return False, "Format file harus .xlsx atau .xls"
    
    # Cek ukuran (maks 10MB)
    file_obj.seek(0, os.SEEK_END)
    file_size = file_obj.tell()
    file_obj.seek(0)
    
    if file_size > 10 * 1024 * 1024:  # 10MB
        return False, "Ukuran file terlalu besar (maks 10MB)"
    
    return True, "File valid"

def clean_dataframe(df):
    """Bersihkan dataframe dari whitespace dan NaN, serta normalisasi string"""
    # Replace NaN dengan string kosong
    df = df.fillna('')
    
    # Fungsi untuk membersihkan string
    def clean_value(x):
        if isinstance(x, str):
            # Hapus spasi berlebih di awal dan akhir
            x = x.strip()
            # Hapus spasi berlebih di dalam (multiple spaces menjadi satu)
            x = ' '.join(x.split())
            return x
        return x
    
    # Gunakan applymap untuk membersihkan semua sel
    # CATATAN: applymap masih ada di pandas versi lama, jika error gunakan alternatif di bawah
    try:
        # Coba gunakan applymap (pandas versi lama)
        df = df.applymap(clean_value)
    except AttributeError:
        # Fallback untuk pandas versi baru (tanpa applymap)
        for col in df.columns:
            if df[col].dtype == 'object':
                df[col] = df[col].apply(clean_value)
    
    return df

# ============================================================================
# FUNGSI BARU: UPLOAD KE CLOUDINARY
# ============================================================================

def upload_to_cloudinary(file, station_id, opponent_id=None, group_id=None):
    """
    Upload file ke Cloudinary dengan organisasi folder yang rapi
    Format folder: stasiun/[operator]/[tahun]/[bulan]/[station_id]/
    """
    if not file or not file.filename:
        return None, "File tidak ditemukan"
    
    if not allowed_file(file.filename):
        return None, "Format file tidak didukung"
    
    try:
        # Dapatkan objek stasiun untuk mendapatkan operator
        station = Stasiun.query.get(station_id)
        if not station:
            return None, "Stasiun tidak ditemukan"
        
        # Bersihkan nama file
        safe_filename = sanitize_filename(file.filename)
        name_without_ext, ext = os.path.splitext(safe_filename)
        
        # Generate timestamp
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        
        # Buat public_id dengan struktur folder yang rapi
        # Format: stasiun/[operator]/[tahun]/[bulan]/[station_id]/[timestamp]_[nama_file]
        year = datetime.now().strftime('%Y')
        month = datetime.now().strftime('%m')
        
        folder_path = f"stasiun/{station.operator}/{year}/{month}/{station_id}"
        public_id = f"{folder_path}/{timestamp}_{name_without_ext}"
        
        print(f"📤 Uploading to Cloudinary: {public_id}")
        
        # Upload ke Cloudinary dengan optimasi otomatis
        upload_result = cloudinary.uploader.upload(
            file,
            public_id=public_id,
            resource_type="image",
            
            # Optimasi otomatis
            eager=[
                {"width": 800, "height": 600, "crop": "limit", "quality": "auto:good"},
                {"width": 400, "height": 300, "crop": "fill", "quality": "auto:eco"}
            ],
            eager_async=True,
            
            # Tags untuk memudahkan pencarian
            tags=[f"station_{station_id}", f"operator_{station.operator}"],
            
            # Context menyimpan metadata
            context={
                "station_id": str(station_id),
                "station_name": station.stasiun_name,
                "operator": station.operator,
                "opponent_id": str(opponent_id) if opponent_id else "",
                "group_id": str(group_id) if group_id else "",
                "original_filename": safe_filename
            },
            
            # Response termasuk metadata
            return_response=True
        )
        
        print(f"✅ Upload successful!")
        print(f"   Public ID: {upload_result['public_id']}")
        print(f"   URL: {upload_result['secure_url']}")
        print(f"   Format: {upload_result['format']}")
        print(f"   Size: {upload_result['bytes']} bytes")
        
        return {
            'public_id': upload_result['public_id'],
            'url': upload_result['secure_url'],
            'format': upload_result['format'],
            'width': upload_result.get('width'),
            'height': upload_result.get('height'),
            'bytes': upload_result.get('bytes'),
            'original_filename': safe_filename
        }, None
        
    except Exception as e:
        print(f"❌ Error uploading to Cloudinary: {e}")
        import traceback
        traceback.print_exc()
        return None, str(e)


# ============================================================================
# FUNGSI BARU: HAPUS DARI CLOUDINARY
# ============================================================================

def delete_from_cloudinary(public_id):
    """
    Hapus file dari Cloudinary
    """
    try:
        result = cloudinary.uploader.destroy(public_id)
        
        if result.get('result') == 'ok':
            print(f"✅ Deleted from Cloudinary: {public_id}")
            return True, None
        else:
            print(f"⚠️ Failed to delete: {result}")
            return False, result.get('result')
            
    except Exception as e:
        print(f"❌ Error deleting from Cloudinary: {e}")
        return False, str(e)


# ============================================================================
# FUNGSI BARU: GET OPTIMIZED URL
# ============================================================================

def get_optimized_image_url(public_id, width=None, height=None, crop="fill", quality="auto"):
    """
    Dapatkan URL gambar yang sudah dioptimasi dari Cloudinary
    """
    options = {
        "quality": quality,
        "fetch_format": "auto",
        "secure": True
    }
    
    if width and height:
        options["width"] = width
        options["height"] = height
        options["crop"] = crop
    elif width:
        options["width"] = width
        options["crop"] = "scale"
    
    url, _ = cloudinary_url(public_id, **options)
    return url


# ============================================================================
# FUNGSI INI TIDAK DIGUNAKAN LAGI (COMMENT)
# ============================================================================

# def save_uploaded_file(file, station_id, opponent_id=None, group_id=None):
#     """TIDAK DIGUNAKAN - DIGANTI DENGAN upload_to_cloudinary"""
#     pass

# def delete_physical_file(filepath):
#     """TIDAK DIGUNAKAN - DIGANTI DENGAN delete_from_cloudinary"""
#     pass

# ============================================================================
# GENERATE TEMPLATE EXCEL UNTUK ADMIN MASTER
# ============================================================================

def generate_admin_master_template():
    """Generate template Excel untuk admin master dengan 6 kolom (FREQ TX dan FREQ RX opsional)"""
    wb = Workbook()
    ws = wb.active
    ws.title = "Template"
    
    # Hapus sheet default yang tidak digunakan
    if 'Sheet' in wb.sheetnames:
        std = wb['Sheet']
        wb.remove(std)
    
    # Header tabel - 6 KOLOM (tanpa tanda kurung di header)
    headers = ['OPERATOR', 'STASIUN NAME', 'STASIUN LAWAN', 'FREQ TX', 'FREQ RX', 'KOTA/KAB']
    ws.append(headers)
    
    # Style header
    for cell in ws[1]:
        cell.font = Font(bold=True, size=11)
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.fill = PatternFill(start_color="C6E0B4", end_color="C6E0B4", fill_type="solid")
    
    # =================== DATA CONTOH DENGAN 6 KOLOM ===================
    contoh_data = [
        # OPERATOR, STASIUN NAME, STASIUN LAWAN, FREQ TX, FREQ RX, KOTA/KAB
        ['TELKOM', 'TELKOM_STN_01', 'TELKOM_LAWAN_01', '1800 MHz', '1800 MHz', 'Kota Samarinda'],
        ['', '', 'TELKOM_LAWAN_02', '', '', ''],
        ['', '', 'TELKOM_LAWAN_03', '900 MHz', '900 MHz', ''],
        
        ['TELKOMSEL', 'TELKOMSEL_STN_01', 'TELKOMSEL_LAWAN_01', '', '', 'Kota Balikpapan'],
        ['', '', 'TELKOMSEL_LAWAN_02', '2100 MHz', '2100 MHz', ''],
        
        ['IOH', 'IOH_STN_01', 'IOH_LAWAN_01', '', '', 'Kota Bontang'],
        ['', '', 'IOH_LAWAN_02', '', '', ''],
        
        ['XL SMART', 'XL_SMART_STN_01', 'XL_SMART_LAWAN_01', '2100 MHz', '2100 MHz', 'Kutai Kartanegara'],
        ['', '', 'XL_SMART_LAWAN_02', '', '', ''],
        ['', '', 'XL_SMART_LAWAN_03', '900 MHz', '900 MHz', ''],
    ]
    
    # Tambahkan data
    for row in contoh_data:
        ws.append(row)
    
    # =================== MERGE CELL ===================
    # Merge untuk TELKOM_STN_01 (baris 2-4, 3 lawan)
    if ws.max_row >= 4:
        ws.merge_cells(start_row=2, start_column=1, end_row=4, end_column=1)  # OPERATOR
        ws.merge_cells(start_row=2, start_column=2, end_row=4, end_column=2)  # STASIUN NAME
        ws.merge_cells(start_row=2, start_column=6, end_row=4, end_column=6)  # KOTA/KAB
    
    # Merge untuk TELKOMSEL_STN_01 (baris 5-6, 2 lawan)
    if ws.max_row >= 6:
        ws.merge_cells(start_row=5, start_column=1, end_row=6, end_column=1)  # OPERATOR
        ws.merge_cells(start_row=5, start_column=2, end_row=6, end_column=2)  # STASIUN NAME
        ws.merge_cells(start_row=5, start_column=6, end_row=6, end_column=6)  # KOTA/KAB
    
    # Merge untuk INDOSAT_STN_01 (baris 7-8, 2 lawan)
    if ws.max_row >= 8:
        ws.merge_cells(start_row=7, start_column=1, end_row=8, end_column=1)  # OPERATOR
        ws.merge_cells(start_row=7, start_column=2, end_row=8, end_column=2)  # STASIUN NAME
        ws.merge_cells(start_row=7, start_column=6, end_row=8, end_column=6)  # KOTA/KAB
    
    # Merge untuk XL_STN_01 (baris 9-11, 3 lawan)
    if ws.max_row >= 11:
        ws.merge_cells(start_row=9, start_column=1, end_row=11, end_column=1)  # OPERATOR
        ws.merge_cells(start_row=9, start_column=2, end_row=11, end_column=2)  # STASIUN NAME
        ws.merge_cells(start_row=9, start_column=6, end_row=11, end_column=6)  # KOTA/KAB
    
    # Center alignment untuk merge cells
    merge_cells = ['A2', 'B2', 'F2', 'A5', 'B5', 'F5', 'A7', 'B7', 'F7', 'A9', 'B9', 'F9']
    for cell_ref in merge_cells:
        if cell_ref in ws:
            ws[cell_ref].alignment = Alignment(vertical='center', horizontal='center')
    
    # Format lebar kolom
    ws.column_dimensions['A'].width = 15  # OPERATOR
    ws.column_dimensions['B'].width = 25  # STASIUN NAME
    ws.column_dimensions['C'].width = 25  # STASIUN LAWAN
    ws.column_dimensions['D'].width = 15  # FREQ TX
    ws.column_dimensions['E'].width = 15  # FREQ RX
    ws.column_dimensions['F'].width = 20  # KOTA/KAB
    
    # =================== SHEET INSTRUKSI ===================
    ws_inst = wb.create_sheet(title="Instruksi")
    
    instruksi = [
        ["INSTRUKSI PENGISIAN - ADMIN MASTER"],
        ["------------------------------------"],
        ["PERHATIAN: File ini untuk Admin Master (SEMUA OPERATOR)"],
        [""],
        ["FORMAT KOLOM (6 KOLOM):"],
        ["1. OPERATOR      : Nama operator (telkom/telkomsel/indosat/xl) - WAJIB"],
        ["2. STASIUN NAME  : Nama stasiun utama - WAJIB"],
        ["3. STASIUN LAWAN : Nama stasiun lawan - WAJIB"],
        ["4. FREQ TX       : Frekuensi Tx - OPSIONAL (boleh dikosongkan)"],
        ["5. FREQ RX       : Frekuensi Rx - OPSIONAL (boleh dikosongkan)"],
        ["6. KOTA/KAB      : Nama kota/kabupaten - WAJIB"],
        [""],
        ["CATATAN PENTING:"],
        ["- Kolom FREQ TX dan FREQ RX bersifat OPSIONAL - boleh dikosongkan"],
        ["- Jika tidak diisi, akan dibiarkan kosong di database"],
        ["- Bisa diisi nanti melalui halaman edit"],
        [""],
        ["CARA PENGISIAN:"],
        ["1. OPERATOR:"],
        ["   - Tulis sekali untuk setiap stasiun dengan banyak lawan"],
        ["   - Gunakan merge cell (contoh: A2:A4 untuk stasiun dengan 3 lawan)"],
        ["   - Pilihan: telkom, telkomsel, indosat, xl"],
        [""],
        ["2. STASIUN NAME:"],
        ["   - Tulis sekali untuk setiap stasiun dengan banyak lawan"],
        ["   - Gunakan merge cell seperti OPERATOR"],
        [""],
        ["3. STASIUN LAWAN:"],
        ["   - Satu baris untuk satu lawan"],
        ["   - Tidak perlu merge cell"],
        [""],
        ["4. FREQ TX  - OPSIONAL:"],
        ["   - Bisa dikosongkan jika tidak ada data"],
        ["   - Contoh format: 1800 MHz, 2100 MHz, 900 MHz, 850 MHz"],
        [""],
        ["5. FREQ RX - OPSIONAL:"],
        ["   - Bisa dikosongkan jika tidak ada data"],
        ["   - Contoh format: 1800 MHz, 2100 MHz, 900 MHz, 850 MHz"],
        [""],
        ["6. KOTA/KAB:"],
        ["   - Tulis sekali untuk setiap stasiun"],
        ["   - Gunakan merge cell seperti OPERATOR dan STASIUN NAME"],
    ]
    
    for row in instruksi:
        ws_inst.append(row)
    
    # Format untuk sheet instruksi
    ws_inst.column_dimensions['A'].width = 80
    ws_inst['A1'].font = Font(bold=True, size=14)
    ws_inst['A2'].font = Font(bold=True)
    
    # =================== SHEET DAFTAR OPERATOR ===================
    ws_op = wb.create_sheet(title="Daftar Operator")
    
    ws_op.append(["DAFTAR OPERATOR YANG TERSEDIA"])
    ws_op.append(["------------------------------"])
    ws_op.append(["Gunakan nama persis seperti di bawah:"])
    ws_op.append([""])
    
    # Daftar operator
    operator_list = ["telkom", "telkomsel", "ioh", "xl smart"]
    for op in operator_list:
        ws_op.append([op])
    
    ws_op.column_dimensions['A'].width = 20
    ws_op['A1'].font = Font(bold=True, size=12)
    
    # =================== SHEET DAFTAR KOTA ===================
    ws_kota = wb.create_sheet(title="Daftar Kota")
    
    ws_kota.append(["DAFTAR KOTA/KAB YANG TERSEDIA"])
    ws_kota.append(["-------------------------------"])
    ws_kota.append(["Gunakan nama persis seperti di bawah:"])
    ws_kota.append([""])
    
    # Daftar kota
    kota_list_display = [
        "Kota Samarinda", "Kota Balikpapan", "Kota Bontang", "Kutai Kartanegara", 
        "Kutai Barat", "Kutai Timur", "Penajam Paser Utara", "Paser",
        "Berau", "Mahakam Ulu"
    ]
    
    kota_list_sorted = sorted(kota_list_display, key=lambda x: x.lower())
    for kota in kota_list_sorted:
        ws_kota.append([kota])
    
    ws_kota.column_dimensions['A'].width = 30
    ws_kota['A1'].font = Font(bold=True, size=12)
    
    # Atur sheet order
    wb.active = wb['Template']
    
    # =================== SAVE FILE ===================
    output = BytesIO()
    
    # Simpan workbook
    wb.save(output)
    output.seek(0)
    
    return send_file(
        output,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True,
        download_name='Template_Admin_Master.xlsx'
    )

def process_excel_upload_admin_master(file_path, user, kota_default='samarinda'):
    """Process Excel upload untuk admin master dengan kolom FREQ TX dan FREQ RX opsional
    
    Fitur:
    - STASIUN NAME: Case-insensitive, jika sama dianggap satu stasiun
    - STASIUN LAWAN: Semua lawan ditambahkan tanpa pengecekan duplikat
    - KOTA/KAB: Setiap baris memiliki kota sendiri (tidak di-ffill)
    """
    try:
        # Baca Excel dengan pandas
        df = pd.read_excel(file_path, dtype=str)
        
        # NORMALISASI NAMA KOLOM (ubah semua ke uppercase)
        df.columns = [str(col).upper().strip() for col in df.columns]
        
        # Validasi kolom minimal dengan case-insensitive
        required_columns = ['OPERATOR', 'STASIUN NAME', 'STASIUN LAWAN']
        
        for col in required_columns:
            if col not in df.columns:
                # Coba cari dengan case-insensitive
                col_found = None
                for df_col in df.columns:
                    if df_col.upper() == col.upper():
                        col_found = df_col
                        break
                
                if col_found:
                    # Rename kolom ke format standar
                    df.rename(columns={col_found: col}, inplace=True)
                else:
                    raise ValueError(f"Kolom '{col}' tidak ditemukan dalam file Excel")
        
        # Deteksi kolom FREQ TX, FREQ RX, dan KOTA/KAB (opsional)
        has_freq_tx_column = False
        has_freq_rx_column = False
        has_kota_column = False
        
        # Normalisasi nama kolom untuk deteksi
        for col in df.columns:
            col_upper = col.upper()
            
            # Cek FREQ TX (case-insensitive)
            if any(x in col_upper for x in ['FREQ TX', 'FREQ_TX', 'FREQ-TX', 'TX']) and 'RX' not in col_upper:
                has_freq_tx_column = True
                if col != 'FREQ TX':
                    df.rename(columns={col: 'FREQ TX'}, inplace=True)
                print(f"DEBUG - Kolom FREQ TX ditemukan: {col}")
            
            # Cek FREQ RX (case-insensitive)
            if any(x in col_upper for x in ['FREQ RX', 'FREQ_RX', 'FREQ-RX', 'RX']):
                has_freq_rx_column = True
                if col != 'FREQ RX':
                    df.rename(columns={col: 'FREQ RX'}, inplace=True)
                print(f"DEBUG - Kolom FREQ RX ditemukan: {col}")
            
            # Cek KOTA (case-insensitive)
            if any(x in col_upper for x in ['KOTA', 'KAB', 'KABUPATEN']):
                has_kota_column = True
                if col != 'KOTA/KAB':
                    df.rename(columns={col: 'KOTA/KAB'}, inplace=True)
                print(f"DEBUG - Kolom KOTA ditemukan: {col}")
        
        # Forward fill untuk merge cell (HANYA UNTUK OPERATOR dan STASIUN NAME)
        df['OPERATOR'] = df['OPERATOR'].ffill()
        df['STASIUN NAME'] = df['STASIUN NAME'].ffill()
        
        # JANGAN forward fill untuk KOTA/KAB - biarkan per baris!
        # if has_kota_column and 'KOTA/KAB' in df.columns:
        #     df['KOTA/KAB'] = df['KOTA/KAB'].ffill()  # COMMENT atau HAPUS!
        
        # Bersihkan data (normalisasi string)
        df = clean_dataframe(df)
        
        # Normalisasi nilai OPERATOR (lowercase)
        df['OPERATOR'] = df['OPERATOR'].str.lower().str.strip()
        
        # Normalisasi STASIUN NAME untuk grouping (lowercase)
        df['STASIUN NAME_NORM'] = df['STASIUN NAME'].str.lower().str.strip()
        
        # Normalisasi KOTA (lowercase) - per baris, tidak di-ffill
        if has_kota_column and 'KOTA/KAB' in df.columns:
            df['KOTA/KAB'] = df['KOTA/KAB'].str.lower().str.strip()
        
        # Filter baris yang valid
        df = df[(df['STASIUN NAME'] != '') | (df['STASIUN LAWAN'] != '')]
        
        if df.empty:
            return 0, 0, ["File Excel tidak mengandung data yang valid"]
        
        success_count = 0
        error_count = 0
        messages = []
        
        # ===== GROUPING BERDASARKAN OPERATOR DAN STASIUN NAME (case-insensitive) =====
        grouped_by_station = {}
        
        for idx, row in df.iterrows():
            if not row['STASIUN NAME'] or not row['STASIUN LAWAN']:
                continue
            
            operator_norm = row['OPERATOR'].lower().strip()
            stasiun_name_norm = row['STASIUN NAME_NORM']
            stasiun_name_original = row['STASIUN NAME']
            
            key = (operator_norm, stasiun_name_norm)
            
            if key not in grouped_by_station:
                grouped_by_station[key] = {
                    'operator': operator_norm,
                    'original_operator': row['OPERATOR'],
                    'original_name': stasiun_name_original,
                    'lawan_list': []
                }
            
            # Tentukan kota untuk baris ini (per baris, tidak di-ffill!)
            if has_kota_column and 'KOTA/KAB' in df.columns and pd.notna(row['KOTA/KAB']) and str(row['KOTA/KAB']).strip():
                kota = str(row['KOTA/KAB']).strip().lower()
            else:
                kota = kota_default.lower()
            
            # Validasi kota (case-insensitive)
            kota_valid = False
            kota_normalized = None
            
            for k in KOTA_LIST:
                if k.lower() == kota.lower():
                    kota_normalized = k.lower()
                    kota_valid = True
                    break
            
            if not kota_valid:
                for k in KOTA_LIST:
                    if k.lower() in kota.lower() or kota.lower() in k.lower():
                        kota_normalized = k.lower()
                        kota_valid = True
                        break
            
            if not kota_valid:
                kota_normalized = kota_default.lower()
                messages.append(f"⚠️ Kota '{kota}' tidak valid, menggunakan default: {kota_default}")
            
            # Ambil freq_tx dan freq_rx (opsional - per baris)
            freq_tx = None
            freq_rx = None
            
            if has_freq_tx_column and 'FREQ TX' in df.columns and pd.notna(row['FREQ TX']) and str(row['FREQ TX']).strip():
                freq_tx = str(row['FREQ TX']).strip()
            
            if has_freq_rx_column and 'FREQ RX' in df.columns and pd.notna(row['FREQ RX']) and str(row['FREQ RX']).strip():
                freq_rx = str(row['FREQ RX']).strip()
            
            grouped_by_station[key]['lawan_list'].append({
                'nama': row['STASIUN LAWAN'],
                'kota': kota_normalized,
                'freq_tx': freq_tx,
                'freq_rx': freq_rx
            })
        
        print(f"DEBUG - Total grup stasiun unik: {len(grouped_by_station)}")
        
        # Process setiap stasiun
        for key, station_data in grouped_by_station.items():
            try:
                operator = station_data['operator']
                stasiun_name = station_data['original_name']
                lawan_list = station_data['lawan_list']
                
                # Validasi operator
                if operator not in OPERATORS:
                    operator_found = None
                    for op in OPERATORS:
                        if op.lower() == operator.lower():
                            operator_found = op
                            break
                    
                    if operator_found:
                        operator = operator_found
                        messages.append(f"ℹ️ Operator '{operator_found}' digunakan untuk stasiun '{stasiun_name}'")
                    else:
                        error_count += 1
                        messages.append(f"❌ Operator '{operator}' tidak valid untuk stasiun '{stasiun_name}'")
                        continue
                
                # Cek duplikasi stasiun (case-insensitive) - TIDAK MEMBUAT STASIUN BARU JIKA SUDAH ADA
                existing = Stasiun.query.filter(
                    db.func.lower(Stasiun.stasiun_name) == key[1],  # key[1] adalah stasiun_name_norm
                    Stasiun.operator == operator
                ).first()
                
                if existing:
                    # ===== STASIUN SUDAH ADA - TAMBAHKAN SEMUA STASIUN LAWAN =====
                    # Tidak ada pengecekan duplikat untuk stasiun lawan!
                    urutan = len(existing.stasiun_lawan_list)
                    added_count = 0
                    kota_used_count = {}
                    
                    for lawan_data in lawan_list:
                        # Selalu tambahkan stasiun lawan baru, tidak peduli duplikat
                        new_lawan = StasiunLawan(
                            stasiun_id=existing.id,
                            nama_stasiun_lawan=lawan_data['nama'],
                            freq_tx=lawan_data['freq_tx'],
                            freq_rx=lawan_data['freq_rx'],
                            group_id=None,
                            urutan=urutan
                        )
                        db.session.add(new_lawan)
                        urutan += 1
                        added_count += 1
                        
                        # Track kota yang digunakan
                        kota = lawan_data['kota']
                        kota_used_count[kota] = kota_used_count.get(kota, 0) + 1
                    
                    db.session.commit()
                    
                    messages.append(f"✅ Stasiun '{stasiun_name}': ditambah {added_count} stasiun lawan")
                    
                    # Tampilkan info kota
                    if kota_used_count:
                        kota_info = ', '.join([f"{k} ({c})" for k, c in kota_used_count.items()])
                        messages.append(f"   📍 Kota lawan: {kota_info}")
                    
                    success_count += 1
                    continue
                
                # ===== STASIUN BARU =====
                # Untuk stasiun baru, tentukan kota stasiun utama dari kota yang paling banyak digunakan
                kota_counts = {}
                for lawan_data in lawan_list:
                    kota = lawan_data['kota']
                    kota_counts[kota] = kota_counts.get(kota, 0) + 1
                
                if kota_counts:
                    # Ambil kota dengan jumlah terbanyak
                    kota_stasiun = max(kota_counts.items(), key=lambda x: x[1])[0]
                else:
                    kota_stasiun = kota_default.lower()
                
                # Buat stasiun baru
                new_station = Stasiun(
                    stasiun_name=stasiun_name,
                    operator=operator,
                    kota=kota_stasiun,
                    created_by=user.id
                )
                db.session.add(new_station)
                db.session.flush()
                
                # Tambah SEMUA stasiun lawan (termasuk duplikat)
                lawan_added = 0
                for i, lawan_data in enumerate(lawan_list):
                    if lawan_data['nama']:
                        new_lawan = StasiunLawan(
                            stasiun_id=new_station.id,
                            nama_stasiun_lawan=lawan_data['nama'],
                            freq_tx=lawan_data['freq_tx'],
                            freq_rx=lawan_data['freq_rx'],
                            group_id=None,
                            urutan=i
                        )
                        db.session.add(new_lawan)
                        lawan_added += 1
                
                success_count += 1
                db.session.commit()
                
                messages.append(f"✅ Stasiun baru '{stasiun_name}' ditambahkan dengan {lawan_added} lawan")
                
                # Tampilkan info kota
                if kota_counts:
                    kota_info = ', '.join([f"{k} ({c})" for k, c in kota_counts.items()])
                    messages.append(f"   📍 Kota stasiun utama: {kota_stasiun}")
                    messages.append(f"   📍 Distribusi kota lawan: {kota_info}")
                
            except Exception as e:
                db.session.rollback()
                error_count += 1
                messages.append(f"❌ Error pada stasiun '{stasiun_name}': {str(e)}")
                import traceback
                traceback.print_exc()
                continue
        
        summary = f"✅ Berhasil memproses {success_count} stasiun!"
        if error_count > 0:
            summary += f" ⚠️ {error_count} error ditemukan."
        
        messages.insert(0, summary)
        return success_count, error_count, messages
        
    except Exception as e:
        db.session.rollback()
        print(f"ERROR FATAL: {str(e)}")
        import traceback
        traceback.print_exc()
        return 0, 0, [f"❌ Error processing file: {str(e)}"]

def process_excel_upload_admin_operator(file_path, user, kota_default=None):
    """Process Excel upload untuk admin operator dengan FREQ TX dan FREQ RX opsional"""
    try:
        operator = user.operator_type
        
        # Baca Excel dengan pandas
        df = pd.read_excel(file_path, dtype=str)
        
        # NORMALISASI NAMA KOLOM (ubah semua ke uppercase)
        df.columns = [str(col).upper().strip() for col in df.columns]
        
        print(f"DEBUG - Nama kolom setelah normalisasi: {list(df.columns)}")
        
        # Mapping kolom dengan case-insensitive
        column_mapping = {}
        
        for col in df.columns:
            col_upper = col.upper()
            
            # Identifikasi STASIUN NAME
            if 'STASIUN NAME' in col_upper or ('NAME' in col_upper and 'STASIUN' in col_upper):
                column_mapping[col] = 'STASIUN NAME'
            elif 'STASIUN' in col_upper and 'LAWAN' in col_upper:
                column_mapping[col] = 'STASIUN LAWAN'
            # Identifikasi FREQ TX
            elif any(x in col_upper for x in ['FREQ TX', 'FREQ_TX', 'FREQ-TX', 'TX']) and 'RX' not in col_upper:
                column_mapping[col] = 'FREQ TX'
            # Identifikasi FREQ RX
            elif any(x in col_upper for x in ['FREQ RX', 'FREQ_RX', 'FREQ-RX', 'RX']):
                column_mapping[col] = 'FREQ RX'
            # Identifikasi KOTA/KAB
            elif any(x in col_upper for x in ['KOTA', 'KAB', 'KABUPATEN']):
                column_mapping[col] = 'KOTA/KAB'
        
        # Rename kolom sesuai mapping
        df = df.rename(columns=column_mapping)
        print(f"DEBUG - Columns after rename: {list(df.columns)}")
        
        # Validasi kolom minimal
        if 'STASIUN NAME' not in df.columns:
            return 0, 0, ["❌ Kolom STASIUN NAME tidak ditemukan dalam file Excel"]
        
        if 'STASIUN LAWAN' not in df.columns:
            return 0, 0, ["❌ Kolom STASIUN LAWAN tidak ditemukan dalam file Excel"]
        
        # Forward fill untuk merge cell (STASIUN NAME)
        df['STASIUN NAME'] = df['STASIUN NAME'].ffill()
        
        # JANGAN forward fill untuk KOTA/KAB - biarkan per baris!
        
        # Bersihkan data
        df = clean_dataframe(df)
        
        # Normalisasi STASIUN NAME (lowercase untuk grouping, tapi simpan original untuk display)
        df['STASIUN NAME_NORM'] = df['STASIUN NAME'].str.lower().str.strip()
        
        # Normalisasi KOTA (lowercase)
        if 'KOTA/KAB' in df.columns:
            df['KOTA/KAB'] = df['KOTA/KAB'].str.lower().str.strip()
        
        # Filter baris yang valid (stasiun lawan tidak boleh kosong)
        df = df[df['STASIUN LAWAN'] != '']
        
        if df.empty:
            return 0, 0, ["❌ File Excel tidak mengandung data stasiun lawan yang valid"]
        
        success_count = 0
        error_count = 0
        messages = []
        
        # ===== PENTING: Grouping berdasarkan STASIUN NAME (case-insensitive) =====
        # Tapi pertahankan semua baris asli untuk stasiun lawan
        grouped_by_station = {}
        
        for idx, row in df.iterrows():
            stasiun_name_norm = row['STASIUN NAME_NORM']
            stasiun_name_original = row['STASIUN NAME']
            
            if stasiun_name_norm not in grouped_by_station:
                grouped_by_station[stasiun_name_norm] = {
                    'original_name': stasiun_name_original,
                    'lawan_list': []
                }
            
            # Tentukan kota yang akan digunakan (per baris!)
            kota_used = None
            
            # 1. Cek dari kolom KOTA/KAB di Excel (per baris)
            if 'KOTA/KAB' in df.columns and pd.notna(row['KOTA/KAB']) and str(row['KOTA/KAB']).strip():
                kota_used = str(row['KOTA/KAB']).strip().lower()
            
            # 2. Jika tidak ada di Excel, cek dari form upload
            if not kota_used and kota_default:
                kota_used = kota_default.lower()
            
            # 3. Jika masih kosong, gunakan default 'samarinda'
            if not kota_used:
                kota_used = 'samarinda'
            
            # Validasi kota (case-insensitive)
            kota_valid = False
            kota_normalized = None
            
            for k in KOTA_LIST:
                if k.lower() == kota_used.lower():
                    kota_normalized = k.lower()
                    kota_valid = True
                    break
            
            if not kota_valid:
                for k in KOTA_LIST:
                    if k.lower() in kota_used.lower() or kota_used.lower() in k.lower():
                        kota_normalized = k.lower()
                        kota_valid = True
                        break
            
            if not kota_valid:
                kota_normalized = 'samarinda'
                messages.append(f"⚠️ Kota '{kota_used}' tidak valid, menggunakan default: samarinda")
            
            # Ambil freq_tx dan freq_rx (opsional - per baris)
            freq_tx = None
            freq_rx = None
            
            if 'FREQ TX' in df.columns and pd.notna(row['FREQ TX']) and str(row['FREQ TX']).strip():
                freq_tx = str(row['FREQ TX']).strip()
            
            if 'FREQ RX' in df.columns and pd.notna(row['FREQ RX']) and str(row['FREQ RX']).strip():
                freq_rx = str(row['FREQ RX']).strip()
            
            grouped_by_station[stasiun_name_norm]['lawan_list'].append({
                'nama': row['STASIUN LAWAN'],
                'kota': kota_normalized,
                'freq_tx': freq_tx,
                'freq_rx': freq_rx
            })
        
        print(f"DEBUG - Total stasiun unik: {len(grouped_by_station)}")
        
        # Process setiap stasiun
        for stasiun_name_norm, station_data in grouped_by_station.items():
            try:
                stasiun_name = station_data['original_name']
                lawan_list = station_data['lawan_list']
                
                # Cek duplikasi stasiun (case-insensitive) - TIDAK MEMBUAT STASIUN BARU JIKA SUD ADA
                existing = Stasiun.query.filter(
                    db.func.lower(Stasiun.stasiun_name) == stasiun_name_norm,
                    Stasiun.operator == operator
                ).first()
                
                if existing:
                    # ===== STASIUN SUDAH ADA - TAMBAHKAN SEMUA STASIUN LAWAN =====
                    # Tidak ada pengecekan duplikat untuk stasiun lawan!
                    urutan = len(existing.stasiun_lawan_list)
                    added_count = 0
                    kota_used_count = {}  # Untuk tracking kota yang digunakan
                    
                    for lawan_data in lawan_list:
                        # Selalu tambahkan stasiun lawan baru, tidak peduli duplikat
                        new_lawan = StasiunLawan(
                            stasiun_id=existing.id,
                            nama_stasiun_lawan=lawan_data['nama'],
                            freq_tx=lawan_data['freq_tx'],
                            freq_rx=lawan_data['freq_rx'],
                            group_id=None,
                            urutan=urutan
                        )
                        db.session.add(new_lawan)
                        urutan += 1
                        added_count += 1
                        
                        # Track kota yang digunakan
                        kota = lawan_data['kota']
                        kota_used_count[kota] = kota_used_count.get(kota, 0) + 1
                    
                    # OPTIONAL: Update kota stasiun jika diperlukan (misal kota default)
                    # Tapi karena kota bisa berbeda per lawan, kita tidak update stasiun utama
                    
                    db.session.commit()
                    
                    messages.append(f"✅ Stasiun '{stasiun_name}': ditambah {added_count} stasiun lawan")
                    
                    # Tampilkan info kota
                    if kota_used_count:
                        kota_info = ', '.join([f"{k} ({c})" for k, c in kota_used_count.items()])
                        messages.append(f"   📍 Kota lawan: {kota_info}")
                    
                    success_count += 1
                    continue
                
                # ===== STASIUN BARU =====
                # Untuk stasiun baru, kita perlu menentukan kota stasiun utama
                # Ambil kota paling umum dari lawan, atau gunakan kota default
                kota_counts = {}
                for lawan_data in lawan_list:
                    kota = lawan_data['kota']
                    kota_counts[kota] = kota_counts.get(kota, 0) + 1
                
                if kota_counts:
                    # Ambil kota dengan jumlah terbanyak
                    kota_stasiun = max(kota_counts.items(), key=lambda x: x[1])[0]
                else:
                    kota_stasiun = 'samarinda'
                
                # Buat stasiun baru
                new_station = Stasiun(
                    stasiun_name=stasiun_name,
                    operator=operator,
                    kota=kota_stasiun,
                    created_by=user.id
                )
                db.session.add(new_station)
                db.session.flush()
                
                # Tambah SEMUA stasiun lawan (termasuk duplikat)
                lawan_added = 0
                for i, lawan_data in enumerate(lawan_list):
                    if lawan_data['nama']:
                        new_lawan = StasiunLawan(
                            stasiun_id=new_station.id,
                            nama_stasiun_lawan=lawan_data['nama'],
                            freq_tx=lawan_data['freq_tx'],
                            freq_rx=lawan_data['freq_rx'],
                            group_id=None,
                            urutan=i
                        )
                        db.session.add(new_lawan)
                        lawan_added += 1
                
                success_count += 1
                db.session.commit()
                
                messages.append(f"✅ Stasiun baru '{stasiun_name}' ditambahkan dengan {lawan_added} lawan")
                
                # Tampilkan info kota
                if kota_counts:
                    kota_info = ', '.join([f"{k} ({c})" for k, c in kota_counts.items()])
                    messages.append(f"   📍 Kota stasiun utama: {kota_stasiun}")
                    messages.append(f"   📍 Distribusi kota lawan: {kota_info}")
                
            except Exception as e:
                db.session.rollback()
                error_count += 1
                messages.append(f"❌ Error pada stasiun '{stasiun_name}': {str(e)}")
                import traceback
                traceback.print_exc()
                continue
        
        summary = f"✅ Berhasil memproses {success_count} stasiun untuk operator {operator.upper()}!"
        if error_count > 0:
            summary += f" ⚠️ {error_count} error ditemukan."
        
        messages.insert(0, summary)
        return success_count, error_count, messages
        
    except Exception as e:
        db.session.rollback()
        print(f"ERROR FATAL: {str(e)}")
        import traceback
        traceback.print_exc()
        return 0, 0, [f"❌ Error processing file: {str(e)}"]
    
def generate_admin_operator_template(operator):
    """Generate template Excel untuk admin operator dengan FREQ TX dan FREQ RX opsional"""
    wb = Workbook()
    ws = wb.active
    ws.title = "Template"
    
    if 'Sheet' in wb.sheetnames:
        std = wb['Sheet']
        wb.remove(std)
    
    # Tentukan operator name
    operator_name = operator.upper()
    
    # Header tabel - 5 KOLOM (tanpa tanda kurung)
    headers = ['STASIUN NAME', 'STASIUN LAWAN', 'FREQ TX', 'FREQ RX', 'KOTA/KAB']
    ws.append(headers)
    
    # Style header
    for cell in ws[1]:
        cell.font = Font(bold=True, size=11)
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.fill = PatternFill(start_color="C6E0B4", end_color="C6E0B4", fill_type="solid")
    
    # =================== DATA CONTOH ===================
    contoh_data = [
        # STASIUN NAME, STASIUN LAWAN, FREQ TX, FREQ RX, KOTA/KAB
        [f'{operator_name}_STN_01', f'{operator_name}_LAWAN_01', '1800 MHz', '1800 MHz', 'Kota Samarinda'],
        ['', f'{operator_name}_LAWAN_02', '', '', ''],
        ['', f'{operator_name}_LAWAN_03', '900 MHz', '900 MHz', ''],
        
        [f'{operator_name}_STN_02', f'{operator_name}_LAWAN_04', '', '', 'Kota Balikpapan'],
        
        [f'{operator_name}_STN_03', f'{operator_name}_LAWAN_05', '2100 MHz', '2100 MHz', 'Kota Bontang'],
        ['', f'{operator_name}_LAWAN_06', '', '', ''],
    ]
    
    for row in contoh_data:
        ws.append(row)
    
    # =================== MERGE CELL ===================
    if ws.max_row >= 3:
        ws.merge_cells(start_row=2, start_column=1, end_row=4, end_column=1)  # STASIUN NAME
        ws.merge_cells(start_row=2, start_column=5, end_row=4, end_column=5)  # KOTA/KAB
    
    if ws.max_row >= 6:
        ws.merge_cells(start_row=5, start_column=1, end_row=6, end_column=1)  # STASIUN NAME
        ws.merge_cells(start_row=5, start_column=5, end_row=6, end_column=5)  # KOTA/KAB
    
    # Center alignment untuk merge cells
    merge_cells = ['A2', 'E2', 'A5', 'E5']
    for cell_ref in merge_cells:
        if cell_ref in ws:
            ws[cell_ref].alignment = Alignment(vertical='center', horizontal='center')
    
    # Format lebar kolom
    ws.column_dimensions['A'].width = 25  # STASIUN NAME
    ws.column_dimensions['B'].width = 25  # STASIUN LAWAN
    ws.column_dimensions['C'].width = 15  # FREQ TX
    ws.column_dimensions['D'].width = 15  # FREQ RX
    ws.column_dimensions['E'].width = 20  # KOTA/KAB
    
    # =================== SHEET INSTRUKSI ===================
    ws_inst = wb.create_sheet(title="Instruksi")
    
    instruksi = [
        [f"INSTRUKSI PENGISIAN - OPERATOR {operator_name}"],
        ["----------------------------------------"],
        [f"PERHATIAN: File ini khusus untuk operator {operator_name}"],
        [""],
        ["FORMAT KOLOM:"],
        ["1. STASIUN NAME  : Nama stasiun utama Anda - WAJIB"],
        ["2. STASIUN LAWAN : Nama stasiun lawan/kompetitor - WAJIB"],
        ["3. FREQ TX       : Frekuensi Tx - OPSIONAL (boleh dikosongkan)"],
        ["4. FREQ RX       : Frekuensi Rx - OPSIONAL (boleh dikosongkan)"],
        ["5. KOTA/KAB      : Nama kota/kabupaten - WAJIB"],
        [""],
        ["CATATAN PENTING:"],
        ["- Kolom FREQ TX dan FREQ RX bersifat OPSIONAL"],
        ["- Boleh dikosongkan jika belum ada data frekuensi"],
        ["- Data frekuensi bisa diisi nanti melalui halaman edit"],
        [""],
        ["CARA PENGISIAN:"],
        ["1. STASIUN NAME:"],
        ["   - Tulis sekali untuk setiap stasiun dengan banyak lawan"],
        ["   - Gunakan merge cell jika diperlukan"],
        ["   - Contoh: " + operator_name + "_STN_01, TOWER_01"],
        [""],
        ["2. STASIUN LAWAN:"],
        ["   - Satu baris untuk satu lawan"],
        ["   - Tidak perlu merge cell"],
        ["   - Contoh: " + operator_name + "_LAWAN_01, COMPETITOR_01"],
        [""],
        ["3. FREQ TX (Opsional):"],
        ["   - Bisa dikosongkan jika tidak ada data"],
        ["   - Contoh: 1800 MHz, 2100 MHz, 900 MHz"],
        [""],
        ["4. FREQ RX (Opsional):"],
        ["   - Bisa dikosongkan jika tidak ada data"],
        ["   - Contoh: 1800 MHz, 2100 MHz, 900 MHz"],
        [""],
        ["5. KOTA/KAB:"],
        ["   - Tulis sekali untuk setiap stasiun"],
        ["   - Gunakan merge cell seperti STASIUN NAME"],
    ]
    
    for row in instruksi:
        ws_inst.append(row)
    
    ws_inst.column_dimensions['A'].width = 80
    ws_inst['A1'].font = Font(bold=True, size=14)
    
    # =================== SHEET DAFTAR KOTA ===================
    ws_kota = wb.create_sheet(title="Daftar Kota")
    
    ws_kota.append(["DAFTAR KOTA/KAB UNTUK OPERATOR " + operator_name])
    ws_kota.append(["---------------------------------"])
    ws_kota.append(["Gunakan nama persis seperti di bawah:"])
    ws_kota.append([""])
    
    kota_list_display = [
        "Kota Samarinda", "Kota Balikpapan", "Kota Bontang", "Kutai Kartanegara", 
        "Kutai Barat", "Kutai Timur", "Penajam Paser Utara", "Paser",
        "Berau", "Mahakam Ulu"
    ]
    
    for kota in sorted(kota_list_display, key=lambda x: x.lower()):
        ws_kota.append([kota])
    
    ws_kota.column_dimensions['A'].width = 30
    ws_kota['A1'].font = Font(bold=True, size=12)
    
    # Atur sheet order
    wb.active = wb['Template']
    
    # =================== SAVE FILE ===================
    output = BytesIO()
    wb.save(output)
    output.seek(0)
    
    return send_file(
        output,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True,
        download_name=f'Template_{operator_name}.xlsx'
    )

# ============================================================================
# REKAP BULANAN DAN TAHUNAN 
# ============================================================================
@app.route('/admin-master/rekap-bulanan-word')
@login_required
@admin_master_required
def admin_master_rekap_bulanan_word():
    """Export rekap bulanan ke Word dengan gambar"""
    try:
        tahun = request.args.get('tahun', datetime.now().year, type=int)
        bulan = request.args.get('bulan', datetime.now().month, type=int)
        operator = request.args.get('operator', 'all')
        kota = request.args.get('kota', 'all')
        
        # Query stasiun
        query = Stasiun.query
        
        if operator != 'all':
            query = query.filter(Stasiun.operator == operator)
        
        if kota != 'all':
            query = query.filter(Stasiun.kota == kota)
        
        stations = query.order_by(Stasiun.operator, Stasiun.kota, Stasiun.stasiun_name).all()
        
        if not stations:
            flash('Tidak ada data untuk periode yang dipilih', 'warning')
            return redirect(url_for('admin_master_rekap_bulanan'))
        
        # Buat dokumen Word
        doc = Document()
        
        # ===== HEADER DOKUMEN =====
        # Judul utama
        title = doc.add_heading('REKAP BULANAN STASIUN MICROWAVE LINK', 0)
        title.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        # Periode
        bulan_nama = ['Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni',
                     'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember']
        periode = doc.add_heading(f'BULAN: {bulan_nama[bulan-1]} {tahun}', level=1)
        periode.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        # Filter info
        filter_text = f"Operator: {operator.upper() if operator != 'all' else 'SEMUA OPERATOR'} | "
        filter_text += f"Kota/Kab: {kota.title() if kota != 'all' else 'Semua Kota/kab'}"
        filter_para = doc.add_paragraph(filter_text)
        filter_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        filter_para.runs[0].bold = True
        
        doc.add_paragraph()  # Spasi
        
        # ===== TABEL REKAP =====
        # Header tabel
        table = doc.add_table(rows=1, cols=5)
        table.style = 'Table Grid'
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        
        # Set lebar kolom
        table.columns[0].width = Inches(1.0)   # No
        table.columns[1].width = Inches(1.5)   # Operator
        table.columns[2].width = Inches(1.8)   # Stasiun Name
        table.columns[3].width = Inches(1.5)   # Kota
        table.columns[4].width = Inches(3.2)   # Stasiun Lawan & Status
        
        # Header row
        header_cells = table.rows[0].cells
        headers = ['No', 'Operator', 'Stasiun Name', 'Kota/Kab', 'Stasiun Lawan & Status']
        for i, header in enumerate(headers):
            header_cells[i].text = header
            header_cells[i].paragraphs[0].runs[0].font.bold = True
            header_cells[i].paragraphs[0].runs[0].font.size = Pt(10)
            header_cells[i].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        # Isi tabel
        no = 1
        total_aktif = 0
        total_tidak_aktif = 0
        total_tidak_sesuai = 0
        total_belum_ada = 0
        
        for station in stations:
            # Ambil semua lawan
            opponents = StasiunLawan.query.filter_by(stasiun_id=station.id).order_by(StasiunLawan.urutan).all()
            
            # Hitung jumlah baris yang diperlukan
            row_count = max(len(opponents), 1)  # Minimal 1 baris
            
            # Tambahkan baris sesuai jumlah lawan
            for i in range(row_count):
                if i == 0:
                    # Baris pertama dengan data stasiun
                    row_cells = table.add_row().cells
                    
                    # No
                    row_cells[0].text = str(no)
                    row_cells[0].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                    
                    # Operator
                    row_cells[1].text = station.operator.upper()
                    row_cells[1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                    
                    # Stasiun Name
                    row_cells[2].text = station.stasiun_name
                    row_cells[2].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.LEFT
                    
                    # Kota
                    row_cells[3].text = station.kota
                    row_cells[3].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.LEFT
                    
                    # Stasiun Lawan & Status (akan diisi per lawan)
                    if opponents:
                        lawan = opponents[i]
                        # Dapatkan status terbaru untuk bulan ini
                        latest_status = StatusUpdate.query\
                            .filter_by(stasiun_lawan_id=lawan.id)\
                            .filter(extract('year', StatusUpdate.updated_at) == tahun)\
                            .filter(extract('month', StatusUpdate.updated_at) == bulan)\
                            .order_by(StatusUpdate.updated_at.desc())\
                            .first()
                        
                        status_text = lawan.nama_stasiun_lawan
                        if latest_status:
                            status_text += f"\nStatus: {STATUS_OPTIONS.get(latest_status.status, latest_status.status)}"
                            if latest_status.catatan:
                                status_text += f"\nCatatan: {latest_status.catatan[:50]}"
                            
                            # Hitung total
                            if latest_status.status == 'aktif':
                                total_aktif += 1
                            elif latest_status.status == 'tidak_aktif':
                                total_tidak_aktif += 1
                            elif latest_status.status == 'tidak_sesuai':
                                total_tidak_sesuai += 1
                        else:
                            status_text += f"\nStatus: Belum Ada"
                            total_belum_ada += 1
                        
                        row_cells[4].text = status_text
                        row_cells[4].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.LEFT
                    else:
                        row_cells[4].text = "Tidak ada stasiun lawan"
                        row_cells[4].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                else:
                    # Baris berikutnya (hanya untuk stasiun lawan tambahan)
                    if opponents and i < len(opponents):
                        row_cells = table.add_row().cells
                        
                        # Kolom 1-4 kosong (merge secara visual)
                        row_cells[0].text = ""
                        row_cells[1].text = ""
                        row_cells[2].text = ""
                        row_cells[3].text = ""
                        
                        # Kolom 5 diisi lawan berikutnya
                        lawan = opponents[i]
                        latest_status = StatusUpdate.query\
                            .filter_by(stasiun_lawan_id=lawan.id)\
                            .filter(extract('year', StatusUpdate.updated_at) == tahun)\
                            .filter(extract('month', StatusUpdate.updated_at) == bulan)\
                            .order_by(StatusUpdate.updated_at.desc())\
                            .first()
                        
                        status_text = lawan.nama_stasiun_lawan
                        if latest_status:
                            status_text += f"\nStatus: {STATUS_OPTIONS.get(latest_status.status, latest_status.status)}"
                            if latest_status.catatan:
                                status_text += f"\nCatatan: {latest_status.catatan[:50]}"
                            
                            if latest_status.status == 'aktif':
                                total_aktif += 1
                            elif latest_status.status == 'tidak_aktif':
                                total_tidak_aktif += 1
                            elif latest_status.status == 'tidak_sesuai':
                                total_tidak_sesuai += 1
                        else:
                            status_text += f"\nStatus: Belum Ada"
                            total_belum_ada += 1
                        
                        row_cells[4].text = status_text
                        row_cells[4].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.LEFT
            
            no += 1
        
        doc.add_paragraph()  # Spasi
        
        # ===== REKAPITULASI STATUS =====
        summary_table = doc.add_table(rows=2, cols=6)
        summary_table.style = 'Table Grid'
        summary_table.alignment = WD_TABLE_ALIGNMENT.CENTER
        
        # Header summary
        summary_headers = summary_table.rows[0].cells
        headers = ['Aktif', 'Tidak Aktif', 'Tidak Sesuai', 'Belum Ada', 'Total']
        for i, header in enumerate(headers):
            summary_headers[i].text = header
            summary_headers[i].paragraphs[0].runs[0].font.bold = True
            summary_headers[i].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        # Isi summary
        summary_values = summary_table.rows[1].cells
        total_semua = total_aktif + total_tidak_aktif + total_tidak_sesuai + total_belum_ada
        values = [total_aktif, total_tidak_aktif, total_tidak_sesuai, total_belum_ada, total_semua]
        for i, value in enumerate(values):
            summary_values[i].text = str(value)
            summary_values[i].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            if i < 5:  # Warna untuk masing-masing status
                if i == 0:  # Aktif
                    summary_values[i].paragraphs[0].runs[0].font.bold = True
                elif i == 1:  # Tidak Aktif
                    summary_values[i].paragraphs[0].runs[0].font.bold = True
        
        doc.add_paragraph()  # Spasi
        
        # ===== FOOTER =====
        footer = doc.add_paragraph()
        footer.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        footer.add_run(f'Dicetak pada: {datetime.now().strftime("%d/%m/%Y")}').italic = True
        footer.add_run('\n').italic = True
        footer.add_run(f'Oleh: {current_user.username}').italic = True
        
        # ===== GAMBAR UNTUK SETIAP STASIUN =====
        doc.add_page_break()
        doc.add_heading('DOKUMENTASI GAMBAR', level=1).alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        for station in stations:
            # Heading stasiun
            doc.add_heading(f'Stasiun: {station.stasiun_name} ({station.operator.upper()})', level=2)
            
            # Ambil semua upload UNIK untuk stasiun ini
            uploads_query = db.session.query(
                UploadGambar.public_id,
                UploadGambar.group_id,
                db.func.max(UploadGambar.id).label('latest_id')
            ).filter_by(stasiun_id=station.id)\
             .group_by(UploadGambar.public_id, UploadGambar.group_id)\
             .subquery()
            
            uploads = UploadGambar.query\
                .join(uploads_query, UploadGambar.id == uploads_query.c.latest_id)\
                .order_by(UploadGambar.uploaded_at.desc())\
                .limit(10)\
                .all()
            
            if uploads:
                # Tabel gambar (maks 3 kolom)
                img_table = doc.add_table(rows=1, cols=3)
                img_table.style = 'Table Grid'
                
                row_cells = img_table.rows[0].cells
                col = 0
                
                for upload in uploads:
                    if col >= 3:
                        # Tambah baris baru
                        img_table.add_row()
                        row_cells = img_table.rows[-1].cells
                        col = 0
                    
                    # Download gambar dari Cloudinary
                    try:
                        response = requests.get(upload.cloudinary_url, timeout=10)
                        if response.status_code == 200:
                            img_data = BytesIO(response.content)
                            
                            # Tambah gambar ke cell
                            paragraph = row_cells[col].paragraphs[0]
                            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                            run = paragraph.add_run()
                            run.add_picture(img_data, width=Inches(1.8))
                            
                            # Tambah keterangan
                            paragraph.add_run().add_break()
                            keterangan = f"{upload.original_filename[:30]}"
                            if upload.original_filename and len(upload.original_filename) > 30:
                                keterangan += "..."
                            paragraph.add_run(keterangan).font.size = Pt(8)
                            
                            # Tambah status
                            paragraph.add_run().add_break()
                            status_text = f"Status: {upload.status if upload.status else 'Belum Ada'}"
                            paragraph.add_run(status_text).font.size = Pt(8)
                            
                            # Tambah tanggal
                            paragraph.add_run().add_break()
                            tgl_text = f"Upload: {upload.uploaded_at.strftime('%d/%m/%Y')}"
                            paragraph.add_run(tgl_text).font.size = Pt(8)
                    except Exception as e:
                        print(f"Error downloading image: {e}")
                        row_cells[col].text = "Gambar tidak tersedia"
                    
                    col += 1
                
                doc.add_paragraph()  # Spasi antar stasiun
            else:
                doc.add_paragraph("Tidak ada gambar untuk stasiun ini")
                doc.add_paragraph()
        
        # ===== SIMPAN DOKUMEN KE MEMORY =====
        file_stream = BytesIO()
        doc.save(file_stream)
        file_stream.seek(0)
        
        # Nama file
        filename = f"Rekap_Bulanan_{bulan_nama[bulan-1]}_{tahun}"
        if operator != 'all':
            filename += f"_{operator.upper()}"
        filename += ".docx"
        
        # Kirim file ke client
        return send_file(
            file_stream,
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            as_attachment=True,
            download_name=filename
        )
        
    except Exception as e:
        flash(f'Error: {str(e)}', 'error')
        print(f"ERROR in rekap bulanan word: {str(e)}")
        import traceback
        traceback.print_exc()
        return redirect(url_for('admin_master_rekap_bulanan'))


@app.route('/admin-master/rekap-tahunan-word')
@login_required
@admin_master_required
def admin_master_rekap_tahunan_word():
    """Export rekap tahunan ke Word dengan gambar"""
    try:
        tahun = request.args.get('tahun', datetime.now().year, type=int)
        operator = request.args.get('operator', 'all')
        kota = request.args.get('kota', 'all')
        
        # Query stasiun
        query = Stasiun.query
        
        if operator != 'all':
            query = query.filter(Stasiun.operator == operator)
        
        if kota != 'all':
            query = query.filter(Stasiun.kota == kota)
        
        stations = query.order_by(Stasiun.operator, Stasiun.kota, Stasiun.stasiun_name).all()
        
        if not stations:
            flash('Tidak ada data untuk periode yang dipilih', 'warning')
            return redirect(url_for('admin_master_rekap_tahunan'))
        
        # Buat dokumen Word
        doc = Document()
        
        # ===== HEADER DOKUMEN =====
        title = doc.add_heading('REKAP TAHUNAN STASIUN MICROWAVE LINK', 0)
        title.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        periode = doc.add_heading(f'TAHUN: {tahun}', level=1)
        periode.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        filter_text = f"Operator: {operator.upper() if operator != 'all' else 'SEMUA OPERATOR'} | "
        filter_text += f"Kota/Kab: {kota.title() if kota != 'all' else 'Semua Kota/kab'}"
        filter_para = doc.add_paragraph(filter_text)
        filter_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        filter_para.runs[0].bold = True
        
        doc.add_paragraph()
        
        # ===== DATA PER BULAN =====
        bulan_nama = ['Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni',
                     'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember']
        
        for bulan in range(1, 13):
            doc.add_heading(f'BULAN: {bulan_nama[bulan-1]} {tahun}', level=2)
            
            # Tabel per bulan
            table = doc.add_table(rows=1, cols=4)
            table.style = 'Table Grid'
            
            # Header
            header_cells = table.rows[0].cells
            headers = ['No', 'Operator', 'Stasiun Name', 'Kota/Kab', 'Jumlah Stasiun Lawan & Status']
            for i, header in enumerate(headers):
                header_cells[i].text = header
                header_cells[i].paragraphs[0].runs[0].font.bold = True
                header_cells[i].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            
            # Isi per stasiun
            no = 1
            total_aktif_bulan = 0
            total_tidak_aktif_bulan = 0
            total_tidak_sesuai_bulan = 0
            total_belum_ada_bulan = 0
            
            for station in stations:
                opponents = StasiunLawan.query.filter_by(stasiun_id=station.id).all()
                
                # Hitung status untuk bulan ini
                aktif_count = 0
                tidak_aktif_count = 0
                tidak_sesuai_count = 0
                belum_ada_count = 0
                
                for opp in opponents:
                    latest_status = StatusUpdate.query\
                        .filter_by(stasiun_lawan_id=opp.id)\
                        .filter(extract('year', StatusUpdate.updated_at) == tahun)\
                        .filter(extract('month', StatusUpdate.updated_at) == bulan)\
                        .order_by(StatusUpdate.updated_at.desc())\
                        .first()
                    
                    if latest_status:
                        if latest_status.status == 'aktif':
                            aktif_count += 1
                        elif latest_status.status == 'tidak_aktif':
                            tidak_aktif_count += 1
                        elif latest_status.status == 'tidak_sesuai':
                            tidak_sesuai_count += 1
                    else:
                        belum_ada_count += 1
                
                row_cells = table.add_row().cells
                
                row_cells[0].text = str(no)
                row_cells[0].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                
                row_cells[1].text = station.operator.upper()
                row_cells[1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                
                row_cells[2].text = station.stasiun_name
                row_cells[2].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.LEFT
                
                row_cells[3].text = station.kota
                row_cells[3].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.LEFT
                
                status_text = f"Total: {len(opponents)} stasiun lawan\n"
                status_text += f"Aktif: {aktif_count}\n"
                status_text += f"Tidak Aktif: {tidak_aktif_count}\n"
                status_text += f"Tidak Sesuai: {tidak_sesuai_count}\n"
                status_text += f"Belum Ada: {belum_ada_count}"
                
                row_cells[4].text = status_text
                row_cells[4].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.LEFT
                
                no += 1
                
                total_aktif_bulan += aktif_count
                total_tidak_aktif_bulan += tidak_aktif_count
                total_tidak_sesuai_bulan += tidak_sesuai_count
                total_belum_ada_bulan += belum_ada_count
            
            # Summary bulan
            doc.add_paragraph()
            summary = doc.add_paragraph()
            summary.add_run(f"Rekapitulasi {bulan_nama[bulan-1]} {tahun}:\n").bold = True
            summary.add_run(f"Aktif: {total_aktif_bulan} | ")
            summary.add_run(f"Tidak Aktif: {total_tidak_aktif_bulan} | ")
            summary.add_run(f"Tidak Sesuai: {total_tidak_sesuai_bulan} | ")
            summary.add_run(f"Belum Ada: {total_belum_ada_bulan}")
            
            doc.add_paragraph()
        
        # ===== REKAPITULASI TAHUNAN =====
        doc.add_page_break()
        doc.add_heading('REKAPITULASI TAHUNAN', level=1).alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        # Tabel rekapitulasi tahunan
        summary_table = doc.add_table(rows=13, cols=6)  # 12 bulan + header
        summary_table.style = 'Table Grid'
        
        # Header
        header_cells = summary_table.rows[0].cells
        headers = ['Bulan', 'Aktif', 'Tidak Aktif', 'Tidak Sesuai', 'Total']
        for i, header in enumerate(headers):
            header_cells[i].text = header
            header_cells[i].paragraphs[0].runs[0].font.bold = True
            header_cells[i].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        # Isi per bulan
        for bulan in range(1, 13):
            row_cells = summary_table.rows[bulan].cells
            
            row_cells[0].text = bulan_nama[bulan-1]
            row_cells[0].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            
            # Hitung total per bulan (contoh sederhana)
            # Di sini Anda bisa menghitung ulang atau menggunakan data yang sudah dikumpulkan
            # Untuk contoh, kita isi dengan 0
            for col in range(1, 6):
                row_cells[col].text = "0"
                row_cells[col].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        doc.add_paragraph()
        
        # ===== FOOTER =====
        footer = doc.add_paragraph()
        footer.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        footer.add_run(f'Dicetak pada: {datetime.now().strftime("%d/%m/%Y")}').italic = True
        footer.add_run('\n').italic = True
        footer.add_run(f'Oleh: {current_user.username}').italic = True
        
        # ===== GAMBAR UNTUK SETIAP STASIUN =====
        doc.add_page_break()
        doc.add_heading('DOKUMENTASI GAMBAR', level=1).alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        for station in stations:
            doc.add_heading(f'Stasiun: {station.stasiun_name} ({station.operator.upper()})', level=2)
            
            uploads_query = db.session.query(
                UploadGambar.public_id,
                UploadGambar.group_id,
                db.func.max(UploadGambar.id).label('latest_id')
            ).filter_by(stasiun_id=station.id)\
             .group_by(UploadGambar.public_id, UploadGambar.group_id)\
             .subquery()
            
            uploads = UploadGambar.query\
                .join(uploads_query, UploadGambar.id == uploads_query.c.latest_id)\
                .order_by(UploadGambar.uploaded_at.desc())\
                .limit(10)\
                .all()
            
            if uploads:
                img_table = doc.add_table(rows=1, cols=3)
                img_table.style = 'Table Grid'
                
                row_cells = img_table.rows[0].cells
                col = 0
                
                for upload in uploads:
                    if col >= 3:
                        img_table.add_row()
                        row_cells = img_table.rows[-1].cells
                        col = 0
                    
                    try:
                        response = requests.get(upload.cloudinary_url, timeout=10)
                        if response.status_code == 200:
                            img_data = BytesIO(response.content)
                            
                            paragraph = row_cells[col].paragraphs[0]
                            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                            run = paragraph.add_run()
                            run.add_picture(img_data, width=Inches(1.8))
                            
                            paragraph.add_run().add_break()
                            keterangan = f"{upload.original_filename[:30]}"
                            if upload.original_filename and len(upload.original_filename) > 30:
                                keterangan += "..."
                            paragraph.add_run(keterangan).font.size = Pt(8)
                            
                            paragraph.add_run().add_break()
                            status_text = f"Status: {upload.status if upload.status else 'Belum Ada'}"
                            paragraph.add_run(status_text).font.size = Pt(8)
                            
                            paragraph.add_run().add_break()
                            tgl_text = f"Upload: {upload.uploaded_at.strftime('%d/%m/%Y')}"
                            paragraph.add_run(tgl_text).font.size = Pt(8)
                    except:
                        row_cells[col].text = "Gambar tidak tersedia"
                    
                    col += 1
                
                doc.add_paragraph()
            else:
                doc.add_paragraph("Tidak ada gambar untuk stasiun ini")
                doc.add_paragraph()
        
        # ===== SIMPAN DOKUMEN =====
        file_stream = BytesIO()
        doc.save(file_stream)
        file_stream.seek(0)
        
        filename = f"Rekap_Tahunan_{tahun}"
        if operator != 'all':
            filename += f"_{operator.upper()}"
        filename += ".docx"
        
        return send_file(
            file_stream,
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            as_attachment=True,
            download_name=filename
        )
        
    except Exception as e:
        flash(f'Error: {str(e)}', 'error')
        print(f"ERROR in rekap tahunan word: {str(e)}")
        import traceback
        traceback.print_exc()
        return redirect(url_for('admin_master_rekap_tahunan'))

@app.route('/admin-master/update-status-pemeriksaan/<int:station_id>', methods=['POST'])
@login_required
@admin_master_required
def admin_master_update_status_pemeriksaan(station_id):
    """Update status pemeriksaan stasiun untuk Admin Master"""
    try:
        station = Stasiun.query.get_or_404(station_id)
        data = request.get_json()
        status = data.get('status')
        
        if status not in ['sudah_diperiksa', 'belum_diperiksa']:
            return jsonify({'success': False, 'message': 'Status tidak valid'})
        
        # Update status
        station.status_pemeriksaan = status
        
        if status == 'sudah_diperiksa':
            station.checked_at = datetime.utcnow()
            station.checked_by = current_user.id
        else:
            station.checked_at = None
            station.checked_by = None
        
        db.session.commit()
        
        # Ambil informasi checker untuk response
        checker_name = current_user.username if status == 'sudah_diperiksa' else None
        checked_at_str = station.checked_at.strftime('%d/%m/%Y') if station.checked_at else None
        
        return jsonify({
            'success': True,
            'message': f'Status pemeriksaan berhasil diupdate menjadi "{status}"',
            'checked_at': checked_at_str,
            'checked_by': checker_name
        })
        
    except Exception as e:
        db.session.rollback()
        return jsonify({'success': False, 'message': str(e)})     

# ============================================================================
# ADMIN MASTER - HASIL LAPORAN
# ============================================================================
@app.route('/admin-master/hasil-laporan', methods=['GET', 'POST'])
@login_required
@admin_master_required
def admin_master_hasil_laporan():
    """Halaman input dan daftar hasil laporan untuk Admin Master"""
    
    if request.method == 'POST':
        try:
            operator = request.form.get('operator')
            nomor_ba = request.form.get('nomor_ba', '').strip()
            tanggal_ba_str = request.form.get('tanggal_ba')
            nomor_spt = request.form.get('nomor_spt', '').strip()
            kota_kab = request.form.get('kota_kab')
            
            # Validasi semua field wajib
            if not all([operator, nomor_ba, tanggal_ba_str, nomor_spt, kota_kab]):
                flash('Semua field wajib harus diisi!', 'error')
                return redirect(url_for('admin_master_hasil_laporan'))
            
            # Cek duplikasi nomor BA
            existing = HasilLaporan.query.filter_by(nomor_ba=nomor_ba).first()
            if existing:
                flash(f'Nomor BA "{nomor_ba}" sudah ada!', 'error')
                return redirect(url_for('admin_master_hasil_laporan'))
            
            # Parse tanggal
            tanggal_ba = datetime.strptime(tanggal_ba_str, '%Y-%m-%d').date()
            
            # Proses upload PDF
            file_pdf_url = None
            file_pdf_path = None

            if 'file_pdf' in request.files:
                pdf_file = request.files['file_pdf']
                if pdf_file and pdf_file.filename:
                    if not pdf_file.filename.lower().endswith('.pdf'):
                        flash('File harus berformat PDF!', 'error')
                        return redirect(url_for('admin_master_hasil_laporan'))
                    
                    result, error = upload_pdf_to_local(pdf_file, operator, nomor_ba)
                    if result:
                        file_pdf_url = result['url']
                        file_pdf_path = result['file_path']
                    else:
                        flash(f'Gagal upload PDF: {error}', 'error')
                        return redirect(url_for('admin_master_hasil_laporan'))

            # Simpan ke database
            new_laporan = HasilLaporan(
                operator=operator,
                nomor_ba=nomor_ba,
                tanggal_ba=tanggal_ba,
                nomor_spt=nomor_spt,
                kota_kab=kota_kab,
                file_pdf_url=file_pdf_url,
                file_pdf_path=file_pdf_path,
                created_by=current_user.id
            )
            
            db.session.add(new_laporan)
            db.session.commit()
            
            flash(f'✅ Laporan dengan Nomor BA "{nomor_ba}" berhasil ditambahkan!', 'success')
            return redirect(url_for('admin_master_hasil_laporan'))
            
        except Exception as e:
            db.session.rollback()
            flash(f'❌ Error: {str(e)}', 'error')
            print(f"Error in admin_master_hasil_laporan POST: {str(e)}")
            import traceback
            traceback.print_exc()
            return redirect(url_for('admin_master_hasil_laporan'))
    
    # ========== GET REQUEST - TAMPILKAN DAFTAR LAPORAN ==========
    page = request.args.get('page', 1, type=int)
    per_page = request.args.get('per_page', 10, type=int)
    
    # Ambil semua parameter filter
    filter_operator = request.args.get('filter_operator', 'all')
    filter_kota = request.args.get('filter_kota', 'all')
    search_nomor_ba = request.args.get('search_nomor_ba', '').strip()
    search_nomor_spt = request.args.get('search_nomor_spt', '').strip()
    tanggal_mulai = request.args.get('tanggal_mulai', '').strip()
    tanggal_sampai = request.args.get('tanggal_sampai', '').strip()
    
    # Buat query dasar
    query = HasilLaporan.query
    
    # Filter berdasarkan operator
    if filter_operator != 'all':
        query = query.filter(HasilLaporan.operator == filter_operator)
    
    # Filter berdasarkan kota/kab
    if filter_kota != 'all':
        query = query.filter(HasilLaporan.kota_kab == filter_kota)
    
    # Filter pencarian Nomor BA (partial match)
    if search_nomor_ba:
        query = query.filter(HasilLaporan.nomor_ba.ilike(f'%{search_nomor_ba}%'))
    
    # Filter pencarian Nomor SPT (partial match)
    if search_nomor_spt:
        query = query.filter(HasilLaporan.nomor_spt.ilike(f'%{search_nomor_spt}%'))
    
    # Filter rentang tanggal mulai
    if tanggal_mulai:
        try:
            tanggal_mulai_date = datetime.strptime(tanggal_mulai, '%Y-%m-%d').date()
            query = query.filter(HasilLaporan.tanggal_ba >= tanggal_mulai_date)
        except ValueError:
            flash(f'Format tanggal mulai tidak valid: {tanggal_mulai}', 'warning')
    
    # Filter rentang tanggal sampai
    if tanggal_sampai:
        try:
            tanggal_sampai_date = datetime.strptime(tanggal_sampai, '%Y-%m-%d').date()
            query = query.filter(HasilLaporan.tanggal_ba <= tanggal_sampai_date)
        except ValueError:
            flash(f'Format tanggal sampai tidak valid: {tanggal_sampai}', 'warning')
    
    # Urutkan berdasarkan tanggal dibuat terbaru
    query = query.order_by(HasilLaporan.created_at.desc())
    
    # Pagination
    pagination = query.paginate(page=page, per_page=per_page, error_out=False)
    laporan_list = pagination.items
    
    # Ambil daftar operator yang pernah membuat laporan (untuk dropdown filter)
    operators_with_laporan = db.session.query(HasilLaporan.operator).distinct().all()
    operators_with_laporan = [op[0] for op in operators_with_laporan if op[0]]
    
    # Ambil daftar kota yang pernah ada dalam laporan (untuk dropdown filter)
    kota_with_laporan = db.session.query(HasilLaporan.kota_kab).distinct().all()
    kota_with_laporan = [k[0] for k in kota_with_laporan if k[0]]
    
    return render_template('admin_master/hasil_laporan.html',
                         laporan_list=laporan_list,
                         pagination=pagination,
                         operators=OPERATORS,
                         kota_list=KOTA_LIST,
                         operators_with_laporan=operators_with_laporan,
                         kota_with_laporan=kota_with_laporan,
                         filter_operator=filter_operator,
                         filter_kota=filter_kota)

@app.route('/admin-master/hasil-laporan/edit/<int:laporan_id>', methods=['GET', 'POST'])
@login_required
@admin_master_required
def admin_master_edit_laporan(laporan_id):
    laporan = HasilLaporan.query.get_or_404(laporan_id)
    
    if request.method == 'POST':
        try:
            operator = request.form.get('operator')
            nomor_ba = request.form.get('nomor_ba', '').strip()
            tanggal_ba_str = request.form.get('tanggal_ba')
            nomor_spt = request.form.get('nomor_spt', '').strip()
            kota_kab = request.form.get('kota_kab')
            
            if not all([operator, nomor_ba, tanggal_ba_str, nomor_spt, kota_kab]):
                flash('Semua field wajib harus diisi!', 'error')
                return redirect(url_for('admin_master_edit_laporan', laporan_id=laporan_id))
            
            existing = HasilLaporan.query.filter(
                HasilLaporan.nomor_ba == nomor_ba,
                HasilLaporan.id != laporan_id
            ).first()
            
            if existing:
                flash(f'Nomor BA "{nomor_ba}" sudah digunakan!', 'error')
                return redirect(url_for('admin_master_edit_laporan', laporan_id=laporan_id))
            
            tanggal_ba = datetime.strptime(tanggal_ba_str, '%Y-%m-%d').date()
            
            # ===== INI BAGIAN YANG DIPERBAIKI =====
            file_pdf_url = laporan.file_pdf_url
            file_pdf_path = laporan.file_pdf_path  # ← PAKAI file_pdf_path
            
            if 'file_pdf' in request.files:
                pdf_file = request.files['file_pdf']
                if pdf_file and pdf_file.filename:
                    if not pdf_file.filename.lower().endswith('.pdf'):
                        flash('File harus berformat PDF!', 'error')
                        return redirect(url_for('admin_master_edit_laporan', laporan_id=laporan_id))
                    
                    if laporan.file_pdf_path:
                        delete_pdf_from_local(laporan.file_pdf_path)
                    
                    result, error = upload_pdf_to_local(pdf_file, operator, nomor_ba)  # ← PAKAI upload_pdf_to_local
                    if result:
                        file_pdf_url = result['url']
                        file_pdf_path = result['file_path']  # ← PAKAI file_pdf_path
            
            laporan.operator = operator
            laporan.nomor_ba = nomor_ba
            laporan.tanggal_ba = tanggal_ba
            laporan.nomor_spt = nomor_spt
            laporan.kota_kab = kota_kab
            laporan.file_pdf_url = file_pdf_url
            laporan.file_pdf_path = file_pdf_path  # ← PAKAI file_pdf_path
            laporan.updated_at = datetime.utcnow()
            
            db.session.commit()
            flash(f'✅ Laporan berhasil diperbarui!', 'success')
            return redirect(url_for('admin_master_hasil_laporan'))
            
        except Exception as e:
            db.session.rollback()
            flash(f'❌ Error: {str(e)}', 'error')
            return redirect(url_for('admin_master_edit_laporan', laporan_id=laporan_id))
    
    return render_template('admin_master/edit_laporan.html',
                         laporan=laporan,
                         operators=OPERATORS,
                         kota_list=KOTA_LIST)

@app.route('/admin-master/hasil-laporan/hapus/<int:laporan_id>')
@login_required
@admin_master_required
def admin_master_hapus_laporan(laporan_id):
    laporan = HasilLaporan.query.get_or_404(laporan_id)
    nomor_ba = laporan.nomor_ba
    
    try:
        if laporan.file_pdf_path:
            delete_pdf_from_local(laporan.file_pdf_path)
        
        db.session.delete(laporan)
        db.session.commit()
        flash(f'✅ Laporan dengan Nomor BA "{nomor_ba}" berhasil dihapus!', 'success')
    except Exception as e:
        db.session.rollback()
        flash(f'❌ Error: {str(e)}', 'error')
    
    return redirect(url_for('admin_master_hasil_laporan'))   

# ============================================================================
# TEMPLATE FILTERS
# ============================================================================

@app.template_filter('format_datetime')
def format_datetime_filter(value, format='%d/%m/%Y'):
    if isinstance(value, datetime):
        return value.strftime(format)
    return value

@app.template_filter('status_badge_class')
def status_badge_class_filter(status):
    """Return CSS class untuk badge status"""
    classes = {
        'aktif': 'bg-success',
        'tidak_aktif': 'bg-danger',
        'tidak_sesuai': 'bg-info',
        'riwayat': 'bg-secondary'
    }
    return classes.get(status, 'bg-secondary')

@app.template_filter('status_display')
def status_display_filter(status_key):
    """Return display text untuk status"""
    display_text = STATUS_OPTIONS.get(status_key, status_key)
    if status_key == 'riwayat':
        return 'Riwayat'
    return display_text

@app.template_filter('operator_color')
def operator_color_filter(operator):
    """Return warna untuk operator"""
    colors = {
    'telkom': '#E1251B',
    'telkomsel': '#E1251B',
    'ioh': '#FF6600',
    'xl smart': '#7F3F98'
}
    return colors.get(operator, '#6c757d')

@app.template_filter('flash_category_class')
def flash_category_class_filter(category):
    """Return CSS class untuk flash message category"""
    return FLASH_CATEGORIES.get(category, 'alert-info')

@app.template_filter('flash_icon')
def flash_icon_filter(category):
    """Return icon untuk flash message category"""
    icons = {
        'success': 'bi-check-circle-fill',
        'error': 'bi-exclamation-triangle-fill',
        'warning': 'bi-exclamation-triangle-fill',
        'info': 'bi-info-circle-fill',
        'debug': 'bi-bug-fill'
    }
    return icons.get(category, 'bi-info-circle-fill')

# ============================================================================
# CLOUDINARY TEMPLATE FILTERS
# ============================================================================

@app.template_filter('cloudinary_thumbnail')
def cloudinary_thumbnail_filter(public_id, width=200, height=200):
    """Filter untuk mendapatkan thumbnail dari Cloudinary"""
    if not public_id:
        return ''
    url, _ = cloudinary_url(public_id, width=width, height=height, crop="fill", quality="auto")
    return url

@app.template_filter('cloudinary_optimized')
def cloudinary_optimized_filter(public_id, width=None):
    """Filter untuk mendapatkan URL optimized dari Cloudinary"""
    if not public_id:
        return ''
    options = {"quality": "auto", "fetch_format": "auto"}
    if width:
        options["width"] = width
        options["crop"] = "scale"
    url, _ = cloudinary_url(public_id, **options)
    return url

@app.template_filter('filesizeformat')
def filesizeformat_filter(bytes_value):
    """Format ukuran file dalam bytes ke format yang mudah dibaca"""
    if not bytes_value:
        return '0 KB'
    
    for unit in ['B', 'KB', 'MB', 'GB']:
        if bytes_value < 1024.0:
            return f"{bytes_value:.1f} {unit}"
        bytes_value /= 1024.0
    return f"{bytes_value:.1f} TB"

# ============================================================================
# VALIDATION FUNCTIONS
# ============================================================================

def validate_image_file(file_obj, max_size_mb=5, log_warnings=False):
    """Validate image file format and size"""
    ALLOWED_EXTENSIONS = {'png', 'jpg', 'jpeg', 'gif', 'webp'}
    
    # Check if file has extension
    if '.' not in file_obj.filename:
        if log_warnings:
            flash(f'File {file_obj.filename} tidak memiliki ekstensi!', 'warning')
        return False
    
    # Extract extension
    try:
        ext = file_obj.filename.rsplit('.', 1)[1].lower()
    except IndexError:
        if log_warnings:
            flash(f'File {file_obj.filename} format tidak valid!', 'warning')
        return False
    
    # Check extension
    if ext not in ALLOWED_EXTENSIONS:
        if log_warnings:
            flash(f'File {file_obj.filename} format tidak didukung!', 'warning')
        return False
    
    # Check file size
    file_content = file_obj.read()
    max_size_bytes = max_size_mb * 1024 * 1024
    
    if len(file_content) > max_size_bytes:
        if log_warnings:
            flash(f'File {file_obj.filename} terlalu besar (maks {max_size_mb}MB)!', 'warning')
        file_obj.seek(0)  # Reset file pointer
        return False
    
    # Reset file pointer for further processing
    file_obj.seek(0)
    return True

# ============================================================================
# AUTHENTICATION ROUTES
# ============================================================================

@app.route('/')
def index():
    if current_user.is_authenticated:
        if current_user.role == 'admin_master':
            return redirect(url_for('admin_master_dashboard'))
        elif current_user.role == 'admin_operator':
            return redirect(url_for('admin_operator_dashboard'))
        elif current_user.role == 'user_operator':
            return redirect(url_for('user_operator_dashboard'))
    return redirect(url_for('login'))

@app.route('/login', methods=['GET', 'POST'])
def login():
    if current_user.is_authenticated:
        return redirect(url_for('index'))
    
    if request.method == 'POST':
        username = request.form.get('username')
        password = request.form.get('password')
        
        user = User.query.filter_by(username=username).first()
        
        if user:
            if check_password_hash(user.password, password):
                login_user(user)
                flash('Login berhasil!', 'success')
                
                # Clear session untuk mencegah caching
                session.permanent = True
                
                if user.role == 'admin_master':
                    return redirect(url_for('admin_master_dashboard'))
                elif user.role == 'admin_operator':
                    return redirect(url_for('admin_operator_dashboard'))
                elif user.role == 'user_operator':
                    return redirect(url_for('user_operator_dashboard'))
            else:
                flash('Username atau password salah!', 'error')
        else:
            flash('Username atau password salah!', 'error')
    
    response = make_response(render_template('login.html'))
    response.headers['Cache-Control'] = 'no-store, no-cache, must-revalidate, post-check=0, pre-check=0, max-age=0'
    response.headers['Pragma'] = 'no-cache'
    response.headers['Expires'] = '-1'
    return response

@app.route('/logout')
@login_required
def logout():
    logout_user()
    # Clear session
    session.clear()
    flash('Anda telah logout!', 'info')
    
    # Buat response dengan header cache-control
    response = make_response(redirect(url_for('login')))
    response.headers['Cache-Control'] = 'no-store, no-cache, must-revalidate, post-check=0, pre-check=0, max-age=0'
    response.headers['Pragma'] = 'no-cache'
    response.headers['Expires'] = '-1'
    return response

# ============================================================================
# ADMIN OPERATOR - REKAP BULANAN DAN TAHUNAN
# ============================================================================

@app.route('/admin-operator/rekap-bulanan')
@login_required
@admin_operator_required
def admin_operator_rekap_bulanan():
    """Halaman rekap bulanan untuk Admin Operator"""
    try:
        operator = current_user.operator_type
        
        # Ambil parameter filter dari URL, default ke tahun dan bulan sekarang
        tahun = request.args.get('tahun', datetime.now().year, type=int)
        bulan = request.args.get('bulan', datetime.now().month, type=int)
        kota = request.args.get('kota', 'all')
        
        # Validasi bulan (1-12)
        if bulan < 1 or bulan > 12:
            bulan = datetime.now().month
        
        # Daftar bulan dalam bahasa Indonesia
        bulan_nama = ['Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni',
                     'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember']
        
        # ===== PERBAIKAN: Ambil daftar kota case-insensitive =====
        from sqlalchemy import func
        
        kota_raw = db.session.query(
            func.lower(Stasiun.kota).label('kota_lower'),
            Stasiun.kota
        ).filter_by(operator=operator)\
         .distinct(func.lower(Stasiun.kota))\
         .order_by(func.lower(Stasiun.kota))\
         .all()
        
        # Mapping untuk menyimpan format standar
        kota_mapping = {}
        for kota_lower, kota_original in kota_raw:
            if kota_lower not in kota_mapping:
                # Cari format standar dari KOTA_LIST
                found = False
                for std_kota in KOTA_LIST:
                    if std_kota.lower() == kota_lower:
                        kota_mapping[kota_lower] = std_kota
                        found = True
                        break
                if not found:
                    # Jika tidak ditemukan di KOTA_LIST, gunakan nilai asli dengan title case
                    kota_mapping[kota_lower] = kota_original.strip().title()
        
        kota_list_operator = list(kota_mapping.values())
        
        # Jika tidak ada kota di database, gunakan KOTA_LIST default
        if not kota_list_operator:
            kota_list_operator = KOTA_LIST
            print(f"⚠️ Tidak ada kota untuk operator {operator}, menggunakan KOTA_LIST default")
        
        # Validasi kota yang dipilih ada di list (case-insensitive)
        if kota != 'all':
            kota_found = False
            for k in kota_list_operator:
                if k.lower() == kota.lower():
                    kota = k
                    kota_found = True
                    break
            if not kota_found:
                print(f"⚠️ Kota '{kota}' tidak ada di list, reset ke 'all'")
                kota = 'all'
        
        print(f"\n{'='*60}")
        print(f"📅 HALAMAN REKAP BULANAN - OPERATOR {operator.upper()}")
        print(f"   Tahun: {tahun}, Bulan: {bulan_nama[bulan-1]}, Kota/Kab: {kota}")
        print(f"   Kota list: {kota_list_operator}")
        print('='*60)
        
        return render_template('admin_operator/rekap_bulanan.html',
                             operator=operator,
                             operator_name=operator.upper(),
                             tahun=tahun,
                             bulan=bulan,
                             bulan_nama=bulan_nama[bulan-1],
                             kota=kota,
                             kota_list=kota_list_operator)
    except Exception as e:
        flash(f'Error: {str(e)}', 'error')
        print(f"❌ ERROR in admin_operator_rekap_bulanan: {str(e)}")
        import traceback
        traceback.print_exc()
        return redirect(url_for('admin_operator_dashboard'))

@app.route('/admin-operator/rekap-tahunan')
@login_required
@admin_operator_required
def admin_operator_rekap_tahunan():
    """Halaman rekap tahunan untuk Admin Operator - HANYA menampilkan form filter dan tombol export"""
    try:
        operator = current_user.operator_type
        
        # Ambil parameter filter dari URL, default ke tahun sekarang
        tahun = request.args.get('tahun', datetime.now().year, type=int)
        kota = request.args.get('kota', 'all')
        
        # ===== PERBAIKAN: Ambil daftar kota case-insensitive =====
        from sqlalchemy import func
        
        kota_raw = db.session.query(
            func.lower(Stasiun.kota).label('kota_lower'),
            Stasiun.kota
        ).filter_by(operator=operator)\
         .distinct(func.lower(Stasiun.kota))\
         .order_by(func.lower(Stasiun.kota))\
         .all()
        
        # Mapping untuk menyimpan format standar
        kota_mapping = {}
        for kota_lower, kota_original in kota_raw:
            if kota_lower not in kota_mapping:
                # Cari format standar dari KOTA_LIST
                found = False
                for std_kota in KOTA_LIST:
                    if std_kota.lower() == kota_lower:
                        kota_mapping[kota_lower] = std_kota
                        found = True
                        break
                if not found:
                    kota_mapping[kota_lower] = kota_original.strip().title()
        
        kota_list_operator = list(kota_mapping.values())
        
        # Jika tidak ada kota di database, gunakan KOTA_LIST default
        if not kota_list_operator:
            kota_list_operator = KOTA_LIST
        
        # Validasi kota yang dipilih
        if kota != 'all':
            kota_found = False
            for k in kota_list_operator:
                if k.lower() == kota.lower():
                    kota = k
                    kota_found = True
                    break
            if not kota_found:
                kota = 'all'
        
        print(f"\n{'='*60}")
        print(f"📅 HALAMAN REKAP TAHUNAN - OPERATOR {operator.upper()}")
        print(f"   Tahun: {tahun}, Kota/Kab: {kota}")
        print('='*60)
        
        return render_template('admin_operator/rekap_tahunan.html',
                             operator=operator,
                             operator_name=operator.upper(),
                             tahun=tahun,
                             kota=kota,
                             kota_list=kota_list_operator)
    except Exception as e:
        flash(f'Error: {str(e)}', 'error')
        print(f"❌ ERROR in admin_operator_rekap_tahunan: {str(e)}")
        import traceback
        traceback.print_exc()
        return redirect(url_for('admin_operator_dashboard'))

@app.route('/admin-operator/rekap-bulanan-word-gambar')
@login_required
@admin_operator_required
def admin_operator_rekap_bulanan_word_gambar():
    """Export rekap bulanan ke Word dengan gambar besar per grup - untuk Admin Operator"""
    try:
        operator = current_user.operator_type
        tahun = request.args.get('tahun', datetime.now().year, type=int)
        bulan = request.args.get('bulan', datetime.now().month, type=int)
        kota = request.args.get('kota', 'all')
        
        print(f"\n{'='*80}")
        print(f"📄 EXPORT REKAP BULANAN WORD - OPERATOR {operator.upper()}")
        print(f"   Tahun: {tahun}, Bulan: {bulan}, Kota: {kota}")
        print('='*80)
        
        # Query stasiun dengan filter case-insensitive
        query = Stasiun.query.filter_by(operator=operator)
        
        if kota != 'all' and kota:
            query = query.filter(Stasiun.kota.ilike(kota))
            print(f"✅ Filter kota dengan ILIKE: {kota}")
        
        all_stations = query.order_by(Stasiun.kota, Stasiun.stasiun_name).all()
        
        # Filter stasiun yang memiliki data di bulan ini
        stations_with_data = []
        for station in all_stations:
            has_data = db.session.query(StatusUpdate)\
                .join(StasiunLawan)\
                .filter(StasiunLawan.stasiun_id == station.id)\
                .filter(extract('year', StatusUpdate.updated_at) == tahun)\
                .filter(extract('month', StatusUpdate.updated_at) == bulan)\
                .first()
            
            if has_data:
                stations_with_data.append(station)
        
        if not stations_with_data:
            flash(f'Tidak ada data untuk periode yang dipilih', 'warning')
            return redirect(url_for('admin_operator_rekap_bulanan', tahun=tahun, bulan=bulan, kota=kota))
        
        # ===== BUAT DOKUMEN WORD =====
        from docx import Document
        from docx.shared import Inches, Pt, RGBColor, Cm
        from docx.enum.text import WD_ALIGN_PARAGRAPH
        from docx.enum.table import WD_TABLE_ALIGNMENT
        from docx.enum.section import WD_ORIENT
        from cloudinary.utils import cloudinary_url
        import requests
        from io import BytesIO
        
        doc = Document()
        
        # ===== HALAMAN 1: LANDSCAPE =====
        section = doc.sections[0]
        original_width = section.page_width
        original_height = section.page_height
        
        section.orientation = WD_ORIENT.LANDSCAPE
        section.page_width = original_height
        section.page_height = original_width
        section.top_margin = Cm(1.5)
        section.bottom_margin = Cm(1.5)
        section.left_margin = Cm(1.5)
        section.right_margin = Cm(1.5)
        
        # ===== HALAMAN COVER =====
        title = doc.add_heading('REKAP BULANAN STASIUN MICROWAVE LINK', 0)
        title.alignment = WD_ALIGN_PARAGRAPH.CENTER
        title.runs[0].font.size = Pt(24)
        title.runs[0].font.bold = True
        title.runs[0].font.color.rgb = RGBColor(0, 51, 102)
        
        doc.add_paragraph('\n' * 2)
        
        operator_heading = doc.add_heading(f'OPERATOR: {operator.upper()}', level=1)
        operator_heading.alignment = WD_ALIGN_PARAGRAPH.CENTER
        operator_heading.runs[0].font.size = Pt(20)
        operator_heading.runs[0].font.color.rgb = RGBColor(0, 102, 204)
        
        doc.add_paragraph()
        
        bulan_nama = ['Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni',
                     'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember']
        
        periode = doc.add_heading(f'{bulan_nama[bulan-1]} {tahun}', level=1)
        periode.alignment = WD_ALIGN_PARAGRAPH.CENTER
        periode.runs[0].font.size = Pt(24)
        periode.runs[0].font.color.rgb = RGBColor(0, 102, 204)
        
        doc.add_paragraph('\n' * 2)
        
        filter_text = f"Kota/Kab: {kota.title() if kota != 'all' else 'Semua Kota/kab'}"
        filter_para = doc.add_paragraph(filter_text)
        filter_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        filter_para.runs[0].font.size = Pt(14)
        filter_para.runs[0].bold = True
        
        doc.add_paragraph('\n' * 2)
        
        footer = doc.add_paragraph()
        footer.alignment = WD_ALIGN_PARAGRAPH.CENTER
        footer.add_run(f'Dicetak pada: {datetime.now().strftime("%d %B %Y")}').italic = True
        footer.add_run('\n')
        footer.add_run(f'Oleh: {current_user.username}').italic = True
        
        doc.add_page_break()
        
        # ===== RINGKASAN DATA =====
        doc.add_heading('RINGKASAN DATA STASIUN LAWAN', level=1).alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        all_rows = []
        row_counter = 0
        
        for station in stations_with_data:
            opponents = StasiunLawan.query.filter_by(stasiun_id=station.id).order_by(StasiunLawan.urutan).all()
            
            for opp in opponents:
                latest_status = StatusUpdate.query\
                    .filter_by(stasiun_lawan_id=opp.id)\
                    .filter(extract('year', StatusUpdate.updated_at) == tahun)\
                    .filter(extract('month', StatusUpdate.updated_at) == bulan)\
                    .order_by(StatusUpdate.updated_at.desc())\
                    .first()
                
                if latest_status:
                    row_counter += 1
                    all_rows.append({
                        'no': row_counter,
                        'stasiun_name': station.stasiun_name,
                        'lawan': opp.nama_stasiun_lawan,
                        'freq_tx': opp.freq_tx if opp.freq_tx else '-',
                        'freq_rx': opp.freq_rx if opp.freq_rx else '-',
                        'kota': station.kota
                    })
        
        if all_rows:
            rows_per_page = 35
            total_rows = len(all_rows)
            total_pages = (total_rows + rows_per_page - 1) // rows_per_page
            
            for page in range(total_pages):
                if page > 0:
                    doc.add_page_break()
                
                start_idx = page * rows_per_page
                end_idx = min(start_idx + rows_per_page, total_rows)
                page_data = all_rows[start_idx:end_idx]
                
                if total_pages > 1:
                    page_info = doc.add_paragraph()
                    page_info.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                    page_info.add_run(f'Halaman {page+1} dari {total_pages} (Ringkasan)').italic = True
                    page_info.runs[0].font.size = Pt(9)
                
                table = doc.add_table(rows=1, cols=6)
                table.style = 'Table Grid'
                table.autofit = False
                
                table.columns[0].width = Cm(1.0)
                table.columns[1].width = Cm(3.5)
                table.columns[2].width = Cm(4.0)
                table.columns[3].width = Cm(2.5)
                table.columns[4].width = Cm(2.5)
                table.columns[5].width = Cm(3.0)
                
                header_cells = table.rows[0].cells
                headers = ['No', 'Stasiun', 'Stasiun Lawan', 'Freq Tx', 'Freq Rx', 'Kota']
                for i, header in enumerate(headers):
                    header_cells[i].text = header
                    header_cells[i].paragraphs[0].runs[0].font.bold = True
                    header_cells[i].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                
                for row in page_data:
                    row_cells = table.add_row().cells
                    row_cells[0].text = str(row['no'])
                    row_cells[0].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                    row_cells[1].text = row['stasiun_name']
                    row_cells[1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.LEFT
                    row_cells[2].text = row['lawan']
                    row_cells[2].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.LEFT
                    row_cells[3].text = row['freq_tx']
                    row_cells[3].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                    row_cells[4].text = row['freq_rx']
                    row_cells[4].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                    row_cells[5].text = row['kota']
                    row_cells[5].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.LEFT
        
        doc.add_page_break()
        
        # ===== HALAMAN DOKUMENTASI GAMBAR (PORTRAIT) =====
        new_section = doc.add_section()
        new_section.orientation = WD_ORIENT.PORTRAIT
        new_section.page_width = original_width
        new_section.page_height = original_height
        new_section.top_margin = Cm(2)
        new_section.bottom_margin = Cm(2)
        new_section.left_margin = Cm(2.5)
        new_section.right_margin = Cm(2.5)
        
        doc.add_heading('DOKUMENTASI GAMBAR PER GRUP', level=1).alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        for station in stations_with_data:
            doc.add_heading(f'STASIUN: {station.stasiun_name}', level=2)
            
            info_para = doc.add_paragraph()
            info_para.add_run(f'Operator: {station.operator.upper()}').bold = True
            info_para.add_run(f' | Kota/Kab: {station.kota}').bold = True
            
            groups = GrupStasiun.query.filter_by(stasiun_id=station.id).all()
            groups_dict = {g.id: g.nama_grup for g in groups}
            
            all_opponents = StasiunLawan.query.filter_by(stasiun_id=station.id).order_by(StasiunLawan.urutan).all()
            
            opponents_by_group = {}
            for opp in all_opponents:
                group_key = opp.group_id if opp.group_id is not None else 'ungrouped'
                if group_key not in opponents_by_group:
                    opponents_by_group[group_key] = []
                
                latest_status = StatusUpdate.query\
                    .filter_by(stasiun_lawan_id=opp.id)\
                    .filter(extract('year', StatusUpdate.updated_at) == tahun)\
                    .filter(extract('month', StatusUpdate.updated_at) == bulan)\
                    .order_by(StatusUpdate.updated_at.desc())\
                    .first()
                
                opponents_by_group[group_key].append({
                    'nama': opp.nama_stasiun_lawan,
                    'freq_tx': opp.freq_tx if opp.freq_tx else '-',
                    'freq_rx': opp.freq_rx if opp.freq_rx else '-',
                    'status': latest_status.status if latest_status else 'belum_ada',
                    'status_display': STATUS_OPTIONS.get(latest_status.status, 'Belum Ada') if latest_status else 'Belum Ada',
                    'catatan': latest_status.catatan if latest_status and latest_status.catatan else '-'
                })
            
            # ===== AMBIL SEMUA UPLOAD (TANPA LIMIT) =====
            uploads_in_bulan = UploadGambar.query\
                .filter_by(stasiun_id=station.id)\
                .filter(extract('year', UploadGambar.uploaded_at) == tahun)\
                .filter(extract('month', UploadGambar.uploaded_at) == bulan)\
                .order_by(UploadGambar.uploaded_at.desc())\
                .all()
            
            uploads_by_group = {}
            seen_uploads = set()
            
            for upload in uploads_in_bulan:
                group_key = upload.group_id if upload.group_id is not None else 'ungrouped'
                upload_key = f"{group_key}_{upload.public_id}"
                
                if upload_key not in seen_uploads:
                    seen_uploads.add(upload_key)
                    
                    if group_key not in uploads_by_group:
                        uploads_by_group[group_key] = []
                    
                    uploads_by_group[group_key].append(upload)
            
            all_group_keys = set(list(opponents_by_group.keys()) + list(uploads_by_group.keys()))
            
            sorted_group_keys = []
            if 'ungrouped' in all_group_keys:
                sorted_group_keys.append('ungrouped')
            
            numeric_keys = [k for k in all_group_keys if k != 'ungrouped' and k is not None]
            numeric_keys.sort(key=lambda x: int(x) if str(x).isdigit() else x)
            sorted_group_keys.extend(numeric_keys)
            
            group_counter = 0
            for group_key in sorted_group_keys:
                group_counter += 1
                
                if group_key == 'ungrouped':
                    group_display = 'Tanpa Grup'
                else:
                    group_display = groups_dict.get(int(group_key), f'Grup {group_key}')
                
                doc.add_heading(f'  Grup {group_counter}: {group_display}', level=3)
                
                group_uploads = uploads_by_group.get(group_key, [])
                
                if group_uploads:
                    doc.add_heading('    GAMBAR:', level=4)
                    
                    for upload in group_uploads:
                        try:
                            # ===== PERUBAHAN: Gunakan thumbnail =====
                            thumbnail_url, _ = cloudinary_url(
                                upload.public_id,
                                width=500,
                                height=375,
                                crop="limit",
                                quality="auto:eco"
                            )
                            response = requests.get(thumbnail_url, timeout=30)
                            if response.status_code == 200:
                                img_data = BytesIO(response.content)
                                
                                paragraph = doc.add_paragraph()
                                paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                                run = paragraph.add_run()
                                run.add_picture(img_data, width=Inches(5.0))
                                
                                info_para = doc.add_paragraph()
                                info_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                                info_text = f"File: {upload.original_filename} | "
                                info_text += f"Upload: {upload.uploaded_at.strftime('%d/%m/%Y')}"
                                info_run = info_para.add_run(info_text)
                                info_run.italic = True
                                info_run.font.size = Pt(9)
                                
                                doc.add_paragraph()
                        except Exception as e:
                            print(f"Error download: {e}")
                            doc.add_paragraph(f"    Gambar tidak tersedia: {upload.original_filename}")
                else:
                    doc.add_paragraph("    Tidak ada gambar untuk grup ini")
                
                group_opponents = opponents_by_group.get(group_key, [])
                if group_opponents:
                    doc.add_heading('    DAFTAR STASIUN LAWAN:', level=4)
                    
                    opp_table = doc.add_table(rows=1, cols=6)
                    opp_table.style = 'Table Grid'
                    opp_table.autofit = True
                    
                    opp_headers = ['No', 'Nama Stasiun Lawan', 'Freq Tx', 'Freq Rx', 'Status', 'Catatan']
                    for i, header in enumerate(opp_headers):
                        opp_table.rows[0].cells[i].text = header
                        opp_table.rows[0].cells[i].paragraphs[0].runs[0].font.bold = True
                        opp_table.rows[0].cells[i].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                    
                    for opp_idx, opponent in enumerate(group_opponents, 1):
                        row_cells = opp_table.add_row().cells
                        row_cells[0].text = str(opp_idx)
                        row_cells[0].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                        row_cells[1].text = opponent['nama']
                        row_cells[1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.LEFT
                        row_cells[2].text = opponent['freq_tx']
                        row_cells[2].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                        row_cells[3].text = opponent['freq_rx']
                        row_cells[3].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                        row_cells[4].text = opponent['status_display']
                        row_cells[4].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                        row_cells[5].text = opponent['catatan']
                        row_cells[5].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.LEFT
                    
                    doc.add_paragraph()
                else:
                    doc.add_paragraph("    Tidak ada stasiun lawan dalam grup ini")
                
                doc.add_paragraph()
            
            doc.add_paragraph()
        
        file_stream = BytesIO()
        doc.save(file_stream)
        file_stream.seek(0)
        
        filename = f"Rekap_Bulanan_{operator.upper()}_{bulan_nama[bulan-1]}_{tahun}"
        if kota != 'all':
            filename += f"_{kota.replace(' ', '_')}"
        filename += ".docx"
        
        return send_file(
            file_stream,
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            as_attachment=True,
            download_name=filename
        )
        
    except Exception as e:
        flash(f'Error: {str(e)}', 'error')
        print(f"\n❌ ERROR: {str(e)}")
        import traceback
        traceback.print_exc()
        return redirect(url_for('admin_operator_rekap_bulanan', tahun=tahun, bulan=bulan, kota=kota))

@app.route('/admin-operator/rekap-tahunan-word-gambar')
@login_required
@admin_operator_required
def admin_operator_rekap_tahunan_word_gambar():
    """Export rekap tahunan ke Word dengan gambar per grup - untuk Admin Operator"""
    try:
        operator = current_user.operator_type
        tahun = request.args.get('tahun', datetime.now().year, type=int)
        kota = request.args.get('kota', 'all')
        
        print(f"\n{'='*80}")
        print(f"📄 EXPORT REKAP TAHUNAN WORD - OPERATOR {operator.upper()}")
        print(f"   Tahun: {tahun}, Kota: {kota}")
        print('='*80)
        
        # Query stasiun dengan filter case-insensitive
        query = Stasiun.query.filter_by(operator=operator)
        
        if kota != 'all' and kota:
            query = query.filter(Stasiun.kota.ilike(kota))
            print(f"✅ Filter kota dengan ILIKE: {kota}")
        
        all_stations = query.order_by(Stasiun.kota, Stasiun.stasiun_name).all()
        
        # Filter stasiun yang memiliki data di tahun tersebut
        stations_with_data = []
        for station in all_stations:
            has_data = db.session.query(StatusUpdate)\
                .join(StasiunLawan)\
                .filter(StasiunLawan.stasiun_id == station.id)\
                .filter(extract('year', StatusUpdate.updated_at) == tahun)\
                .first()
            
            if has_data:
                stations_with_data.append(station)
        
        if not stations_with_data:
            flash(f'Tidak ada data untuk tahun {tahun}', 'warning')
            return redirect(url_for('admin_operator_rekap_tahunan', tahun=tahun, kota=kota))
        
        # ===== KUMPULKAN DATA PER BULAN =====
        bulan_nama = ['Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni',
                     'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember']
        
        bulan_data = {}
        bulan_dengan_data = []
        
        for bulan in range(1, 13):
            bulan_key = f"{tahun}-{bulan:02d}"
            bulan_data[bulan_key] = {
                'nama': bulan_nama[bulan-1],
                'stations': [],
                'total_uploads': 0,
                'has_data': False
            }
            
            for station in stations_with_data:
                status_count = db.session.query(StatusUpdate)\
                    .join(StasiunLawan)\
                    .filter(StasiunLawan.stasiun_id == station.id)\
                    .filter(extract('year', StatusUpdate.updated_at) == tahun)\
                    .filter(extract('month', StatusUpdate.updated_at) == bulan)\
                    .count()
                
                upload_count = UploadGambar.query\
                    .filter_by(stasiun_id=station.id)\
                    .filter(extract('year', UploadGambar.uploaded_at) == tahun)\
                    .filter(extract('month', UploadGambar.uploaded_at) == bulan)\
                    .count()
                
                if status_count > 0 or upload_count > 0:
                    bulan_data[bulan_key]['has_data'] = True
                    bulan_data[bulan_key]['stations'].append(station)
                    bulan_data[bulan_key]['total_uploads'] += upload_count
            
            if bulan_data[bulan_key]['has_data']:
                bulan_dengan_data.append(bulan)
                print(f"  ✅ {bulan_nama[bulan-1]}: {len(bulan_data[bulan_key]['stations'])} stasiun, {bulan_data[bulan_key]['total_uploads']} uploads")
        
        # ===== BUAT DOKUMEN WORD =====
        from docx import Document
        from docx.shared import Inches, Pt, RGBColor, Cm
        from docx.enum.text import WD_ALIGN_PARAGRAPH
        from docx.enum.table import WD_TABLE_ALIGNMENT
        from docx.enum.section import WD_ORIENT
        from cloudinary.utils import cloudinary_url
        import requests
        from io import BytesIO
        
        doc = Document()
        
        section = doc.sections[0]
        original_width = section.page_width
        original_height = section.page_height
        
        section.orientation = WD_ORIENT.LANDSCAPE
        section.page_width = original_height
        section.page_height = original_width
        section.top_margin = Cm(1.5)
        section.bottom_margin = Cm(1.5)
        section.left_margin = Cm(1.5)
        section.right_margin = Cm(1.5)
        
        title = doc.add_heading('REKAP TAHUNAN STASIUN MICROWAVE LINK', 0)
        title.alignment = WD_ALIGN_PARAGRAPH.CENTER
        title.runs[0].font.size = Pt(24)
        title.runs[0].font.bold = True
        title.runs[0].font.color.rgb = RGBColor(0, 51, 102)
        
        doc.add_paragraph('\n' * 4)
        
        operator_heading = doc.add_heading(f'OPERATOR: {operator.upper()}', level=1)
        operator_heading.alignment = WD_ALIGN_PARAGRAPH.CENTER
        operator_heading.runs[0].font.size = Pt(20)
        operator_heading.runs[0].font.color.rgb = RGBColor(0, 102, 204)
        
        doc.add_paragraph()
        
        tahun_heading = doc.add_heading(f'TAHUN {tahun}', level=1)
        tahun_heading.alignment = WD_ALIGN_PARAGRAPH.CENTER
        tahun_heading.runs[0].font.size = Pt(28)
        tahun_heading.runs[0].font.color.rgb = RGBColor(0, 102, 204)
        
        doc.add_paragraph('\n' * 4)
        
        filter_text = f"Kota/Kab: {kota.title() if kota != 'all' else 'Semua Kota/kab'}"
        filter_para = doc.add_paragraph(filter_text)
        filter_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        filter_para.runs[0].font.size = Pt(14)
        filter_para.runs[0].bold = True
        
        doc.add_paragraph('\n' * 2)
        
        footer = doc.add_paragraph()
        footer.alignment = WD_ALIGN_PARAGRAPH.CENTER
        footer.add_run(f'Dicetak pada: {datetime.now().strftime("%d %B %Y")}').italic = True
        footer.add_run('\n')
        footer.add_run(f'Oleh: {current_user.username}').italic = True
        
        doc.add_page_break()
        
        doc.add_heading('RINGKASAN STATISTIK TAHUNAN', level=1).alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        all_rows_data = []
        row_counter = 0
        
        for station in stations_with_data:
            opponents = StasiunLawan.query.filter_by(stasiun_id=station.id).order_by(StasiunLawan.urutan).all()
            
            for opp in opponents:
                row_counter += 1
                all_rows_data.append({
                    'no': row_counter,
                    'stasiun_name': station.stasiun_name,
                    'stasiun_lawan': opp.nama_stasiun_lawan,
                    'freq_tx': opp.freq_tx if opp.freq_tx else '-',
                    'freq_rx': opp.freq_rx if opp.freq_rx else '-',
                    'kota': station.kota
                })
        
        if all_rows_data:
            rows_per_page = 35
            total_rows = len(all_rows_data)
            total_pages = (total_rows + rows_per_page - 1) // rows_per_page
            
            for page in range(total_pages):
                if page > 0:
                    doc.add_page_break()
                
                start_idx = page * rows_per_page
                end_idx = min(start_idx + rows_per_page, total_rows)
                page_data = all_rows_data[start_idx:end_idx]
                
                if total_pages > 1:
                    page_info = doc.add_paragraph()
                    page_info.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                    page_info.add_run(f'Halaman {page+1} dari {total_pages} (Ringkasan)').italic = True
                    page_info.runs[0].font.size = Pt(9)
                
                main_table = doc.add_table(rows=1, cols=6)
                main_table.style = 'Table Grid'
                main_table.autofit = False
                
                main_table.columns[0].width = Cm(1.0)
                main_table.columns[1].width = Cm(3.5)
                main_table.columns[2].width = Cm(4.0)
                main_table.columns[3].width = Cm(2.5)
                main_table.columns[4].width = Cm(2.5)
                main_table.columns[5].width = Cm(3.0)
                
                header_cells = main_table.rows[0].cells
                headers = ['No', 'Stasiun Name', 'Stasiun Lawan', 'Freq Tx', 'Freq Rx', 'Kota/Kab']
                for i, header in enumerate(headers):
                    header_cells[i].text = header
                    header_cells[i].paragraphs[0].runs[0].font.bold = True
                    header_cells[i].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                
                for data in page_data:
                    row_cells = main_table.add_row().cells
                    row_cells[0].text = str(data['no'])
                    row_cells[0].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                    row_cells[1].text = data['stasiun_name']
                    row_cells[1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.LEFT
                    row_cells[2].text = data['stasiun_lawan']
                    row_cells[2].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.LEFT
                    row_cells[3].text = data['freq_tx']
                    row_cells[3].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                    row_cells[4].text = data['freq_rx']
                    row_cells[4].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                    row_cells[5].text = data['kota']
                    row_cells[5].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.LEFT
        
        doc.add_page_break()
        
        new_section = doc.add_section()
        new_section.orientation = WD_ORIENT.PORTRAIT
        new_section.page_width = original_width
        new_section.page_height = original_height
        new_section.top_margin = Cm(2)
        new_section.bottom_margin = Cm(2)
        new_section.left_margin = Cm(2.5)
        new_section.right_margin = Cm(2.5)
        
        doc.add_heading('DOKUMENTASI GAMBAR PER BULAN', level=1).alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        if not bulan_dengan_data:
            doc.add_paragraph(f"Tidak ada data untuk tahun {tahun}")
        else:
            for bulan in sorted(bulan_dengan_data):
                bulan_key = f"{tahun}-{bulan:02d}"
                bulan_info = bulan_data[bulan_key]
                
                doc.add_heading(f'{bulan_info["nama"].upper()} {tahun}', level=2)
                doc.add_paragraph()
                
                for station in bulan_info['stations']:
                    station_heading = doc.add_heading(f'STASIUN: {station.stasiun_name}', level=3)
                    station_heading.runs[0].font.color.rgb = RGBColor(67, 97, 238)
                    
                    info_para = doc.add_paragraph()
                    info_para.add_run(f'Operator: {station.operator.upper()}').bold = True
                    info_para.add_run(f' | Kota/Kab: {station.kota}').bold = True
                    
                    groups = GrupStasiun.query.filter_by(stasiun_id=station.id).all()
                    groups_dict = {g.id: g.nama_grup for g in groups}
                    
                    all_opponents = StasiunLawan.query.filter_by(stasiun_id=station.id).order_by(StasiunLawan.urutan).all()
                    
                    opponents_by_group = {}
                    for opp in all_opponents:
                        group_key = opp.group_id if opp.group_id is not None else 'ungrouped'
                        if group_key not in opponents_by_group:
                            opponents_by_group[group_key] = []
                        
                        latest_status = StatusUpdate.query\
                            .filter_by(stasiun_lawan_id=opp.id)\
                            .filter(extract('year', StatusUpdate.updated_at) == tahun)\
                            .filter(extract('month', StatusUpdate.updated_at) == bulan)\
                            .order_by(StatusUpdate.updated_at.desc())\
                            .first()
                        
                        opponents_by_group[group_key].append({
                            'nama': opp.nama_stasiun_lawan,
                            'freq_tx': opp.freq_tx if opp.freq_tx else '-',
                            'freq_rx': opp.freq_rx if opp.freq_rx else '-',
                            'status': latest_status.status if latest_status else 'belum_ada',
                            'status_display': STATUS_OPTIONS.get(latest_status.status, 'Belum Ada') if latest_status else 'Belum Ada',
                            'catatan': latest_status.catatan if latest_status and latest_status.catatan else '-'
                        })
                    
                    # ===== AMBIL SEMUA UPLOAD (TANPA LIMIT) =====
                    uploads_in_bulan = UploadGambar.query\
                        .filter_by(stasiun_id=station.id)\
                        .filter(extract('year', UploadGambar.uploaded_at) == tahun)\
                        .filter(extract('month', UploadGambar.uploaded_at) == bulan)\
                        .order_by(UploadGambar.uploaded_at.desc())\
                        .all()
                    
                    uploads_by_group = {}
                    seen_uploads = set()
                    
                    for upload in uploads_in_bulan:
                        group_key = upload.group_id if upload.group_id is not None else 'ungrouped'
                        upload_key = f"{group_key}_{upload.public_id}"
                        
                        if upload_key not in seen_uploads:
                            seen_uploads.add(upload_key)
                            
                            if group_key not in uploads_by_group:
                                uploads_by_group[group_key] = []
                            
                            uploads_by_group[group_key].append(upload)
                    
                    all_group_keys = set(list(opponents_by_group.keys()) + list(uploads_by_group.keys()))
                    
                    sorted_group_keys = []
                    if 'ungrouped' in all_group_keys:
                        sorted_group_keys.append('ungrouped')
                    
                    numeric_keys = [k for k in all_group_keys if k != 'ungrouped' and k is not None]
                    numeric_keys.sort(key=lambda x: int(x) if str(x).isdigit() else x)
                    sorted_group_keys.extend(numeric_keys)
                    
                    group_counter = 0
                    for group_key in sorted_group_keys:
                        group_counter += 1
                        
                        if group_key == 'ungrouped':
                            group_display = 'Tanpa Grup'
                        else:
                            group_display = groups_dict.get(int(group_key), f'Grup {group_key}')
                        
                        doc.add_heading(f'  Grup {group_counter}: {group_display}', level=4)
                        
                        group_uploads = uploads_by_group.get(group_key, [])
                        
                        if group_uploads:
                            doc.add_heading('    GAMBAR:', level=5)
                            
                            for upload in group_uploads:
                                try:
                                    # ===== PERUBAHAN: Gunakan thumbnail =====
                                    thumbnail_url, _ = cloudinary_url(
                                        upload.public_id,
                                        width=500,
                                        height=375,
                                        crop="limit",
                                        quality="auto:eco"
                                    )
                                    response = requests.get(thumbnail_url, timeout=30)
                                    if response.status_code == 200:
                                        img_data = BytesIO(response.content)
                                        
                                        paragraph = doc.add_paragraph()
                                        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                                        run = paragraph.add_run()
                                        run.add_picture(img_data, width=Inches(5.0))
                                        
                                        info_para = doc.add_paragraph()
                                        info_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                                        info_text = f"File: {upload.original_filename} | "
                                        info_text += f"Upload: {upload.uploaded_at.strftime('%d/%m/%Y')}"
                                        info_run = info_para.add_run(info_text)
                                        info_run.italic = True
                                        info_run.font.size = Pt(9)
                                        
                                        doc.add_paragraph()
                                except Exception as e:
                                    print(f"Error download: {e}")
                                    doc.add_paragraph(f"    Gambar tidak tersedia: {upload.original_filename}")
                        else:
                            doc.add_paragraph("    Tidak ada gambar untuk grup ini")
                        
                        group_opponents = opponents_by_group.get(group_key, [])
                        if group_opponents:
                            doc.add_heading('    DAFTAR STASIUN LAWAN:', level=5)
                            
                            opp_table = doc.add_table(rows=1, cols=6)
                            opp_table.style = 'Table Grid'
                            opp_table.autofit = True
                            
                            opp_headers = ['No', 'Nama Stasiun Lawan', 'Freq Tx', 'Freq Rx', 'Status', 'Catatan']
                            for i, header in enumerate(opp_headers):
                                opp_table.rows[0].cells[i].text = header
                                opp_table.rows[0].cells[i].paragraphs[0].runs[0].font.bold = True
                                opp_table.rows[0].cells[i].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                            
                            for opp_idx, opponent in enumerate(group_opponents, 1):
                                row_cells = opp_table.add_row().cells
                                row_cells[0].text = str(opp_idx)
                                row_cells[0].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                                row_cells[1].text = opponent['nama']
                                row_cells[1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.LEFT
                                row_cells[2].text = opponent['freq_tx']
                                row_cells[2].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                                row_cells[3].text = opponent['freq_rx']
                                row_cells[3].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                                row_cells[4].text = opponent['status_display']
                                row_cells[4].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                                row_cells[5].text = opponent['catatan']
                                row_cells[5].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.LEFT
                            
                            doc.add_paragraph()
                        else:
                            doc.add_paragraph("    Tidak ada stasiun lawan dalam grup ini")
                        
                        doc.add_paragraph()
                    
                    doc.add_paragraph()
                
                if bulan != bulan_dengan_data[-1]:
                    doc.add_page_break()
        
        file_stream = BytesIO()
        doc.save(file_stream)
        file_stream.seek(0)
        
        filename = f"Rekap_Tahunan_{operator.upper()}_{tahun}"
        if kota != 'all':
            filename += f"_{kota.replace(' ', '_')}"
        filename += ".docx"
        
        return send_file(
            file_stream,
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            as_attachment=True,
            download_name=filename
        )
        
    except Exception as e:
        flash(f'Error: {str(e)}', 'error')
        print(f"\n❌ ERROR: {str(e)}")
        import traceback
        traceback.print_exc()
        return redirect(url_for('admin_operator_rekap_tahunan', tahun=tahun, kota=kota))
    
# ============================================================================
# ADMIN OPERATOR - HASIL LAPORAN 
# ============================================================================

@app.route('/admin-operator/hasil-laporan', methods=['GET', 'POST'])
@login_required
@admin_operator_required
def admin_operator_hasil_laporan():
    """Halaman input dan daftar hasil laporan untuk Admin Operator"""
    operator = current_user.operator_type
    
    if request.method == 'POST':
        try:
            # Operator otomatis dari user yang login
            nomor_ba = request.form.get('nomor_ba', '').strip()
            tanggal_ba_str = request.form.get('tanggal_ba')
            nomor_spt = request.form.get('nomor_spt', '').strip()
            kota_kab = request.form.get('kota_kab')
            
            # Validasi semua field wajib
            if not all([nomor_ba, tanggal_ba_str, nomor_spt, kota_kab]):
                flash('Semua field wajib harus diisi!', 'error')
                return redirect(url_for('admin_operator_hasil_laporan'))
            
            # Cek duplikasi nomor BA
            existing = HasilLaporan.query.filter_by(nomor_ba=nomor_ba).first()
            if existing:
                flash(f'Nomor BA "{nomor_ba}" sudah ada!', 'error')
                return redirect(url_for('admin_operator_hasil_laporan'))
            
            # Parse tanggal
            tanggal_ba = datetime.strptime(tanggal_ba_str, '%Y-%m-%d').date()
            
            # ===== PERBAIKAN DI SINI =====
            # Proses upload PDF
            file_pdf_url = None
            file_pdf_path = None  # ← GANTI: dari public_id ke path
            
            if 'file_pdf' in request.files:
                pdf_file = request.files['file_pdf']
                if pdf_file and pdf_file.filename:
                    if not pdf_file.filename.lower().endswith('.pdf'):
                        flash('File harus berformat PDF!', 'error')
                        return redirect(url_for('admin_operator_hasil_laporan'))
                    
                    result, error = upload_pdf_to_local(pdf_file, operator, nomor_ba)  # ← GANTI: dari cloudinary ke local
                    if result:
                        file_pdf_url = result['url']
                        file_pdf_path = result['file_path']  # ← GANTI: dari public_id ke path
                    else:
                        flash(f'Gagal upload PDF: {error}', 'error')
                        return redirect(url_for('admin_operator_hasil_laporan'))
            
            # Simpan ke database
            new_laporan = HasilLaporan(
                operator=operator,
                nomor_ba=nomor_ba,
                tanggal_ba=tanggal_ba,
                nomor_spt=nomor_spt,
                kota_kab=kota_kab,
                file_pdf_url=file_pdf_url,
                file_pdf_path=file_pdf_path,  # ← GANTI: dari public_id ke path
                created_by=current_user.id
            )
            
            db.session.add(new_laporan)
            db.session.commit()
            
            flash(f'✅ Laporan dengan Nomor BA "{nomor_ba}" berhasil ditambahkan!', 'success')
            return redirect(url_for('admin_operator_hasil_laporan'))
            
        except Exception as e:
            db.session.rollback()
            flash(f'❌ Error: {str(e)}', 'error')
            print(f"Error in admin_operator_hasil_laporan POST: {str(e)}")
            import traceback
            traceback.print_exc()
            return redirect(url_for('admin_operator_hasil_laporan'))
    
    # ========== GET REQUEST - TAMPILKAN DAFTAR LAPORAN ==========
    page = request.args.get('page', 1, type=int)
    per_page = request.args.get('per_page', 10, type=int)
    
    # Ambil semua parameter filter (tanpa filter operator karena sudah tetap)
    filter_kota = request.args.get('filter_kota', 'all')
    search_nomor_ba = request.args.get('search_nomor_ba', '').strip()
    search_nomor_spt = request.args.get('search_nomor_spt', '').strip()
    tanggal_mulai = request.args.get('tanggal_mulai', '').strip()
    tanggal_sampai = request.args.get('tanggal_sampai', '').strip()
    
    # Query hanya untuk operator yang sama
    query = HasilLaporan.query.filter(HasilLaporan.operator == operator)
    
    # Filter berdasarkan kota/kab
    if filter_kota != 'all':
        query = query.filter(HasilLaporan.kota_kab == filter_kota)
    
    # Filter pencarian Nomor BA
    if search_nomor_ba:
        query = query.filter(HasilLaporan.nomor_ba.ilike(f'%{search_nomor_ba}%'))
    
    # Filter pencarian Nomor SPT
    if search_nomor_spt:
        query = query.filter(HasilLaporan.nomor_spt.ilike(f'%{search_nomor_spt}%'))
    
    # Filter rentang tanggal mulai
    if tanggal_mulai:
        try:
            tanggal_mulai_date = datetime.strptime(tanggal_mulai, '%Y-%m-%d').date()
            query = query.filter(HasilLaporan.tanggal_ba >= tanggal_mulai_date)
        except ValueError:
            pass
    
    # Filter rentang tanggal sampai
    if tanggal_sampai:
        try:
            tanggal_sampai_date = datetime.strptime(tanggal_sampai, '%Y-%m-%d').date()
            query = query.filter(HasilLaporan.tanggal_ba <= tanggal_sampai_date)
        except ValueError:
            pass
    
    # Urutkan berdasarkan tanggal dibuat terbaru
    query = query.order_by(HasilLaporan.created_at.desc())
    
    # Pagination
    pagination = query.paginate(page=page, per_page=per_page, error_out=False)
    laporan_list = pagination.items
    
    # Ambil daftar kota yang pernah ada dalam laporan untuk operator ini
    kota_with_laporan = db.session.query(HasilLaporan.kota_kab)\
        .filter(HasilLaporan.operator == operator)\
        .distinct().all()
    kota_with_laporan = [k[0] for k in kota_with_laporan if k[0]]
    
    return render_template('admin_operator/hasil_laporan.html',
                         laporan_list=laporan_list,
                         pagination=pagination,
                         operator=operator,
                         operator_name=operator.upper(),
                         kota_list=KOTA_LIST,
                         kota_with_laporan=kota_with_laporan,
                         filter_kota=filter_kota)

@app.route('/admin-operator/hasil-laporan/edit/<int:laporan_id>', methods=['GET', 'POST'])
@login_required
@admin_operator_required
def admin_operator_edit_laporan(laporan_id):
    """Edit hasil laporan untuk Admin Operator"""
    laporan = HasilLaporan.query.get_or_404(laporan_id)
    
    # Validasi kepemilikan
    if laporan.operator != current_user.operator_type:
        flash('Akses ditolak! Laporan ini bukan milik operator Anda.', 'error')
        return redirect(url_for('admin_operator_hasil_laporan'))
    
    if request.method == 'POST':
        try:
            nomor_ba = request.form.get('nomor_ba', '').strip()
            tanggal_ba_str = request.form.get('tanggal_ba')
            nomor_spt = request.form.get('nomor_spt', '').strip()
            kota_kab = request.form.get('kota_kab')
            
            if not all([nomor_ba, tanggal_ba_str, nomor_spt, kota_kab]):
                flash('Semua field wajib harus diisi!', 'error')
                return redirect(url_for('admin_operator_edit_laporan', laporan_id=laporan_id))
            
            # Cek duplikasi nomor BA (kecuali untuk laporan ini sendiri)
            existing = HasilLaporan.query.filter(
                HasilLaporan.nomor_ba == nomor_ba,
                HasilLaporan.id != laporan_id
            ).first()
            
            if existing:
                flash(f'Nomor BA "{nomor_ba}" sudah digunakan!', 'error')
                return redirect(url_for('admin_operator_edit_laporan', laporan_id=laporan_id))
            
            tanggal_ba = datetime.strptime(tanggal_ba_str, '%Y-%m-%d').date()
            
            # Proses upload PDF baru jika ada
            file_pdf_url = laporan.file_pdf_url
            file_pdf_path = laporan.file_pdf_path
            
            if 'file_pdf' in request.files:
                pdf_file = request.files['file_pdf']
                if pdf_file and pdf_file.filename:
                    if not pdf_file.filename.lower().endswith('.pdf'):
                        flash('File harus berformat PDF!', 'error')
                        return redirect(url_for('admin_operator_edit_laporan', laporan_id=laporan_id))
                    
                    # Hapus file lama jika ada
                    if laporan.file_pdf_path:
                        delete_pdf_from_local(laporan.file_pdf_path)
                    
                    result, error = upload_pdf_to_local(pdf_file, laporan.operator, nomor_ba)
                    if result:
                        file_pdf_url = result['url']
                        file_pdf_path = result['file_path']
                    else:
                        flash(f'Gagal upload PDF: {error}', 'error')
                        return redirect(url_for('admin_operator_edit_laporan', laporan_id=laporan_id))
            
            # Update data
            laporan.nomor_ba = nomor_ba
            laporan.tanggal_ba = tanggal_ba
            laporan.nomor_spt = nomor_spt
            laporan.kota_kab = kota_kab
            laporan.file_pdf_url = file_pdf_url
            laporan.file_pdf_path = file_pdf_path
            laporan.updated_at = datetime.utcnow()
            
            db.session.commit()
            flash(f'✅ Laporan berhasil diperbarui!', 'success')
            return redirect(url_for('admin_operator_hasil_laporan'))
            
        except Exception as e:
            db.session.rollback()
            flash(f'❌ Error: {str(e)}', 'error')
            return redirect(url_for('admin_operator_edit_laporan', laporan_id=laporan_id))
    
    return render_template('admin_operator/edit_laporan.html',
                         laporan=laporan,
                         kota_list=KOTA_LIST)


@app.route('/admin-operator/hasil-laporan/hapus/<int:laporan_id>')
@login_required
@admin_operator_required
def admin_operator_hapus_laporan(laporan_id):
    """Hapus hasil laporan untuk Admin Operator"""
    laporan = HasilLaporan.query.get_or_404(laporan_id)
    
    # Validasi kepemilikan
    if laporan.operator != current_user.operator_type:
        flash('Akses ditolak! Laporan ini bukan milik operator Anda.', 'error')
        return redirect(url_for('admin_operator_hasil_laporan'))
    
    nomor_ba = laporan.nomor_ba
    
    try:
        if laporan.file_pdf_path:
            delete_pdf_from_local(laporan.file_pdf_path)
        
        db.session.delete(laporan)
        db.session.commit()
        flash(f'✅ Laporan dengan Nomor BA "{nomor_ba}" berhasil dihapus!', 'success')
    except Exception as e:
        db.session.rollback()
        flash(f'❌ Error: {str(e)}', 'error')
    
    return redirect(url_for('admin_operator_hasil_laporan'))

# ============================================================================
# ADMIN MASTER ROUTES
# ============================================================================
@app.route('/admin-master/dashboard')
@login_required
@admin_master_required
def admin_master_dashboard():
    total_stations = Stasiun.query.count()
    total_opponents = StasiunLawan.query.count()
    
    # ===== STATISTIK UPLOAD =====
    stations_with_uploads = Stasiun.query.filter(Stasiun.uploads.any()).count()
    stations_without_uploads = total_stations - stations_with_uploads
    
    # ===== STATISTIK PEMERIKSAAN =====
    stations_checked = Stasiun.query.filter(Stasiun.status_pemeriksaan == 'sudah_diperiksa').count()
    stations_unchecked = Stasiun.query.filter(Stasiun.status_pemeriksaan == 'belum_diperiksa').count()
    
    # ===== STATISTIK PER OPERATOR =====
    # Telkom
    telkom_stations = Stasiun.query.filter_by(operator='telkom').count()
    telkom_uploaded = Stasiun.query.filter(Stasiun.operator == 'telkom', Stasiun.uploads.any()).count()
    telkom_not_uploaded = telkom_stations - telkom_uploaded
    telkom_checked = Stasiun.query.filter(Stasiun.operator == 'telkom', Stasiun.status_pemeriksaan == 'sudah_diperiksa').count()
    telkom_unchecked = telkom_stations - telkom_checked
    
    # Telkomsel
    telkomsel_stations = Stasiun.query.filter_by(operator='telkomsel').count()
    telkomsel_uploaded = Stasiun.query.filter(Stasiun.operator == 'telkomsel', Stasiun.uploads.any()).count()
    telkomsel_not_uploaded = telkomsel_stations - telkomsel_uploaded
    telkomsel_checked = Stasiun.query.filter(Stasiun.operator == 'telkomsel', Stasiun.status_pemeriksaan == 'sudah_diperiksa').count()
    telkomsel_unchecked = telkomsel_stations - telkomsel_checked
    
    # IOH (ganti dari indosat)
    ioh_stations = Stasiun.query.filter_by(operator='ioh').count()
    ioh_uploaded = Stasiun.query.filter(Stasiun.operator == 'ioh', Stasiun.uploads.any()).count()
    ioh_not_uploaded = ioh_stations - ioh_uploaded
    ioh_checked = Stasiun.query.filter(Stasiun.operator == 'ioh', Stasiun.status_pemeriksaan == 'sudah_diperiksa').count()
    ioh_unchecked = ioh_stations - ioh_checked
    
    # XL Smart (ganti dari xl)
    xl_smart_stations = Stasiun.query.filter_by(operator='xl smart').count()
    xl_smart_uploaded = Stasiun.query.filter(Stasiun.operator == 'xl smart', Stasiun.uploads.any()).count()
    xl_smart_not_uploaded = xl_smart_stations - xl_smart_uploaded
    xl_smart_checked = Stasiun.query.filter(Stasiun.operator == 'xl smart', Stasiun.status_pemeriksaan == 'sudah_diperiksa').count()
    xl_smart_unchecked = xl_smart_stations - xl_smart_checked
    
    # ===== STATISTIK OPPONENTS PER OPERATOR =====
    telkom_opponents = StasiunLawan.query.join(Stasiun).filter(Stasiun.operator == 'telkom').count()
    telkomsel_opponents = StasiunLawan.query.join(Stasiun).filter(Stasiun.operator == 'telkomsel').count()
    ioh_opponents = StasiunLawan.query.join(Stasiun).filter(Stasiun.operator == 'ioh').count()
    xl_smart_opponents = StasiunLawan.query.join(Stasiun).filter(Stasiun.operator == 'xl smart').count()
    
    recent_stations = Stasiun.query.order_by(Stasiun.created_at.desc()).limit(5).all()
    recent_users = User.query.order_by(User.created_at.desc()).limit(5).all()
    current_date = datetime.now().strftime('%d %B %Y')
    
    # Get user statistics
    user_stats = {
        'admin_master': User.query.filter_by(role='admin_master').count(),
        'admin_operator': User.query.filter_by(role='admin_operator').count(),
        'user_operator': User.query.filter_by(role='user_operator').count(),
    }
    
    # ===== PERBAIKAN: Ambil daftar kota case-insensitive =====
    from sqlalchemy import func
    
    kota_raw = db.session.query(
        func.lower(Stasiun.kota).label('kota_lower'),
        Stasiun.kota
    ).distinct(func.lower(Stasiun.kota))\
     .order_by(func.lower(Stasiun.kota))\
     .all()
    
    # Mapping untuk menyimpan format standar
    kota_mapping = {}
    for kota_lower, kota_original in kota_raw:
        if kota_lower not in kota_mapping:
            # Cari format standar dari KOTA_LIST
            found = False
            for std_kota in KOTA_LIST:
                if std_kota.lower() == kota_lower:
                    kota_mapping[kota_lower] = std_kota
                    found = True
                    break
            if not found:
                kota_mapping[kota_lower] = kota_original.strip().title()
    
    kota_list = list(kota_mapping.values())
    
    # Jika tidak ada kota di database, gunakan KOTA_LIST sebagai fallback
    if not kota_list:
        kota_list = KOTA_LIST
    
    return render_template('admin_master/dashboard.html',
                         total_stations=total_stations,
                         total_opponents=total_opponents,
                         # Statistik Upload
                         stations_with_uploads=stations_with_uploads,
                         stations_without_uploads=stations_without_uploads,
                         # Statistik Pemeriksaan
                         stations_checked=stations_checked,
                         stations_unchecked=stations_unchecked,
                         # Per operator - stasiun
                         telkom_stations=telkom_stations,
                         telkomsel_stations=telkomsel_stations,
                         ioh_stations=ioh_stations,
                         xl_smart_stations=xl_smart_stations,
                         # Per operator - upload
                         telkom_uploaded=telkom_uploaded,
                         telkom_not_uploaded=telkom_not_uploaded,
                         telkomsel_uploaded=telkomsel_uploaded,
                         telkomsel_not_uploaded=telkomsel_not_uploaded,
                         ioh_uploaded=ioh_uploaded,
                         ioh_not_uploaded=ioh_not_uploaded,
                         xl_smart_uploaded=xl_smart_uploaded,
                         xl_smart_not_uploaded=xl_smart_not_uploaded,
                         # Per operator - pemeriksaan
                         telkom_checked=telkom_checked,
                         telkom_unchecked=telkom_unchecked,
                         telkomsel_checked=telkomsel_checked,
                         telkomsel_unchecked=telkomsel_unchecked,
                         ioh_checked=ioh_checked,
                         ioh_unchecked=ioh_unchecked,
                         xl_smart_checked=xl_smart_checked,
                         xl_smart_unchecked=xl_smart_unchecked,
                         # Per operator - opponents
                         telkom_opponents=telkom_opponents,
                         telkomsel_opponents=telkomsel_opponents,
                         ioh_opponents=ioh_opponents,
                         xl_smart_opponents=xl_smart_opponents,
                         recent_stations=recent_stations,
                         recent_users=recent_users,
                         user_stats=user_stats,
                         operators=OPERATORS,
                         kota_list=kota_list,
                         current_date=current_date)

# ============================================================================
# AI ANALYTICS & CHAT ROUTES
# ============================================================================

@app.route('/analisis-ai')
@login_required
def analisis_ai():
    """Universal redirector to AI analytics based on user role"""
    if current_user.role == 'admin_master':
        return redirect(url_for('admin_master_analisis_ai'))
    elif current_user.role == 'admin_operator':
        return redirect(url_for('admin_operator_analisis_ai'))
    elif current_user.role == 'user_operator':
        return redirect(url_for('user_operator_analisis_ai'))
    return redirect(url_for('index'))

@app.route('/admin-master/analisis-ai')
@login_required
@admin_master_required
def admin_master_analisis_ai():
    selected_operator = request.args.get('operator', 'all')
    selected_kota = request.args.get('kota', 'all')
    
    analytics = get_ai_analytics_data(
        db, Stasiun, StasiunLawan, UploadGambar, HasilLaporan,
        OPERATORS, KOTA_LIST,
        target_operator=selected_operator,
        target_kota=selected_kota
    )
    return render_template('admin_master/analisis_ai.html', analytics=analytics)

@app.route('/admin-operator/analisis-ai')
@login_required
@admin_operator_required
def admin_operator_analisis_ai():
    default_op = current_user.operator_type if current_user.operator_type else 'all'
    selected_operator = request.args.get('operator', default_op)
    selected_kota = request.args.get('kota', 'all')
    
    analytics = get_ai_analytics_data(
        db, Stasiun, StasiunLawan, UploadGambar, HasilLaporan,
        OPERATORS, KOTA_LIST,
        target_operator=selected_operator,
        target_kota=selected_kota
    )
    return render_template('admin_operator/analisis_ai.html', analytics=analytics)

@app.route('/user-operator/analisis-ai')
@login_required
@user_operator_required
def user_operator_analisis_ai():
    default_op = current_user.operator_type if current_user.operator_type else 'all'
    selected_operator = request.args.get('operator', default_op)
    selected_kota = request.args.get('kota', 'all')
    
    analytics = get_ai_analytics_data(
        db, Stasiun, StasiunLawan, UploadGambar, HasilLaporan,
        OPERATORS, KOTA_LIST,
        target_operator=selected_operator,
        target_kota=selected_kota
    )
    return render_template('user_operator/analisis_ai.html', analytics=analytics)

@app.route('/api/ai-chat', methods=['POST'])
@login_required
def api_ai_chat():
    try:
        data = request.get_json() or {}
        user_message = data.get('message', '').strip()
        if not user_message:
            return jsonify({'status': 'error', 'message': 'Pesan tidak boleh kosong'}), 400

        target_op = current_user.operator_type if current_user.role != 'admin_master' else 'all'
        analytics = get_ai_analytics_data(
            db, Stasiun, StasiunLawan, UploadGambar, HasilLaporan,
            OPERATORS, KOTA_LIST,
            target_operator=target_op
        )
        
        result = process_ai_query(user_message, analytics)
        return jsonify(result)
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)}), 500

# ============================================================================
# OCR (CAPTURE TO TEXT) ROUTES
# ============================================================================

@app.route('/admin-master/ocr', methods=['GET'])
@login_required
@admin_master_required
def admin_master_ocr():
    return render_template('admin_master/ocr.html')

@app.route('/admin-operator/ocr', methods=['GET'])
@login_required
@admin_operator_required
def admin_operator_ocr():
    return render_template('admin_operator/ocr.html')

@app.route('/user-operator/ocr', methods=['GET'])
@login_required
@user_operator_required
def user_operator_ocr():
    return render_template('user_operator/ocr.html')

rapid_ocr_instance = None

def get_rapid_ocr():
    global rapid_ocr_instance
    if rapid_ocr_instance is None:
        try:
            from rapidocr_onnxruntime import RapidOCR
            rapid_ocr_instance = RapidOCR(
                intra_op_num_threads=1,
                inter_op_num_threads=1
            )
        except Exception as e:
            print(f"Error initializing RapidOCR: {e}")
    return rapid_ocr_instance

# Pre-warm RapidOCR asynchronously at server startup so first request has zero cold-start delay
import threading
def _prewarm_rapidocr():
    try:
        get_rapid_ocr()
    except Exception as e:
        print(f"RapidOCR prewarm notice: {e}")

threading.Thread(target=_prewarm_rapidocr, daemon=True).start()

@app.route('/api/ocr/rapidocr', methods=['POST'])
@login_required
def api_rapidocr():
    """
    Endpoint pemrosesan OCR menggunakan mesin MWLink_Inspector (RapidOCR).
    Mendukung single image atau multiple images (misal Port Setting + Radio Configuration).
    Mengekstrak teks mentah OCR sekaligus 6 parameter microwave link secara terstruktur.
    """
    try:
        data = request.get_json()
        if not data:
            return jsonify({'status': 'error', 'message': 'Data payload tidak ditemukan'}), 400

        raw_images = []
        if 'images' in data and isinstance(data['images'], list) and data['images']:
            raw_images = data['images']
        elif 'image' in data and data['image']:
            raw_images = [data['image']]

        if not raw_images:
            return jsonify({'status': 'error', 'message': 'Data gambar tidak ditemukan'}), 400

        import base64
        import tempfile
        import shutil
        import io
        from PIL import Image
        from mwlink_ocr_engine import process_microwave_ocr

        temp_dir = tempfile.mkdtemp()
        temp_paths = []
        try:
            for idx, img_b64 in enumerate(raw_images):
                if ',' in img_b64:
                    img_b64 = img_b64.split(',')[1]
                
                img_bytes = base64.b64decode(img_b64)
                img = Image.open(io.BytesIO(img_bytes)).convert('RGB')
                
                # Batasi dimensi maksimal 1800px untuk stabilitas memori & kecepatan
                max_dim = 1800
                if max(img.size) > max_dim:
                    scale = max_dim / max(img.size)
                    img = img.resize((int(img.size[0] * scale), int(img.size[1] * scale)), Image.Resampling.LANCZOS)
                
                fpath = os.path.join(temp_dir, f"capture_{idx+1}.jpg")
                img.save(fpath, format='JPEG', quality=90)
                temp_paths.append(fpath)

            station_hint = data.get('station_hint')
            result_data = process_microwave_ocr(temp_paths, station_name_hint=station_hint)

            return jsonify({
                'status': 'success',
                'text': result_data.get('text', ''),
                'confidence': result_data.get('confidence', 95.0),
                'links': result_data.get('links', []),
                'elapse': result_data.get('elapse', 0.0)
            })
        finally:
            shutil.rmtree(temp_dir, ignore_errors=True)

    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)}), 500

@app.route('/admin-master/users')
@login_required
@admin_master_required
def admin_master_manage_users():
    users = User.query.order_by(User.created_at.desc()).all()
    return render_template('admin_master/manage_users.html', users=users)

@app.route('/admin-master/users/add', methods=['GET', 'POST'])
@login_required
@admin_master_required
def admin_master_add_user():
    if request.method == 'POST':
        try:
            username = request.form.get('username')
            password = request.form.get('password')
            role = request.form.get('role')
            operator_type = request.form.get('operator_type')
            
            if not username or not password or not role:
                flash('Username, password, dan role harus diisi!', 'error')
                return redirect(url_for('admin_master_add_user'))
            
            # Cek apakah username sudah ada
            existing_user = User.query.filter_by(username=username).first()
            if existing_user:
                flash('Username sudah digunakan!', 'error')
                return redirect(url_for('admin_master_add_user'))
            
            # Hash password
            hashed_password = generate_password_hash(password, method='pbkdf2:sha256')
            
            # Validasi operator_type berdasarkan role
            if role == 'admin_master':
                operator_type = None
            elif role in ['admin_operator', 'user_operator']:
                if not operator_type or operator_type not in OPERATORS:
                    flash('Operator type harus dipilih!', 'error')
                    return redirect(url_for('admin_master_add_user'))
            
            # Buat user baru
            new_user = User(
                username=username,
                password=hashed_password,
                role=role,
                operator_type=operator_type
            )
            
            db.session.add(new_user)
            db.session.commit()
            
            flash(f'User {username} berhasil ditambahkan!', 'success')
            return redirect(url_for('admin_master_manage_users'))
            
        except Exception as e:
            db.session.rollback()
            flash(f'Error: {str(e)}', 'error')
            return redirect(url_for('admin_master_add_user'))
    
    return render_template('admin_master/add_user.html', 
                         roles=['admin_master', 'admin_operator', 'user_operator'],
                         operators=OPERATORS)

@app.route('/admin-master/users/edit/<int:user_id>', methods=['GET', 'POST'])
@login_required
@admin_master_required
def admin_master_edit_user(user_id):
    user = User.query.get_or_404(user_id)
    
    if request.method == 'POST':
        try:
            username = request.form.get('username')
            password = request.form.get('password')
            role = request.form.get('role')
            operator_type = request.form.get('operator_type')
            
            if not username or not role:
                flash('Username dan role harus diisi!', 'error')
                return redirect(url_for('admin_master_edit_user', user_id=user_id))
            
            # Cek apakah username sudah ada (kecuali untuk user ini)
            existing_user = User.query.filter(User.username == username, User.id != user_id).first()
            if existing_user:
                flash('Username sudah digunakan!', 'error')
                return redirect(url_for('admin_master_edit_user', user_id=user_id))
            
            # Update user
            user.username = username
            user.role = role
            
            # Update password jika diisi
            if password:
                user.password = generate_password_hash(password, method='pbkdf2:sha256')
            
            # Update operator_type berdasarkan role
            if role == 'admin_master':
                user.operator_type = None
            elif role in ['admin_operator', 'user_operator']:
                if not operator_type or operator_type not in OPERATORS:
                    flash('Operator type harus dipilih!', 'error')
                    return redirect(url_for('admin_master_edit_user', user_id=user_id))
                user.operator_type = operator_type
            
            db.session.commit()
            
            flash(f'User {username} berhasil diperbarui!', 'success')
            return redirect(url_for('admin_master_manage_users'))
            
        except Exception as e:
            db.session.rollback()
            flash(f'Error: {str(e)}', 'error')
            return redirect(url_for('admin_master_edit_user', user_id=user_id))
    
    return render_template('admin_master/edit_user.html', 
                         user=user,
                         roles=['admin_master', 'admin_operator', 'user_operator'],
                         operators=OPERATORS)

@app.route('/admin-master/users/delete/<int:user_id>')
@login_required
@admin_master_required
def admin_master_delete_user(user_id):
    user = User.query.get_or_404(user_id)
    
    # Tidak boleh menghapus diri sendiri
    if user.id == current_user.id:
        flash('Tidak dapat menghapus akun sendiri!', 'error')
        return redirect(url_for('admin_master_manage_users'))
    
    try:
        db.session.delete(user)
        db.session.commit()
        flash(f'User {user.username} berhasil dihapus!', 'success')
    except Exception as e:
        db.session.rollback()
        flash(f'Error: {str(e)}', 'error')
    
    return redirect(url_for('admin_master_manage_users'))

# ============================================================================
# HAPUS SEMUA DATA (BULK DELETE)
# ============================================================================

@app.route('/admin-master/hapus-semua-data', methods=['POST'])
@login_required
@admin_master_required
def admin_master_hapus_semua_data():
    """Hapus semua data stasiun untuk Admin Master"""
    try:
        # Validasi konfirmasi
        confirm = request.form.get('confirm_delete')
        if confirm != 'HAPUS SEMUA':
            flash('Konfirmasi tidak valid! Ketik "HAPUS SEMUA" untuk melanjutkan.', 'error')
            return redirect(url_for('admin_master_daftar_stasiun'))
        
        # Hitung jumlah data sebelum dihapus
        total_stations = Stasiun.query.count()
        total_opponents = StasiunLawan.query.count()
        total_uploads = UploadGambar.query.count()
        total_status = StatusUpdate.query.count()
        total_groups = GrupStasiun.query.count()
        
        print(f"🗑️ MEMULAI PROSES HAPUS SEMUA DATA - ADMIN MASTER")
        print(f"   - Stasiun: {total_stations}")
        print(f"   - Stasiun Lawan: {total_opponents}")
        print(f"   - Upload Gambar: {total_uploads}")
        print(f"   - Status Update: {total_status}")
        print(f"   - Grup Stasiun: {total_groups}")
        
        # ===== 1. HAPUS SEMUA FILE DARI CLOUDINARY =====
        all_uploads = UploadGambar.query.all()
        cloudinary_deleted = 0
        cloudinary_failed = 0
        
        # Koleksi public_id unik
        unique_public_ids = set()
        for upload in all_uploads:
            unique_public_ids.add(upload.public_id)
        
        print(f"   - File unik di Cloudinary: {len(unique_public_ids)}")
        
        for public_id in unique_public_ids:
            try:
                success, msg = delete_from_cloudinary(public_id)
                if success:
                    cloudinary_deleted += 1
                    print(f"     ✓ {public_id}")
                else:
                    cloudinary_failed += 1
                    print(f"     ✗ {public_id}: {msg}")
            except Exception as e:
                cloudinary_failed += 1
                print(f"     ✗ {public_id}: {e}")
        
        # ===== 2. HAPUS SEMUA DATA DARI DATABASE =====
        # Hapus status updates
        StatusUpdate.query.delete()
        
        # Hapus upload gambar
        UploadGambar.query.delete()
        
        # Hapus stasiun lawan
        StasiunLawan.query.delete()
        
        # Hapus grup stasiun
        GrupStasiun.query.delete()
        
        # Hapus stasiun
        Stasiun.query.delete()
        
        # Commit perubahan
        db.session.commit()
        
        # ===== 3. TAMPILKAN HASIL =====
        message = f"""
        ✅ BERHASIL MENGHAPUS SEMUA DATA!
        
        📊 RINCIAN PENGHAPUSAN:
        • Stasiun: {total_stations} data
        • Stasiun Lawan: {total_opponents} data
        • Upload Gambar: {total_uploads} data
        • Status Update: {total_status} data
        • Grup Stasiun: {total_groups} data
        
        ☁️ CLOUDINARY:
        • File berhasil dihapus: {cloudinary_deleted} file
        • File gagal dihapus: {cloudinary_failed} file
        """
        
        flash(message, 'success')
        
    except Exception as e:
        db.session.rollback()
        error_msg = f"❌ ERROR: {str(e)}"
        print(error_msg)
        import traceback
        traceback.print_exc()
        flash(error_msg, 'error')
    
    return redirect(url_for('admin_master_daftar_stasiun'))

@app.route('/admin-operator/hapus-semua-data', methods=['POST'])
@login_required
@admin_operator_required
def admin_operator_hapus_semua_data():
    """Hapus semua data stasiun untuk operator tertentu"""
    try:
        operator = current_user.operator_type
        
        # Validasi konfirmasi
        confirm = request.form.get('confirm_delete')
        expected_confirm = f'HAPUS {operator.upper()}'
        if confirm != expected_confirm:
            flash(f'Konfirmasi tidak valid! Ketik "{expected_confirm}" untuk melanjutkan.', 'error')
            return redirect(url_for('admin_operator_daftar_stasiun'))
        
        # Hitung jumlah data sebelum dihapus
        stations = Stasiun.query.filter_by(operator=operator).all()
        station_ids = [s.id for s in stations]
        
        total_stations = len(stations)
        total_opponents = StasiunLawan.query.filter(StasiunLawan.stasiun_id.in_(station_ids)).count() if station_ids else 0
        
        # Hitung upload dan status
        total_uploads = 0
        total_status = 0
        total_groups = 0
        
        if station_ids:
            total_uploads = UploadGambar.query.filter(UploadGambar.stasiun_id.in_(station_ids)).count()
            total_groups = GrupStasiun.query.filter(GrupStasiun.stasiun_id.in_(station_ids)).count()
            
            # Hitung status updates
            opponent_ids = [o.id for o in StasiunLawan.query.filter(StasiunLawan.stasiun_id.in_(station_ids)).all()]
            if opponent_ids:
                total_status = StatusUpdate.query.filter(StatusUpdate.stasiun_lawan_id.in_(opponent_ids)).count()
        
        print(f"🗑️ MEMULAI PROSES HAPUS SEMUA DATA - OPERATOR {operator.upper()}")
        print(f"   - Stasiun: {total_stations}")
        print(f"   - Stasiun Lawan: {total_opponents}")
        print(f"   - Upload Gambar: {total_uploads}")
        print(f"   - Status Update: {total_status}")
        print(f"   - Grup Stasiun: {total_groups}")
        
        if not station_ids:
            flash(f'Tidak ada data untuk operator {operator.upper()}', 'info')
            return redirect(url_for('admin_operator_daftar_stasiun'))
        
        # ===== 1. HAPUS SEMUA FILE DARI CLOUDINARY =====
        all_uploads = UploadGambar.query.filter(UploadGambar.stasiun_id.in_(station_ids)).all()
        cloudinary_deleted = 0
        cloudinary_failed = 0
        
        # Koleksi public_id unik
        unique_public_ids = set()
        for upload in all_uploads:
            unique_public_ids.add(upload.public_id)
        
        print(f"   - File unik di Cloudinary: {len(unique_public_ids)}")
        
        for public_id in unique_public_ids:
            try:
                # Cek apakah public_id ini masih digunakan oleh operator lain
                other_refs = UploadGambar.query\
                    .filter(UploadGambar.public_id == public_id,
                           ~UploadGambar.stasiun_id.in_(station_ids))\
                    .count()
                
                if other_refs == 0:
                    success, msg = delete_from_cloudinary(public_id)
                    if success:
                        cloudinary_deleted += 1
                        print(f"     ✓ {public_id}")
                    else:
                        cloudinary_failed += 1
                        print(f"     ✗ {public_id}: {msg}")
                else:
                    print(f"     ⚠ {public_id} masih digunakan operator lain, dilewati")
                    cloudinary_deleted += 1  # Hitung sebagai berhasil karena tidak perlu dihapus
            except Exception as e:
                cloudinary_failed += 1
                print(f"     ✗ {public_id}: {e}")
        
        # ===== 2. HAPUS DATA DARI DATABASE =====
        # Hapus status updates
        if opponent_ids:
            StatusUpdate.query.filter(StatusUpdate.stasiun_lawan_id.in_(opponent_ids)).delete()
        
        # Hapus upload gambar
        UploadGambar.query.filter(UploadGambar.stasiun_id.in_(station_ids)).delete()
        
        # Hapus stasiun lawan
        StasiunLawan.query.filter(StasiunLawan.stasiun_id.in_(station_ids)).delete()
        
        # Hapus grup stasiun
        GrupStasiun.query.filter(GrupStasiun.stasiun_id.in_(station_ids)).delete()
        
        # Hapus stasiun
        for station in stations:
            db.session.delete(station)
        
        # Commit perubahan
        db.session.commit()
        
        # ===== 3. TAMPILKAN HASIL =====
        message = f"""
        ✅ BERHASIL MENGHAPUS SEMUA DATA OPERATOR {operator.upper()}!
        
        📊 RINCIAN PENGHAPUSAN:
        • Stasiun: {total_stations} data
        • Stasiun Lawan: {total_opponents} data
        • Upload Gambar: {total_uploads} data
        • Status Update: {total_status} data
        • Grup Stasiun: {total_groups} data
        
        ☁️ CLOUDINARY:
        • File berhasil dihapus: {cloudinary_deleted} file
        • File gagal dihapus: {cloudinary_failed} file
        """
        
        flash(message, 'success')
        
    except Exception as e:
        db.session.rollback()
        error_msg = f"❌ ERROR: {str(e)}"
        print(error_msg)
        import traceback
        traceback.print_exc()
        flash(error_msg, 'error')
    
    return redirect(url_for('admin_operator_daftar_stasiun'))

@app.route('/admin-master/tambah-data', methods=['GET', 'POST'])
@login_required
@admin_master_required
def admin_master_tambah_data():
    if request.method == 'POST':
        try:
            stasiun_name = request.form.get('stasiun_name').strip()
            operator = request.form.get('operator')
            kota = request.form.get('kota')
            
            stasiun_lawan_list = request.form.getlist('stasiun_lawan[]')
            freq_tx_list = request.form.getlist('freq_tx[]')
            freq_rx_list = request.form.getlist('freq_rx[]')
            
            if not stasiun_name:
                flash('Nama stasiun harus diisi!', 'error')
                return redirect(url_for('admin_master_tambah_data'))
            
            new_station = Stasiun(
                stasiun_name=stasiun_name,
                operator=operator,
                kota=kota,
                created_by=current_user.id
            )
            db.session.add(new_station)
            db.session.flush()
            
            for i, nama_lawan in enumerate(stasiun_lawan_list):
                nama_lawan = nama_lawan.strip()
                if nama_lawan:
                    freq_tx = freq_tx_list[i].strip() if i < len(freq_tx_list) else ''
                    freq_rx = freq_rx_list[i].strip() if i < len(freq_rx_list) else ''
                    
                    new_lawan = StasiunLawan(
                        stasiun_id=new_station.id,
                        nama_stasiun_lawan=nama_lawan,
                        freq_tx=freq_tx if freq_tx else None,
                        freq_rx=freq_rx if freq_rx else None,
                        group_id=None,
                        urutan=i
                    )
                    db.session.add(new_lawan)
            
            db.session.commit()
            flash(f'Stasiun "{stasiun_name}" berhasil ditambahkan!', 'success')
            return redirect(url_for('admin_master_daftar_stasiun'))
            
        except Exception as e:
            db.session.rollback()
            flash(f'Error: {str(e)}', 'error')
            return redirect(url_for('admin_master_tambah_data'))
    
    return render_template('admin_master/tambah_data.html',
                         operators=OPERATORS,
                         kota_list=KOTA_LIST)

def _handle_admin_master_daftar_stasiun(template_name='admin_master/daftar_stasiun.html'):
    operator = request.args.get('operator', 'all')
    kota = request.args.get('kota', 'all')
    search_stasiun = request.args.get('search_stasiun', '').strip()
    status_pemeriksaan = request.args.get('status_pemeriksaan', 'all')
    status_lengkap = request.args.get('status_lengkap', 'all')
    
    # ===== PARAMETER PAGINATION =====
    per_page = request.args.get('per_page', 10, type=int)
    if per_page not in [10, 20, 50, 100]:
        per_page = 10
    page = request.args.get('page', 1, type=int)
    
    # ===== QUERY DASAR DENGAN PAGINATION =====
    query = Stasiun.query
    
    if operator != 'all':
        query = query.filter_by(operator=operator)
    
    if kota != 'all' and kota:
        query = query.filter(Stasiun.kota.ilike(f'%{kota}%'))
    
    if search_stasiun:
        query = query.filter(Stasiun.stasiun_name.ilike(f'%{search_stasiun}%'))
    
    if status_pemeriksaan != 'all':
        query = query.filter(Stasiun.status_pemeriksaan == status_pemeriksaan)
    
    # ===== PAGINATION DI DATABASE LEVEL =====
    pagination = query.order_by(Stasiun.created_at.desc()).paginate(
        page=page, per_page=per_page, error_out=False
    )
    stations = pagination.items
    
    if not stations:
        # Jika tidak ada data, tetap tampilkan halaman kosong
        from sqlalchemy import func
        kota_raw = db.session.query(
            func.lower(Stasiun.kota).label('kota_lower'),
            Stasiun.kota
        ).distinct(func.lower(Stasiun.kota)).order_by(func.lower(Stasiun.kota)).all()
        
        kota_mapping = {}
        for kota_lower, kota_original in kota_raw:
            if kota_lower not in kota_mapping:
                found = False
                for std_kota in KOTA_LIST:
                    if std_kota.lower() == kota_lower:
                        kota_mapping[kota_lower] = std_kota
                        found = True
                        break
                if not found:
                    kota_mapping[kota_lower] = kota_original.strip().title()
        
        kota_list = list(kota_mapping.values())
        if not kota_list:
            kota_list = KOTA_LIST
        
        selected_kota = kota
        if kota != 'all':
            found = False
            for k in kota_list:
                if k.lower() == kota.lower():
                    selected_kota = k
                    found = True
                    break
            if not found:
                selected_kota = 'all'
        
        return render_template(template_name,
                             stations=[],
                             pagination=pagination,
                             operators=OPERATORS,
                             kota_list=kota_list,
                             selected_operator=operator,
                             selected_kota=selected_kota,
                             selected_status=status_pemeriksaan,
                             selected_status_lengkap=status_lengkap,
                             status_summary={'lengkap': 0, 'belum_lengkap': 0},
                             status_options=STATUS_OPTIONS)
    
    # ===== BATCH QUERY UNTUK SEMUA STASIUN DALAM 1 HALAMAN =====
    station_ids = [s.id for s in stations]
    
    # 1. Batch query semua stasiun lawan (1 query)
    all_opponents = StasiunLawan.query\
        .filter(StasiunLawan.stasiun_id.in_(station_ids))\
        .order_by(StasiunLawan.urutan)\
        .all()
    
    # 2. Batch query semua status terbaru (1 query)
    opponent_ids = [o.id for o in all_opponents]
    latest_status_dict = {}
    
    if opponent_ids:
        from sqlalchemy import func
        subquery = db.session.query(
            StatusUpdate.stasiun_lawan_id,
            func.max(StatusUpdate.updated_at).label('max_updated')
        ).filter(StatusUpdate.stasiun_lawan_id.in_(opponent_ids))\
         .group_by(StatusUpdate.stasiun_lawan_id)\
         .subquery()
        
        latest_statuses = db.session.query(StatusUpdate)\
            .join(subquery,
                  (StatusUpdate.stasiun_lawan_id == subquery.c.stasiun_lawan_id) &
                  (StatusUpdate.updated_at == subquery.c.max_updated))\
            .all()
        
        for status in latest_statuses:
            latest_status_dict[status.stasiun_lawan_id] = status
    
    # 3. Batch query semua grup (1 query)
    all_groups = GrupStasiun.query.filter(GrupStasiun.stasiun_id.in_(station_ids)).all()
    groups_by_station = {}
    for group in all_groups:
        if group.stasiun_id not in groups_by_station:
            groups_by_station[group.stasiun_id] = []
        groups_by_station[group.stasiun_id].append(group)
    
    # 4. Batch query user untuk checked_by (1 query)
    user_ids = [s.checked_by for s in stations if s.checked_by]
    users_dict = {}
    if user_ids:
        users = User.query.filter(User.id.in_(user_ids)).all()
        users_dict = {u.id: u for u in users}
    
    # ===== Kelompokkan lawan per stasiun =====
    opponents_by_station = {}
    for opponent in all_opponents:
        if opponent.stasiun_id not in opponents_by_station:
            opponents_by_station[opponent.stasiun_id] = []
        opponents_by_station[opponent.stasiun_id].append(opponent)
    
    # ===== Proses setiap stasiun =====
    for station in stations:
        station.stasiun_lawan_list = opponents_by_station.get(station.id, [])
        station.grup_list = groups_by_station.get(station.id, [])
        station.checked_by_user = users_dict.get(station.checked_by)
        
        # Hitung status untuk setiap lawan
        total_lawan = len(station.stasiun_lawan_list)
        lawan_tanpa_status = 0
        
        for lawan in station.stasiun_lawan_list:
            latest_status = latest_status_dict.get(lawan.id)
            lawan.latest_status = latest_status
            if not latest_status:
                lawan_tanpa_status += 1
        
        station.has_lawan_without_status = lawan_tanpa_status > 0
        
        # ===== Build groups display =====
        groups_dict = {g.id: g.nama_grup for g in station.grup_list}
        
        lawan_by_group = {}
        group_stats = {}
        
        for lawan in station.stasiun_lawan_list:
            group_key = lawan.group_id if lawan.group_id is not None else 'ungrouped'
            if group_key not in lawan_by_group:
                lawan_by_group[group_key] = []
                group_stats[group_key] = {
                    'total': 0, 'aktif': 0, 'tidak_aktif': 0, 'tidak_sesuai': 0, 'belum_ada': 0
                }
            
            lawan_by_group[group_key].append(lawan)
            group_stats[group_key]['total'] += 1
            
            if lawan.latest_status:
                status = lawan.latest_status.status
                if status in group_stats[group_key]:
                    group_stats[group_key][status] += 1
            else:
                group_stats[group_key]['belum_ada'] += 1
        
        groups_display = []
        
        if 'ungrouped' in lawan_by_group:
            groups_display.append({
                'group_id': None,
                'group_name': 'Tanpa Grup',
                'lawans': lawan_by_group['ungrouped'],
                'stats': group_stats['ungrouped'],
                'is_ungrouped': True,
                'badge_color': 'warning'
            })
        
        for group_key, lawans in lawan_by_group.items():
            if group_key != 'ungrouped':
                try:
                    group_id = int(group_key)
                    group_name = groups_dict.get(group_id, f'Grup {group_id}')
                    groups_display.append({
                        'group_id': group_id,
                        'group_name': group_name,
                        'lawans': lawans,
                        'stats': group_stats[group_key],
                        'is_ungrouped': False,
                        'badge_color': 'success'
                    })
                except:
                    continue
        
        groups_display.sort(key=lambda x: (x['is_ungrouped'], x['group_id'] if x['group_id'] is not None else 0))
        station.groups_display = groups_display
    
    # ===== Filter berdasarkan status_lengkap =====
    if status_lengkap != 'all':
        filtered_stations = []
        for station in stations:
            total_lawan = len(station.stasiun_lawan_list)
            lawan_tanpa_status = sum(1 for l in station.stasiun_lawan_list if not l.latest_status)
            is_lengkap = (lawan_tanpa_status == 0 and total_lawan > 0)
            
            if status_lengkap == 'lengkap' and is_lengkap:
                filtered_stations.append(station)
            elif status_lengkap == 'belum_lengkap' and not is_lengkap and total_lawan > 0:
                filtered_stations.append(station)
        
        stations = filtered_stations
        pagination.items = stations
        pagination.total = len(filtered_stations)
    
    # ===== Hitung status summary untuk badge =====
    status_summary = {'lengkap': 0, 'belum_lengkap': 0}
    for station in stations:
        total_lawan = len(station.stasiun_lawan_list)
        lawan_tanpa_status = sum(1 for l in station.stasiun_lawan_list if not l.latest_status)
        if lawan_tanpa_status == 0 and total_lawan > 0:
            status_summary['lengkap'] += 1
        elif total_lawan > 0:
            status_summary['belum_lengkap'] += 1
    
    # ===== Ambil daftar kota untuk filter =====
    from sqlalchemy import func
    
    kota_raw = db.session.query(
        func.lower(Stasiun.kota).label('kota_lower'),
        Stasiun.kota
    ).distinct(func.lower(Stasiun.kota)).order_by(func.lower(Stasiun.kota)).all()
    
    kota_mapping = {}
    for kota_lower, kota_original in kota_raw:
        if kota_lower not in kota_mapping:
            found = False
            for std_kota in KOTA_LIST:
                if std_kota.lower() == kota_lower:
                    kota_mapping[kota_lower] = std_kota
                    found = True
                    break
            if not found:
                kota_mapping[kota_lower] = kota_original.strip().title()
    
    kota_list = list(kota_mapping.values())
    if not kota_list:
        kota_list = KOTA_LIST
    
    selected_kota = kota
    if kota != 'all':
        found = False
        for k in kota_list:
            if k.lower() == kota.lower():
                selected_kota = k
                found = True
                break
        if not found:
            selected_kota = 'all'
    
    return render_template(template_name,
                         stations=stations,
                         pagination=pagination,
                         operators=OPERATORS,
                         kota_list=kota_list,
                         selected_operator=operator,
                         selected_kota=selected_kota,
                         selected_status=status_pemeriksaan,
                         selected_status_lengkap=status_lengkap,
                         status_summary=status_summary,
                         status_options=STATUS_OPTIONS)

@app.route('/admin-master/daftar-stasiun')
@login_required
@admin_master_required
def admin_master_daftar_stasiun():
    return _handle_admin_master_daftar_stasiun('admin_master/daftar_stasiun.html')

@app.route('/admin-master/daftar-stasiun-v2')
@login_required
@admin_master_required
def admin_master_daftar_stasiun_v2():
    return _handle_admin_master_daftar_stasiun('admin_master/daftar_stasiun_v2.html')

def build_stasiun_lawan_dict(lawan, station, latest_status=None, status_history=None, updater_info=None):
    """Helper terpusat untuk membentuk kamus data stasiun lawan dengan data SIMS dan hasil pemeriksaan"""
    op_lower = (station.operator or '').lower()
    if 'telkomsel' in op_lower:
        default_client = '00105553 - PT TELEKOMUNIKASI SELULAR'
    elif 'telkom' in op_lower:
        default_client = '00105553 - TELEKOMUNIKASI INDONESIA TBK'
    elif 'ioh' in op_lower or 'indosat' in op_lower:
        default_client = '00104442 - INDOSAT OOREDOO HUTCHISON'
    elif 'xl' in op_lower:
        default_client = '00103331 - XL AXIATA TBK'
    else:
        default_client = f"00105553 - {station.operator.upper()}"

    client_operator = lawan.client_operator or default_client
    nomor_isr = lawan.nomor_isr or '02877478-0005U/2620242029'
    link_id = lawan.link_id or '15560635'
    id_sims = lawan.id_sims or '3500864'
    sertifikat_perangkat = lawan.sertifikat_perangkat or 'Ada'
    mulai_beroperasi = lawan.mulai_beroperasi or '-'

    def _normalize_bw_mhz(val, default='28'):
        if not val:
            return default
        s = str(val).strip()
        try:
            num = float(s)
            if num > 500:
                num = num / 1000.0
            return str(int(round(num))) if num.is_integer() else str(round(num, 2))
        except ValueError:
            return s.replace('kHz', '').replace('khz', '').strip() or default

    # Baseline Nilai ISR (SIMS) dari data stasiun / stasiun lawan
    isr_freq_tx = lawan.freq_tx or '-'
    isr_freq_rx = lawan.freq_rx or '-'
    isr_bw_mhz = _normalize_bw_mhz(lawan.bw_mhz, default='28')
    isr_tipe_perangkat = lawan.tipe_perangkat or 'NEC iPasolink VR10 Radio Frekuensi'
    isr_longitude = lawan.longitude or '116.71388889'
    isr_latitude = lawan.latitude or '-0.97194444'

    # Hasil Pemeriksaan (Dapat Diedit)
    hasil_freq_tx = lawan.hasil_freq_tx if lawan.hasil_freq_tx is not None else (lawan.freq_tx or '15215')
    hasil_freq_rx = lawan.hasil_freq_rx if lawan.hasil_freq_rx is not None else (lawan.freq_rx or '14725')
    hasil_bw_mhz = _normalize_bw_mhz(lawan.hasil_bw_mhz if lawan.hasil_bw_mhz is not None else lawan.bw_mhz, default='28')
    hasil_tipe_perangkat = lawan.hasil_tipe_perangkat if lawan.hasil_tipe_perangkat is not None else (lawan.tipe_perangkat or 'NEC iPasolink VR10 Radio Frekuensi')
    hasil_longitude = lawan.hasil_longitude if lawan.hasil_longitude is not None else (lawan.longitude or '116.71388889')
    hasil_latitude = lawan.hasil_latitude if lawan.hasil_latitude is not None else (lawan.latitude or '-0.97194444')

    return {
        'id': lawan.id,
        'nama': lawan.nama_stasiun_lawan,
        'client_operator': client_operator,
        'nomor_isr': nomor_isr,
        'link_id': link_id,
        'id_sims': id_sims,
        'sertifikat_perangkat': sertifikat_perangkat,
        'mulai_beroperasi': mulai_beroperasi,
        'isr_freq_tx': isr_freq_tx,
        'isr_freq_rx': isr_freq_rx,
        'isr_bw_mhz': isr_bw_mhz,
        'isr_tipe_perangkat': isr_tipe_perangkat,
        'isr_longitude': isr_longitude,
        'isr_latitude': isr_latitude,
        'hasil_freq_tx': hasil_freq_tx,
        'hasil_freq_rx': hasil_freq_rx,
        'hasil_bw_mhz': hasil_bw_mhz,
        'hasil_tipe_perangkat': hasil_tipe_perangkat,
        'hasil_longitude': hasil_longitude,
        'hasil_latitude': hasil_latitude,
        'freq_tx': lawan.freq_tx,
        'freq_rx': lawan.freq_rx,
        'bw_mhz': isr_bw_mhz,
        'tipe_perangkat': isr_tipe_perangkat,
        'hasil_pemeriksaan': lawan.hasil_pemeriksaan or 'Sesuai ISR',
        'metode_pemeriksaan': lawan.metode_pemeriksaan or 'Pemeriksaan melalui Remote Site',
        'tanggal_pemeriksaan': lawan.tanggal_pemeriksaan.strftime('%Y-%m-%d') if lawan.tanggal_pemeriksaan else None,
        'catatan_pemeriksaan': lawan.catatan_pemeriksaan or '',
        'latest_status': latest_status.status if latest_status else None,
        'status_display': STATUS_OPTIONS.get(latest_status.status, 'Belum Ada') if latest_status else 'Belum Ada',
        'catatan': latest_status.catatan if latest_status else '',
        'status_updated_at': latest_status.updated_at if latest_status else None,
        'status_updated_by': updater_info,
        'updated_at': lawan.updated_at or (latest_status.updated_at if latest_status else None),
        'updated_by': updater_info,
        'status_history': [{
            'status': s.status,
            'status_display': STATUS_OPTIONS.get(s.status, s.status),
            'catatan': s.catatan,
            'updated_at': s.updated_at,
            'updated_by': s.updated_by
        } for s in status_history] if status_history else []
    }


@app.route('/admin-master/stasiun/<int:station_id>')
@login_required
@admin_master_required


def admin_master_detail_stasiun(station_id):
    station = Stasiun.query.get_or_404(station_id)
    
    # ===== AMBIL PARAMETER NEXT UNTUK KEMBALI KE HALAMAN SEBELUMNYA =====
    next_url = request.args.get('next', '')
    
    # ===== AMBIL INFORMASI PEMERIKSA =====
    station.checked_by_user = None
    if station.checked_by:
        station.checked_by_user = User.query.get(station.checked_by)
    
    # ===== AMBIL SEMUA LAWAN =====
    all_lawans = StasiunLawan.query\
        .filter_by(stasiun_id=station_id)\
        .order_by(StasiunLawan.urutan)\
        .all()
    
    # ===== AMBIL SEMUA GRUP =====
    all_groups = GrupStasiun.query.filter_by(stasiun_id=station_id).all()
    groups_dict = {g.id: g.nama_grup for g in all_groups}
    
    # ===== AMBIL SEMUA UPLOAD UNIK =====
    uploads_query = db.session.query(
        UploadGambar.group_id,
        UploadGambar.public_id,
        db.func.max(UploadGambar.id).label('latest_id')
    ).filter_by(stasiun_id=station_id)\
     .group_by(UploadGambar.group_id, UploadGambar.public_id)\
     .subquery()
    
    all_uploads = UploadGambar.query\
        .join(uploads_query, UploadGambar.id == uploads_query.c.latest_id)\
        .order_by(UploadGambar.uploaded_at.desc())\
        .all()
    
    # ===== KELOMPOKKAN DATA PER GRUP =====
    groups = {}
    
    # Proses lawan per grup
    for lawan in all_lawans:
        group_key = lawan.group_id if lawan.group_id is not None else 'ungrouped'
        
        if group_key not in groups:
            groups[group_key] = {
                'group_id': lawan.group_id,
                'nama_grup': groups_dict.get(lawan.group_id) if lawan.group_id else None,
                'opponents': [],
                'uploads': [],
                'status_summary': {'aktif': 0, 'tidak_aktif': 0, 'tidak_sesuai': 0}
            }
        
        # Ambil status terbaru
        latest_status = StatusUpdate.query\
            .filter_by(stasiun_lawan_id=lawan.id)\
            .order_by(StatusUpdate.updated_at.desc())\
            .first()
        
        # Ambil status update sebelumnya untuk riwayat (opsional)
        status_history = StatusUpdate.query\
            .filter_by(stasiun_lawan_id=lawan.id)\
            .order_by(StatusUpdate.updated_at.desc())\
            .limit(5)\
            .all()
        
        lawan_data = build_stasiun_lawan_dict(
            lawan, station, latest_status, status_history, 
            updater_info=latest_status.updated_by if latest_status else None
        )
        
        groups[group_key]['opponents'].append(lawan_data)
        
        # Hitung statistik
        if latest_status and latest_status.status in groups[group_key]['status_summary']:
            groups[group_key]['status_summary'][latest_status.status] += 1
    
    # Proses upload per grup
    for upload in all_uploads:
        group_key = upload.group_id if upload.group_id is not None else 'ungrouped'
        
        if group_key not in groups:
            groups[group_key] = {
                'group_id': upload.group_id,
                'nama_grup': groups_dict.get(upload.group_id) if upload.group_id else None,
                'opponents': [],
                'uploads': [],
                'status_summary': {'aktif': 0, 'tidak_aktif': 0, 'tidak_sesuai': 0}
            }
        
        # Cek duplikasi public_id dalam grup yang sama
        existing_public_ids = [u.get('public_id') for u in groups[group_key]['uploads']]
        if upload.public_id not in existing_public_ids:
            # Dapatkan informasi uploader
            uploader = User.query.get(upload.uploaded_by) if upload.uploaded_by else None
            
            # Dapatkan thumbnail URL
            thumbnail_url = None
            if upload.public_id:
                thumbnail_url, _ = cloudinary_url(
                    upload.public_id, 
                    width=200, 
                    height=200, 
                    crop="fill", 
                    quality="auto"
                )
            
            upload_data = {
                'id': upload.id,
                'public_id': upload.public_id,
                'cloudinary_url': upload.cloudinary_url,
                'thumbnail_url': thumbnail_url,
                'original_filename': upload.original_filename,
                'uploaded_at': upload.uploaded_at,
                'uploaded_by': uploader.username if uploader else 'System',
                'status': upload.status,
                'format': upload.format,
                'width': upload.width,
                'height': upload.height,
                'bytes_size': upload.bytes_size,
                'size_mb': round(upload.bytes_size / (1024 * 1024), 2) if upload.bytes_size else None
            }
            groups[group_key]['uploads'].append(upload_data)
    
    # ===== URUTKAN GRUP =====
    sorted_groups = {}
    
    # Ungrouped dulu
    if 'ungrouped' in groups:
        sorted_groups['ungrouped'] = groups.pop('ungrouped')
    
    # Grup numeric diurutkan
    numeric_groups = {}
    for key, data in groups.items():
        if key != 'ungrouped':
            try:
                numeric_groups[int(key)] = data
            except:
                numeric_groups[key] = data
    
    for key in sorted(numeric_groups.keys()):
        sorted_groups[str(key)] = numeric_groups[key]
    
    # ===== HITUNG STATISTIK TOTAL =====
    status_counts = {'aktif': 0, 'tidak_aktif': 0, 'tidak_sesuai': 0}
    for lawan in all_lawans:
        latest_status = StatusUpdate.query\
            .filter_by(stasiun_lawan_id=lawan.id)\
            .order_by(StatusUpdate.updated_at.desc())\
            .first()
        if latest_status and latest_status.status in status_counts:
            status_counts[latest_status.status] += 1
    
    # ===== HITUNG STATISTIK UPLOAD =====
    total_uploads = len(all_uploads)
    total_upload_size = sum([u.bytes_size for u in all_uploads if u.bytes_size]) if all_uploads else 0
    total_upload_size_mb = round(total_upload_size / (1024 * 1024), 2) if total_upload_size else 0
    
    # ===== STATISTIK PER BULAN (untuk grafik) =====
    uploads_by_month = {}
    status_updates_by_month = {}
    
    for upload in all_uploads:
        month_key = upload.uploaded_at.strftime('%Y-%m')
        uploads_by_month[month_key] = uploads_by_month.get(month_key, 0) + 1
    
    for lawan in all_lawans:
        status_updates = StatusUpdate.query\
            .filter_by(stasiun_lawan_id=lawan.id)\
            .order_by(StatusUpdate.updated_at)\
            .all()
        
        for status_update in status_updates:
            month_key = status_update.updated_at.strftime('%Y-%m')
            if month_key not in status_updates_by_month:
                status_updates_by_month[month_key] = {'aktif': 0, 'tidak_aktif': 0, 'tidak_sesuai': 0}
            if status_update.status in status_updates_by_month[month_key]:
                status_updates_by_month[month_key][status_update.status] += 1
    
    # ===== HITUNG TOTAL OPPONENTS DENGAN STATUS =====
    total_opponents = len(all_lawans)
    opponents_with_status = sum(1 for lawan in all_lawans if StatusUpdate.query.filter_by(stasiun_lawan_id=lawan.id).first())
    opponents_without_status = total_opponents - opponents_with_status
    
    # ===== AMBIL UPLOAD TERBARU UNTUK PREVIEW =====
    recent_uploads = sorted(all_uploads, key=lambda x: x.uploaded_at, reverse=True)[:5]
    
    # ===== TAMBAHKAN station.grup_list UNTUK TEMPLATE =====
    station.grup_list = all_groups
    
    return render_template('admin_master/detail_stasiun.html',
                         station=station,
                         groups=sorted_groups,
                         all_uploads=all_uploads,
                         recent_uploads=recent_uploads,
                         total_opponents=total_opponents,
                         status_counts=status_counts,
                         uploads_by_month=uploads_by_month,
                         status_updates_by_month=status_updates_by_month,
                         total_uploads=total_uploads,
                         total_upload_size_mb=total_upload_size_mb,
                         opponents_with_status=opponents_with_status,
                         opponents_without_status=opponents_without_status,
                         checked_by_user=station.checked_by_user,
                         next_url=next_url)

@app.route('/admin-master/edit-stasiun/<int:station_id>', methods=['GET', 'POST'])
@login_required
@admin_master_required
def admin_master_edit_stasiun(station_id):
    station = Stasiun.query.get_or_404(station_id)
    
    # ===== AMBIL PARAMETER NEXT UNTUK KEMBALI KE HALAMAN SEBELUMNYA =====
    next_url = request.args.get('next', '')
    
    # ===== DEBUG: TAMPILKAN NILAI KOTA UNTUK MEMASTIKAN =====
    print(f"\n{'='*60}")
    print(f"📝 EDIT STASIUN - ID: {station_id}")
    print(f"   Nama Stasiun: {station.stasiun_name}")
    print(f"   Operator: {station.operator}")
    print(f"   Kota di Database: '{station.kota}'")
    print(f"   Next URL: {next_url}")
    print(f"{'='*60}\n")
    
    # ===== AMBIL DAFTAR KOTA DARI DATABASE UNTUK DROPDOWN =====
    kota_list_from_db = db.session.query(Stasiun.kota).distinct().order_by(Stasiun.kota).all()
    kota_list = [k[0] for k in kota_list_from_db if k[0]]
    
    # Jika tidak ada kota di database, gunakan KOTA_LIST sebagai fallback
    if not kota_list:
        kota_list = KOTA_LIST
        print(f"⚠️ Tidak ada kota di database, menggunakan KOTA_LIST default")
    
    # ===== PASTIKAN station.kota ADA DI DALAM kota_list =====
    if station.kota and station.kota not in kota_list:
        print(f"⚠️ Peringatan: Kota '{station.kota}' tidak ada di kota_list!")
        print(f"   Menambahkan kota '{station.kota}' ke dalam dropdown")
        kota_list.append(station.kota)
        kota_list.sort()
    
    print(f"📋 Daftar kota yang tersedia: {kota_list}")
    print(f"📍 Kota yang dipilih: '{station.kota}'")
    print(f"{'='*60}\n")
    
    if request.method == 'POST':
        try:
            # ===== VALIDASI INPUT =====
            stasiun_name = request.form.get('stasiun_name', '').strip()
            operator = request.form.get('operator', '').strip()
            kota = request.form.get('kota', '').strip()
            
            print(f"\n📝 POST DATA:")
            print(f"   Nama: {stasiun_name}")
            print(f"   Operator: {operator}")
            print(f"   Kota: '{kota}'")
            print(f"{'-'*40}")
            
            if not stasiun_name:
                flash('❌ Nama stasiun harus diisi!', 'error')
                return redirect(url_for('admin_master_edit_stasiun', station_id=station_id, next=next_url))
            
            if not operator:
                flash('❌ Operator harus dipilih!', 'error')
                return redirect(url_for('admin_master_edit_stasiun', station_id=station_id, next=next_url))
            
            # ===== VALIDASI KOTA - PASTIKAN TIDAK KOSONG =====
            if not kota:
                flash('❌ Kota/Kabupaten harus dipilih!', 'error')
                return redirect(url_for('admin_master_edit_stasiun', station_id=station_id, next=next_url))
            
            # ===== NORMALISASI KOTA =====
            kota_normalized = None
            for k in KOTA_LIST:
                if k.lower() == kota.lower():
                    kota_normalized = k
                    break
            
            # Coba cari yang mirip (misal: "Samarinda" vs "Kota Samarinda")
            if not kota_normalized:
                for k in KOTA_LIST:
                    if kota.lower() in k.lower() or k.lower() in kota.lower():
                        kota_normalized = k
                        break
            
            if not kota_normalized:
                # Jika tidak ditemukan, gunakan nilai asli dengan peringatan
                kota_normalized = kota
                flash(f'⚠️ Kota "{kota}" tidak ada dalam daftar standar, menggunakan nilai asli', 'warning')
            
            # Catat perubahan untuk pesan
            changes = []
            
            # ===== 1. UPDATE DATA STASIUN UTAMA =====
            if station.stasiun_name != stasiun_name:
                changes.append(f"📝 Nama stasiun: '{station.stasiun_name}' → '{stasiun_name}'")
                station.stasiun_name = stasiun_name
            
            if station.operator != operator:
                changes.append(f"🔄 Operator: '{station.operator}' → '{operator}'")
                station.operator = operator
            
            if station.kota != kota_normalized:
                changes.append(f"📍 Kota/Kab: '{station.kota}' → '{kota_normalized}'")
                station.kota = kota_normalized
            
            # ===== 2. AMBIL DATA DARI FORM =====
            lawan_ids = request.form.getlist('lawan_id[]')
            stasiun_lawan_list = request.form.getlist('stasiun_lawan[]')
            freq_tx_list = request.form.getlist('freq_tx[]')
            freq_rx_list = request.form.getlist('freq_rx[]')
            deleted_lawan_ids = request.form.getlist('deleted_lawan_ids[]')
            
            print(f"   Jumlah lawan yang ada: {len(station.stasiun_lawan_list)}")
            print(f"   Jumlah lawan dari form: {len(stasiun_lawan_list)}")
            print(f"   Jumlah lawan yang dihapus: {len(deleted_lawan_ids)}")
            
            # ===== VALIDASI MINIMAL 1 STASIUN LAWAN =====
            valid_lawan = [n for n in stasiun_lawan_list if n.strip()]
            if not valid_lawan:
                flash('❌ Minimal harus ada 1 stasiun lawan!', 'error')
                return redirect(url_for('admin_master_edit_stasiun', station_id=station_id, next=next_url))
            
            # ===== 3. HAPUS LAWAN YANG DITANDAI =====
            deleted_count = 0
            if deleted_lawan_ids:
                for lawan_id in deleted_lawan_ids:
                    if lawan_id:
                        lawan = StasiunLawan.query.get(int(lawan_id))
                        if lawan and lawan.stasiun_id == station_id:
                            print(f"   🗑️ Menghapus lawan: {lawan.nama_stasiun_lawan} (ID: {lawan.id})")
                            
                            # Putuskan relasi upload, jangan hapus file
                            uploads = UploadGambar.query.filter_by(stasiun_lawan_id=lawan.id).all()
                            for upload in uploads:
                                upload.stasiun_lawan_id = None
                            
                            # Hapus status updates
                            StatusUpdate.query.filter_by(stasiun_lawan_id=lawan.id).delete()
                            
                            changes.append(f"❌ HAPUS: '{lawan.nama_stasiun_lawan}' (status dihapus, gambar tetap ada)")
                            
                            db.session.delete(lawan)
                            deleted_count += 1
            
            # ===== 4. PROSES DATA YANG ADA (UPDATE) DAN DATA BARU (TAMBAH) =====
            updated_count = 0
            new_count = 0
            freq_tx_updated_count = 0
            freq_rx_updated_count = 0
            
            for i, nama_lawan in enumerate(stasiun_lawan_list):
                nama_lawan = nama_lawan.strip()
                if not nama_lawan:
                    continue
                
                freq_tx = freq_tx_list[i].strip() if i < len(freq_tx_list) else ''
                freq_rx = freq_rx_list[i].strip() if i < len(freq_rx_list) else ''
                
                # Cek apakah ini data existing (punya ID) atau baru
                if i < len(lawan_ids) and lawan_ids[i]:
                    # UPDATE DATA EXISTING
                    lawan_id = int(lawan_ids[i])
                    lawan = StasiunLawan.query.get(lawan_id)
                    
                    if lawan and lawan.stasiun_id == station_id:
                        # Cek perubahan nama
                        if lawan.nama_stasiun_lawan != nama_lawan:
                            changes.append(f"✏️ Ubah nama: '{lawan.nama_stasiun_lawan}' → '{nama_lawan}'")
                            lawan.nama_stasiun_lawan = nama_lawan
                            updated_count += 1
                        
                        # Cek perubahan frekuensi TX
                        old_freq_tx = lawan.freq_tx or ''
                        new_freq_tx = freq_tx if freq_tx else ''
                        if old_freq_tx != new_freq_tx:
                            if old_freq_tx and new_freq_tx:
                                changes.append(f"📡 Ubah Freq Tx '{lawan.nama_stasiun_lawan}': '{old_freq_tx}' → '{new_freq_tx}'")
                            elif new_freq_tx and not old_freq_tx:
                                changes.append(f"📡 Tambah Freq Tx '{lawan.nama_stasiun_lawan}': '{new_freq_tx}'")
                            elif old_freq_tx and not new_freq_tx:
                                changes.append(f"📡 Hapus Freq Tx '{lawan.nama_stasiun_lawan}' (sebelumnya '{old_freq_tx}')")
                            
                            lawan.freq_tx = new_freq_tx if new_freq_tx else None
                            freq_tx_updated_count += 1
                        
                        # Cek perubahan frekuensi RX
                        old_freq_rx = lawan.freq_rx or ''
                        new_freq_rx = freq_rx if freq_rx else ''
                        if old_freq_rx != new_freq_rx:
                            if old_freq_rx and new_freq_rx:
                                changes.append(f"📡 Ubah Freq Rx '{lawan.nama_stasiun_lawan}': '{old_freq_rx}' → '{new_freq_rx}'")
                            elif new_freq_rx and not old_freq_rx:
                                changes.append(f"📡 Tambah Freq Rx '{lawan.nama_stasiun_lawan}': '{new_freq_rx}'")
                            elif old_freq_rx and not new_freq_rx:
                                changes.append(f"📡 Hapus Freq Rx '{lawan.nama_stasiun_lawan}' (sebelumnya '{old_freq_rx}')")
                            
                            lawan.freq_rx = new_freq_rx if new_freq_rx else None
                            freq_rx_updated_count += 1
                else:
                    # DATA BARU
                    last_urutan = db.session.query(db.func.max(StasiunLawan.urutan))\
                        .filter_by(stasiun_id=station_id).scalar() or 0
                    
                    new_lawan = StasiunLawan(
                        stasiun_id=station_id,
                        nama_stasiun_lawan=nama_lawan,
                        freq_tx=freq_tx if freq_tx else None,
                        freq_rx=freq_rx if freq_rx else None,
                        group_id=None,
                        urutan=last_urutan + 1
                    )
                    db.session.add(new_lawan)
                    new_count += 1
                    freq_info = []
                    if freq_tx:
                        freq_info.append(f"Tx:{freq_tx}")
                    if freq_rx:
                        freq_info.append(f"Rx:{freq_rx}")
                    freq_str = f" ({', '.join(freq_info)})" if freq_info else ""
                    changes.append(f"✅ TAMBAH: '{nama_lawan}'{freq_str}")
            
            # ===== 5. COMMIT SEMUA PERUBAHAN =====
            db.session.commit()
            
            print(f"\n📊 SUMMARY:")
            print(f"   - Nama diubah: {updated_count}")
            print(f"   - Lawan baru: {new_count}")
            print(f"   - Freq Tx diubah: {freq_tx_updated_count}")
            print(f"   - Freq Rx diubah: {freq_rx_updated_count}")
            print(f"   - Lawan dihapus: {deleted_count}")
            print(f"{'='*60}\n")
            
            # ===== 6. TAMPILKAN RINGKASAN PERUBAHAN =====
            if changes:
                flash('📋 RINCIAN PERUBAHAN:', 'info')
                for change in changes:
                    flash(change, 'debug')
            
            # Ringkasan statistik
            summary = []
            if new_count > 0:
                summary.append(f"✅ {new_count} lawan baru")
            if updated_count > 0:
                summary.append(f"✏️ {updated_count} nama diubah")
            if freq_tx_updated_count > 0:
                summary.append(f"📡 {freq_tx_updated_count} Freq Tx diubah")
            if freq_rx_updated_count > 0:
                summary.append(f"📡 {freq_rx_updated_count} Freq Rx diubah")
            if deleted_count > 0:
                summary.append(f"❌ {deleted_count} lawan dihapus")
            
            if summary:
                flash(' ✓ '.join(summary), 'success')
            else:
                flash('ℹ️ Tidak ada perubahan data.', 'info')
            
            # ===== 7. REDIRECT KE NEXT_URL JIKA ADA =====
            if next_url:
                print(f"✅ Redirecting to next_url: {next_url}")
                return redirect(next_url)
            else:
                print(f"✅ Redirecting to detail page")
                return redirect(url_for('admin_master_detail_stasiun', station_id=station_id))
            
        except Exception as e:
            db.session.rollback()
            error_msg = str(e)
            flash(f'❌ ERROR: {error_msg}', 'error')
            print(f"\n❌ ERROR in admin_master_edit_stasiun: {error_msg}")
            import traceback
            traceback.print_exc()
            return redirect(url_for('admin_master_edit_stasiun', station_id=station_id, next=next_url))
    
    # ===== GET REQUEST - TAMPILKAN FORM EDIT =====
    return render_template('admin_master/edit_stasiun.html',
                         station=station,
                         operators=OPERATORS,
                         kota_list=kota_list,
                         next_url=next_url)

@app.route('/admin-master/hapus-stasiun/<int:station_id>')
@login_required
@admin_master_required
def admin_master_hapus_stasiun(station_id):
    station = Stasiun.query.get_or_404(station_id)
    
    try:
        # 1. HAPUS DULU SEMUA GRUP YANG TERKAIT DENGAN STASIUN INI
        groups = GrupStasiun.query.filter_by(stasiun_id=station_id).all()
        for group in groups:
            # Update lawan yang merujuk ke grup ini
            StasiunLawan.query.filter_by(grup_id=group.id).update({
                'group_id': None,
                'grup_id': None
            })
            
            # Update upload yang merujuk ke grup ini
            UploadGambar.query.filter_by(group_id=group.id).update({
                'group_id': None
            })
            
            # Hapus grup
            db.session.delete(group)
        
        # 2. AMBIL SEMUA LAWAN
        opponents = StasiunLawan.query.filter_by(stasiun_id=station_id).all()
        opponent_ids = [opponent.id for opponent in opponents]
        
        # 3. HAPUS STATUS UPDATES
        if opponent_ids:
            StatusUpdate.query.filter(StatusUpdate.stasiun_lawan_id.in_(opponent_ids)).delete()
        
        # 4. HAPUS UPLOAD GAMBAR (FILE FISIK DAN RECORD) - UPDATE UNTUK CLOUDINARY
        all_uploads = UploadGambar.query.filter_by(stasiun_id=station_id).all()
        for upload in all_uploads:
            # Hapus dari Cloudinary
            delete_from_cloudinary(upload.public_id)
            db.session.delete(upload)
        
        # 5. HAPUS STASIUN LAWAN
        StasiunLawan.query.filter_by(stasiun_id=station_id).delete()
        
        # 6. TERAKHIR, HAPUS STASIUN
        db.session.delete(station)
        db.session.commit()
        
        flash(f'Stasiun "{station.stasiun_name}" berhasil dihapus!', 'success')
    except Exception as e:
        db.session.rollback()
        flash(f'Error: {str(e)}', 'error')
        print(f"Error menghapus stasiun: {str(e)}")
        import traceback
        traceback.print_exc()
    
    return redirect(url_for('admin_master_daftar_stasiun'))

@app.route('/admin-master/upload-excel', methods=['GET', 'POST'])
@login_required
@admin_master_required
def admin_master_upload_excel():
    """Halaman upload Excel untuk admin master"""
    
    if request.method == 'POST':
        try:
            if 'excel_file' not in request.files:
                flash('Tidak ada file yang dipilih!', 'error')
                return redirect(url_for('admin_master_upload_excel'))
            
            file = request.files['excel_file']
            
            if file.filename == '':
                flash('Tidak ada file yang dipilih!', 'error')
                return redirect(url_for('admin_master_upload_excel'))
            
            # Validasi file
            is_valid, error_msg = validate_excel_file(file)
            if not is_valid:
                flash(error_msg, 'error')
                return redirect(url_for('admin_master_upload_excel'))
            
            # Get kota default
            kota_default = request.form.get('kota_default', 'samarinda')
            
            # Simpan file sementara
            temp_dir = os.path.join(app.config['UPLOAD_FOLDER'], 'temp')
            os.makedirs(temp_dir, exist_ok=True)
            
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            temp_filename = f"admin_master_{timestamp}_{secure_filename(file.filename)}"
            temp_path = os.path.join(temp_dir, temp_filename)
            
            file.save(temp_path)
            
            # Process file
            success_count, error_count, messages = process_excel_upload_admin_master(
                temp_path, 
                current_user,
                kota_default
            )
            
            # Hapus file temp
            try:
                os.remove(temp_path)
            except:
                pass
            
            # Tampilkan hasil
            for msg in messages:
                if 'Error' in msg or 'error' in msg:
                    flash(msg, 'error')
                elif 'sukses' in msg or 'berhasil' in msg:
                    flash(msg, 'success')
                else:
                    flash(msg, 'info')
            
            return redirect(url_for('admin_master_daftar_stasiun'))
            
        except Exception as e:
            flash(f'Error: {str(e)}', 'error')
            return redirect(url_for('admin_master_upload_excel'))
    
    # GET request
    total_stations = Stasiun.query.count()
    
    # Hitung total operator unik
    operators_count = db.session.query(Stasiun.operator).distinct().count()
    
    return render_template('admin_master/upload_excel.html',
                         total_stations=total_stations,
                         total_operators=operators_count,
                         operators=OPERATORS,
                         kota_list=KOTA_LIST)

@app.route('/admin-master/download-template')
@login_required
@admin_master_required
def admin_master_download_template():
    """Download template Excel untuk admin master"""
    return generate_admin_master_template()

@app.route('/debug-check-duplicates')
@login_required
@admin_master_required
def debug_check_duplicates():
    """Cek apakah ada file dengan nama yang sama - UPDATE untuk Cloudinary"""
    uploads = UploadGambar.query.all()
    
    # Kelompokkan berdasarkan original_filename
    filename_groups = {}
    for upload in uploads:
        original_name = upload.original_filename or 'unknown'
        
        if original_name not in filename_groups:
            filename_groups[original_name] = []
        filename_groups[original_name].append(upload)
    
    result = []
    result.append("=== CEK DUPLIKASI NAMA FILE ===")
    result.append("")
    
    for original_name, uploads_list in filename_groups.items():
        if len(uploads_list) > 1:
            result.append(f"⚠️ {original_name} - {len(uploads_list)} file:")
            for upload in uploads_list:
                station = Stasiun.query.get(upload.stasiun_id)
                result.append(f"   - {upload.public_id} (Stasiun: {station.stasiun_name if station else 'Unknown'})")
            result.append("")
    
    return "<pre>" + "\n".join(result) + "</pre>"

# ============================================================================
# ADMIN OPERATOR ROUTES
# ============================================================================

@app.route('/admin-operator/dashboard')
@login_required
@admin_operator_required
def admin_operator_dashboard():
    operator = current_user.operator_type
    
    total_stations = Stasiun.query.filter_by(operator=operator).count()
    
    total_opponents = StasiunLawan.query\
        .join(Stasiun)\
        .filter(Stasiun.operator == operator)\
        .count()
    
    stations_with_uploads = Stasiun.query\
        .filter(Stasiun.operator == operator, Stasiun.uploads.any())\
        .count()
    
    # BARU: Hitung stasiun yang sudah diperiksa dan belum diperiksa
    stations_checked = Stasiun.query\
        .filter(Stasiun.operator == operator, Stasiun.status_pemeriksaan == 'sudah_diperiksa')\
        .count()
    
    stations_unchecked = Stasiun.query\
        .filter(Stasiun.operator == operator, Stasiun.status_pemeriksaan == 'belum_diperiksa')\
        .count()
    
    recent_stations = Stasiun.query\
        .filter_by(operator=operator)\
        .order_by(Stasiun.created_at.desc())\
        .limit(5).all()
    
    # ===== PERBAIKAN: Ambil daftar kota case-insensitive =====
    from sqlalchemy import func
    
    kota_raw = db.session.query(
        func.lower(Stasiun.kota).label('kota_lower'),
        Stasiun.kota
    ).filter_by(operator=operator)\
     .distinct(func.lower(Stasiun.kota))\
     .order_by(func.lower(Stasiun.kota))\
     .all()
    
    # Mapping untuk menyimpan format standar
    kota_mapping = {}
    for kota_lower, kota_original in kota_raw:
        if kota_lower not in kota_mapping:
            # Cari format standar dari KOTA_LIST
            found = False
            for std_kota in KOTA_LIST:
                if std_kota.lower() == kota_lower:
                    kota_mapping[kota_lower] = std_kota
                    found = True
                    break
            if not found:
                kota_mapping[kota_lower] = kota_original.strip().title()
    
    kota_list_operator = list(kota_mapping.values())
    
    # Jika tidak ada kota di database, gunakan KOTA_LIST sebagai fallback
    if not kota_list_operator:
        kota_list_operator = KOTA_LIST
    
    current_date = datetime.now().strftime('%d %B %Y')
    
    return render_template('admin_operator/dashboard.html',
                         operator=operator,
                         operator_name=operator.upper(),
                         current_date=current_date,
                         total_stations=total_stations,
                         total_opponents=total_opponents,
                         stations_with_uploads=stations_with_uploads,
                         stations_checked=stations_checked,
                         stations_unchecked=stations_unchecked,
                         recent_stations=recent_stations,
                         kota_list=kota_list_operator)

@app.route('/admin-operator/tambah-data', methods=['GET', 'POST'])
@login_required
@admin_operator_required
def admin_operator_tambah_data():
    operator = current_user.operator_type
    
    if request.method == 'POST':
        try:
            stasiun_name = request.form.get('stasiun_name').strip()
            kota = request.form.get('kota')
            
            stasiun_lawan_list = request.form.getlist('stasiun_lawan[]')
            freq_tx_list = request.form.getlist('freq_tx[]')
            freq_rx_list = request.form.getlist('freq_rx[]')
            
            if not stasiun_name:
                flash('Nama stasiun harus diisi!', 'error')
                return redirect(url_for('admin_operator_tambah_data'))
            
            # Validasi minimal 1 stasiun lawan
            valid_lawan = [n for n in stasiun_lawan_list if n.strip()]
            if not valid_lawan:
                flash('Minimal harus ada 1 stasiun lawan!', 'error')
                return redirect(url_for('admin_operator_tambah_data'))
            
            new_station = Stasiun(
                stasiun_name=stasiun_name,
                operator=operator,
                kota=kota,
                created_by=current_user.id
            )
            db.session.add(new_station)
            db.session.flush()
            
            for i, nama_lawan in enumerate(stasiun_lawan_list):
                nama_lawan = nama_lawan.strip()
                if nama_lawan:
                    freq_tx = freq_tx_list[i].strip() if i < len(freq_tx_list) else ''
                    freq_rx = freq_rx_list[i].strip() if i < len(freq_rx_list) else ''
                    
                    new_lawan = StasiunLawan(
                        stasiun_id=new_station.id,
                        nama_stasiun_lawan=nama_lawan,
                        freq_tx=freq_tx if freq_tx else None,
                        freq_rx=freq_rx if freq_rx else None,
                        group_id=None,
                        urutan=i
                    )
                    db.session.add(new_lawan)
            
            db.session.commit()
            flash(f'Stasiun "{stasiun_name}" berhasil ditambahkan!', 'success')
            return redirect(url_for('admin_operator_daftar_stasiun'))
            
        except Exception as e:
            db.session.rollback()
            flash(f'Error: {str(e)}', 'error')
            return redirect(url_for('admin_operator_tambah_data'))
    
    return render_template('admin_operator/tambah_data.html',
                         kota_list=KOTA_LIST,
                         operator=operator)

def _handle_admin_operator_daftar_stasiun(template_name='admin_operator/daftar_stasiun.html'):
    operator = current_user.operator_type
    kota = request.args.get('kota', 'all')
    search_stasiun = request.args.get('search_stasiun', '').strip()
    status_pemeriksaan = request.args.get('status_pemeriksaan', 'all')
    status_update_filter = request.args.get('status_update_filter', 'all')
    
    # ===== PARAMETER PAGINATION =====
    per_page = request.args.get('per_page', 10, type=int)
    if per_page not in [10, 20, 50, 100]:
        per_page = 10
    page = request.args.get('page', 1, type=int)
    
    # ===== QUERY DASAR DENGAN PAGINATION =====
    query = Stasiun.query.filter_by(operator=operator)
    
    if search_stasiun:
        query = query.filter(Stasiun.stasiun_name.ilike(f'%{search_stasiun}%'))
    
    if kota != 'all' and kota:
        query = query.filter(Stasiun.kota.ilike(kota))
    
    if status_pemeriksaan != 'all':
        query = query.filter(Stasiun.status_pemeriksaan == status_pemeriksaan)
    
    # ===== PAGINATION DI DATABASE LEVEL =====
    pagination = query.order_by(Stasiun.created_at.desc()).paginate(
        page=page, per_page=per_page, error_out=False
    )
    stations = pagination.items
    
    if not stations:
        # Handle empty case
        from sqlalchemy import func
        kota_raw = db.session.query(
            func.lower(Stasiun.kota).label('kota_lower'),
            Stasiun.kota
        ).filter_by(operator=operator).distinct(func.lower(Stasiun.kota)).order_by(func.lower(Stasiun.kota)).all()
        
        kota_mapping = {}
        for kota_lower, kota_original in kota_raw:
            if kota_lower not in kota_mapping:
                found = False
                for std_kota in KOTA_LIST:
                    if std_kota.lower() == kota_lower:
                        kota_mapping[kota_lower] = std_kota
                        found = True
                        break
                if not found:
                    kota_mapping[kota_lower] = kota_original.strip().title()
        
        kota_list_operator = list(kota_mapping.values())
        if not kota_list_operator:
            kota_list_operator = KOTA_LIST
        
        return render_template(template_name,
                             stations=[],
                             pagination=pagination,
                             operator=operator,
                             operator_name=operator.upper(),
                             kota_list=kota_list_operator,
                             selected_kota=kota,
                             selected_status=status_pemeriksaan,
                             selected_status_update=status_update_filter,
                             status_summary={'lengkap': 0, 'belum_lengkap': 0},
                             status_options=STATUS_OPTIONS)
    
    # ===== BATCH QUERY UNTUK SEMUA STASIUN DALAM 1 HALAMAN =====
    station_ids = [s.id for s in stations]
    
    # 1. Batch query semua stasiun lawan (1 query)
    all_opponents = StasiunLawan.query\
        .filter(StasiunLawan.stasiun_id.in_(station_ids))\
        .order_by(StasiunLawan.urutan)\
        .all()
    
    # 2. Batch query semua status terbaru (1 query)
    opponent_ids = [o.id for o in all_opponents]
    latest_status_dict = {}
    
    if opponent_ids:
        from sqlalchemy import func
        subquery = db.session.query(
            StatusUpdate.stasiun_lawan_id,
            func.max(StatusUpdate.updated_at).label('max_updated')
        ).filter(StatusUpdate.stasiun_lawan_id.in_(opponent_ids))\
         .group_by(StatusUpdate.stasiun_lawan_id)\
         .subquery()
        
        latest_statuses = db.session.query(StatusUpdate)\
            .join(subquery,
                  (StatusUpdate.stasiun_lawan_id == subquery.c.stasiun_lawan_id) &
                  (StatusUpdate.updated_at == subquery.c.max_updated))\
            .all()
        
        for status in latest_statuses:
            latest_status_dict[status.stasiun_lawan_id] = status
    
    # 3. Batch query semua grup (1 query)
    all_groups = GrupStasiun.query.filter(GrupStasiun.stasiun_id.in_(station_ids)).all()
    groups_by_station = {}
    for group in all_groups:
        if group.stasiun_id not in groups_by_station:
            groups_by_station[group.stasiun_id] = []
        groups_by_station[group.stasiun_id].append(group)
    
    # 4. Batch query user untuk checked_by (1 query)
    user_ids = [s.checked_by for s in stations if s.checked_by]
    users_dict = {}
    if user_ids:
        users = User.query.filter(User.id.in_(user_ids)).all()
        users_dict = {u.id: u for u in users}
    
    # ===== Kelompokkan lawan per stasiun =====
    opponents_by_station = {}
    for opponent in all_opponents:
        if opponent.stasiun_id not in opponents_by_station:
            opponents_by_station[opponent.stasiun_id] = []
        opponents_by_station[opponent.stasiun_id].append(opponent)
    
    # ===== Proses setiap stasiun =====
    for station in stations:
        station.stasiun_lawan_list = opponents_by_station.get(station.id, [])
        station.grup_list = groups_by_station.get(station.id, [])
        station.checked_by_user = users_dict.get(station.checked_by)
        
        # Hitung status untuk setiap lawan
        total_lawan = len(station.stasiun_lawan_list)
        lawan_tanpa_status = 0
        station.status_summary = {'aktif': 0, 'tidak_aktif': 0, 'tidak_sesuai': 0}
        
        for lawan in station.stasiun_lawan_list:
            latest_status = latest_status_dict.get(lawan.id)
            lawan.latest_status = latest_status
            
            if latest_status:
                if latest_status.status in station.status_summary:
                    station.status_summary[latest_status.status] += 1
            else:
                lawan_tanpa_status += 1
        
        station.has_lawan_without_status = lawan_tanpa_status > 0
        
        # ===== Build groups display =====
        groups_dict = {g.id: g.nama_grup for g in station.grup_list}
        
        lawan_by_group = {}
        group_stats = {}
        
        for lawan in station.stasiun_lawan_list:
            group_key = lawan.group_id if lawan.group_id is not None else 'ungrouped'
            if group_key not in lawan_by_group:
                lawan_by_group[group_key] = []
                group_stats[group_key] = {
                    'total': 0, 'aktif': 0, 'tidak_aktif': 0, 'tidak_sesuai': 0, 'belum_ada': 0
                }
            
            lawan_by_group[group_key].append(lawan)
            group_stats[group_key]['total'] += 1
            
            if lawan.latest_status:
                status = lawan.latest_status.status
                if status in group_stats[group_key]:
                    group_stats[group_key][status] += 1
            else:
                group_stats[group_key]['belum_ada'] += 1
        
        groups_display = []
        
        if 'ungrouped' in lawan_by_group:
            groups_display.append({
                'group_id': None,
                'group_name': 'Tanpa Grup',
                'lawans': lawan_by_group['ungrouped'],
                'stats': group_stats['ungrouped'],
                'is_ungrouped': True,
                'badge_color': 'warning'
            })
        
        for group_key, lawans in lawan_by_group.items():
            if group_key != 'ungrouped':
                try:
                    group_id = int(group_key)
                    group_name = groups_dict.get(group_id, f'Grup {group_id}')
                    groups_display.append({
                        'group_id': group_id,
                        'group_name': group_name,
                        'lawans': lawans,
                        'stats': group_stats[group_key],
                        'is_ungrouped': False,
                        'badge_color': 'success'
                    })
                except:
                    continue
        
        groups_display.sort(key=lambda x: (x['is_ungrouped'], x['group_id'] if x['group_id'] is not None else 0))
        station.groups_display = groups_display
    
    # ===== Filter berdasarkan status_update_filter =====
    if status_update_filter != 'all':
        filtered_stations = []
        for station in stations:
            total_lawan = len(station.stasiun_lawan_list)
            lawan_tanpa_status = sum(1 for l in station.stasiun_lawan_list if not l.latest_status)
            is_lengkap = (lawan_tanpa_status == 0 and total_lawan > 0)
            
            if status_update_filter == 'lengkap' and is_lengkap:
                filtered_stations.append(station)
            elif status_update_filter == 'belum_lengkap' and not is_lengkap and total_lawan > 0:
                filtered_stations.append(station)
        
        stations = filtered_stations
        pagination.items = stations
        pagination.total = len(filtered_stations)
    
    # ===== Hitung status summary untuk badge =====
    status_summary = {'lengkap': 0, 'belum_lengkap': 0}
    for station in stations:
        total_lawan = len(station.stasiun_lawan_list)
        lawan_tanpa_status = sum(1 for l in station.stasiun_lawan_list if not l.latest_status)
        if lawan_tanpa_status == 0 and total_lawan > 0:
            status_summary['lengkap'] += 1
        elif total_lawan > 0:
            status_summary['belum_lengkap'] += 1
    
    # ===== Ambil daftar kota untuk filter =====
    from sqlalchemy import func
    
    kota_raw = db.session.query(
        func.lower(Stasiun.kota).label('kota_lower'),
        Stasiun.kota
    ).filter_by(operator=operator)\
     .distinct(func.lower(Stasiun.kota))\
     .order_by(func.lower(Stasiun.kota))\
     .all()
    
    kota_mapping = {}
    for kota_lower, kota_original in kota_raw:
        if kota_lower not in kota_mapping:
            found = False
            for std_kota in KOTA_LIST:
                if std_kota.lower() == kota_lower:
                    kota_mapping[kota_lower] = std_kota
                    found = True
                    break
            if not found:
                kota_mapping[kota_lower] = kota_original.strip().title()
    
    kota_list_operator = list(kota_mapping.values())
    if not kota_list_operator:
        kota_list_operator = KOTA_LIST
    
    selected_kota_normalized = kota
    if kota != 'all':
        found = False
        for k in kota_list_operator:
            if k.lower() == kota.lower():
                selected_kota_normalized = k
                found = True
                break
        if not found:
            selected_kota_normalized = 'all'
    
    return render_template(template_name,
                         stations=stations,
                         pagination=pagination,
                         operator=operator,
                         operator_name=operator.upper(),
                         kota_list=kota_list_operator,
                         selected_kota=selected_kota_normalized,
                         selected_status=status_pemeriksaan,
                         selected_status_update=status_update_filter,
                         status_summary=status_summary,
                         status_options=STATUS_OPTIONS)

@app.route('/admin-operator/daftar-stasiun')
@login_required
@admin_operator_required
def admin_operator_daftar_stasiun():
    return _handle_admin_operator_daftar_stasiun('admin_operator/daftar_stasiun.html')

@app.route('/admin-operator/daftar-stasiun-v2')
@login_required
@admin_operator_required
def admin_operator_daftar_stasiun_v2():
    return _handle_admin_operator_daftar_stasiun('admin_operator/daftar_stasiun_v2.html')
                         
@app.route('/admin-operator/stasiun/<int:station_id>')
@login_required
@admin_operator_required
def admin_operator_detail_stasiun(station_id):
    """Detail stasiun untuk Admin Operator dengan tampilan per grup"""
    station = Stasiun.query.get_or_404(station_id)
    
    # Validasi kepemilikan stasiun
    if station.operator != current_user.operator_type:
        flash('Stasiun ini bukan milik operator Anda!', 'error')
        return redirect(url_for('admin_operator_dashboard'))
    
    # ===== AMBIL PARAMETER NEXT =====
    next_url = request.args.get('next', '')
    
    # ===== AMBIL INFORMASI PEMERIKSA =====
    station.checked_by_user = None
    if station.checked_by:
        station.checked_by_user = User.query.get(station.checked_by)
    
    # ===== AMBIL SEMUA LAWAN =====
    all_lawans = StasiunLawan.query\
        .filter_by(stasiun_id=station_id)\
        .order_by(StasiunLawan.urutan)\
        .all()
    
    # ===== AMBIL SEMUA GRUP =====
    all_groups = GrupStasiun.query.filter_by(stasiun_id=station_id).all()
    groups_dict = {g.id: g.nama_grup for g in all_groups}
    
    # ===== AMBIL SEMUA UPLOAD UNIK (group by public_id) =====
    uploads_query = db.session.query(
        UploadGambar.group_id,
        UploadGambar.public_id,
        db.func.max(UploadGambar.id).label('latest_id')
    ).filter_by(stasiun_id=station_id)\
     .group_by(UploadGambar.group_id, UploadGambar.public_id)\
     .subquery()
    
    all_uploads = UploadGambar.query\
        .join(uploads_query, UploadGambar.id == uploads_query.c.latest_id)\
        .order_by(UploadGambar.uploaded_at.desc())\
        .all()
    
    # ===== KELOMPOKKAN DATA PER GRUP =====
    groups = {}
    
    # ===== PROSES LAWAN PER GRUP =====
    for lawan in all_lawans:
        group_key = lawan.group_id if lawan.group_id is not None else 'ungrouped'
        
        if group_key not in groups:
            groups[group_key] = {
                'group_id': lawan.group_id,
                'nama_grup': groups_dict.get(lawan.group_id) if lawan.group_id else None,
                'opponents': [],
                'uploads': [],
                'status_summary': {
                    'aktif': 0,
                    'tidak_aktif': 0,
                    'tidak_sesuai': 0
                }
            }
        
        # Ambil status terbaru
        latest_status = StatusUpdate.query\
            .filter_by(stasiun_lawan_id=lawan.id)\
            .order_by(StatusUpdate.updated_at.desc())\
            .first()
        
        # Ambil riwayat status (5 terakhir)
        status_history = StatusUpdate.query\
            .filter_by(stasiun_lawan_id=lawan.id)\
            .order_by(StatusUpdate.updated_at.desc())\
            .limit(5)\
            .all()
        
        # Ambil informasi updater
        updater_info = None
        if latest_status and latest_status.updated_by:
            updater = db.session.get(User, latest_status.updated_by)
            updater_info = updater.username if updater else None
        
        lawan_data = build_stasiun_lawan_dict(
            lawan, station, latest_status, status_history, updater_info=updater_info
        )
        
        groups[group_key]['opponents'].append(lawan_data)
        
        # Hitung statistik status per grup
        if latest_status and latest_status.status in groups[group_key]['status_summary']:
            groups[group_key]['status_summary'][latest_status.status] += 1
    
    # ===== PROSES UPLOAD PER GRUP =====
    for upload in all_uploads:
        group_key = upload.group_id if upload.group_id is not None else 'ungrouped'
        
        if group_key not in groups:
            groups[group_key] = {
                'group_id': upload.group_id,
                'nama_grup': groups_dict.get(upload.group_id) if upload.group_id else None,
                'opponents': [],
                'uploads': [],
                'status_summary': {
                    'aktif': 0,
                    'tidak_aktif': 0,
                    'tidak_sesuai': 0
                }
            }
        
        # Cek duplikasi public_id dalam grup yang sama
        existing_public_ids = [u.get('public_id') for u in groups[group_key]['uploads']]
        if upload.public_id not in existing_public_ids:
            # Dapatkan informasi uploader
            uploader = db.session.get(User, upload.uploaded_by) if upload.uploaded_by else None
            
            # Dapatkan thumbnail URL
            thumbnail_url = None
            if upload.public_id:
                thumbnail_url, _ = cloudinary_url(
                    upload.public_id, 
                    width=200, 
                    height=200, 
                    crop="fill", 
                    quality="auto"
                )
            
            upload_data = {
                'id': upload.id,
                'public_id': upload.public_id,
                'cloudinary_url': upload.cloudinary_url,
                'thumbnail_url': thumbnail_url,
                'original_filename': upload.original_filename,
                'uploaded_at': upload.uploaded_at,
                'uploaded_by': uploader.username if uploader else 'System',
                'uploaded_by_id': upload.uploaded_by,
                'status': upload.status,
                'format': upload.format,
                'width': upload.width,
                'height': upload.height,
                'bytes_size': upload.bytes_size,
                'size_mb': round(upload.bytes_size / (1024 * 1024), 2) if upload.bytes_size else None
            }
            groups[group_key]['uploads'].append(upload_data)
    
    # ===== URUTKAN GRUP =====
    sorted_groups = {}
    
    # Ungrouped dulu
    if 'ungrouped' in groups:
        sorted_groups['ungrouped'] = groups.pop('ungrouped')
    
    # Grup numeric diurutkan
    numeric_groups = {}
    for key, data in groups.items():
        if key != 'ungrouped':
            try:
                numeric_groups[int(key)] = data
            except:
                numeric_groups[key] = data
    
    for key in sorted(numeric_groups.keys()):
        sorted_groups[str(key)] = numeric_groups[key]
    
    # ===== HITUNG STATISTIK TOTAL =====
    status_counts = {'aktif': 0, 'tidak_aktif': 0, 'tidak_sesuai': 0}
    for lawan in all_lawans:
        latest_status = StatusUpdate.query\
            .filter_by(stasiun_lawan_id=lawan.id)\
            .order_by(StatusUpdate.updated_at.desc())\
            .first()
        if latest_status and latest_status.status in status_counts:
            status_counts[latest_status.status] += 1
    
    # ===== HITUNG STATISTIK UPLOAD =====
    total_uploads = len(all_uploads)
    total_upload_size = sum([u.bytes_size for u in all_uploads if u.bytes_size]) if all_uploads else 0
    total_upload_size_mb = round(total_upload_size / (1024 * 1024), 2) if total_upload_size else 0
    
    # ===== STATISTIK PER BULAN (untuk grafik) =====
    uploads_by_month = {}
    status_updates_by_month = {}
    
    for upload in all_uploads:
        month_key = upload.uploaded_at.strftime('%Y-%m')
        uploads_by_month[month_key] = uploads_by_month.get(month_key, 0) + 1
    
    for lawan in all_lawans:
        status_updates = StatusUpdate.query\
            .filter_by(stasiun_lawan_id=lawan.id)\
            .order_by(StatusUpdate.updated_at)\
            .all()
        
        for status_update in status_updates:
            month_key = status_update.updated_at.strftime('%Y-%m')
            if month_key not in status_updates_by_month:
                status_updates_by_month[month_key] = {
                    'aktif': 0,
                    'tidak_aktif': 0,
                    'tidak_sesuai': 0
                }
            if status_update.status in status_updates_by_month[month_key]:
                status_updates_by_month[month_key][status_update.status] += 1
    
    # ===== HITUNG TOTAL OPPONENTS DENGAN STATUS =====
    total_opponents = len(all_lawans)
    opponents_with_status = sum(1 for lawan in all_lawans if StatusUpdate.query.filter_by(stasiun_lawan_id=lawan.id).first())
    opponents_without_status = total_opponents - opponents_with_status
    
    # ===== AMBIL UPLOAD TERBARU UNTUK PREVIEW =====
    recent_uploads = sorted(all_uploads, key=lambda x: x.uploaded_at, reverse=True)[:5]
    
    # ===== TAMBAHKAN station.grup_list UNTUK TEMPLATE =====
    station.grup_list = all_groups
    
    # ===== HITUNG STATISTIK GRUP =====
    total_groups = len(sorted_groups)
    groups_with_uploads = sum(1 for g in sorted_groups.values() if g['uploads'])
    groups_without_uploads = total_groups - groups_with_uploads
    
    return render_template('admin_operator/detail_stasiun.html',
                         station=station,
                         groups=sorted_groups,
                         all_uploads=all_uploads,
                         recent_uploads=recent_uploads,
                         total_opponents=total_opponents,
                         status_counts=status_counts,
                         uploads_by_month=uploads_by_month,
                         status_updates_by_month=status_updates_by_month,
                         total_uploads=total_uploads,
                         total_upload_size_mb=total_upload_size_mb,
                         opponents_with_status=opponents_with_status,
                         opponents_without_status=opponents_without_status,
                         checked_by_user=station.checked_by_user,
                         total_groups=total_groups,
                         groups_with_uploads=groups_with_uploads,
                         groups_without_uploads=groups_without_uploads,
                         next_url=next_url)

@app.route('/admin-operator/update-status-pemeriksaan/<int:station_id>', methods=['POST'])
@login_required
@admin_operator_required
def admin_operator_update_status_pemeriksaan(station_id):
    """Update status pemeriksaan stasiun untuk Admin Operator"""
    try:
        station = Stasiun.query.get_or_404(station_id)
        
        # Validasi kepemilikan
        if station.operator != current_user.operator_type:
            return jsonify({'success': False, 'message': 'Stasiun ini bukan milik operator Anda!'})
        
        data = request.get_json()
        status = data.get('status')
        
        if status not in ['sudah_diperiksa', 'belum_diperiksa']:
            return jsonify({'success': False, 'message': 'Status tidak valid'})
        
        # Update status
        old_status = station.status_pemeriksaan
        station.status_pemeriksaan = status
        
        if status == 'sudah_diperiksa':
            station.checked_at = datetime.utcnow()
            station.checked_by = current_user.id
            print(f"✅ Status diubah: {old_status} → {status} oleh {current_user.username}")
        else:
            station.checked_at = None
            station.checked_by = None
            print(f"✅ Status diubah: {old_status} → {status} (reset)")
        
        db.session.commit()
        
        # Ambil informasi checker untuk response
        checker_name = current_user.username if status == 'sudah_diperiksa' else None
        checked_at_str = station.checked_at.strftime('%d/%m/%Y') if station.checked_at else None
        
        return jsonify({
            'success': True,
            'message': f'Status pemeriksaan berhasil diupdate menjadi "{status}"',
            'checked_at': checked_at_str,
            'checked_by': checker_name
        })
        
    except Exception as e:
        db.session.rollback()
        print(f"❌ ERROR: {str(e)}")
        return jsonify({'success': False, 'message': str(e)})                        

@app.route('/admin-operator/edit-stasiun/<int:station_id>', methods=['GET', 'POST'])
@login_required
@admin_operator_required
def admin_operator_edit_stasiun(station_id):
    station = Stasiun.query.get_or_404(station_id)
    
    if station.operator != current_user.operator_type:
        flash('❌ Akses ditolak! Stasiun ini bukan milik operator Anda.', 'error')
        return redirect(url_for('admin_operator_dashboard'))
    
    # ===== AMBIL PARAMETER NEXT =====
    next_url = request.args.get('next', '')
    
    # ===== DEBUG: TAMPILKAN NILAI KOTA =====
    print(f"\n{'='*60}")
    print(f"📝 EDIT STASIUN (ADMIN OPERATOR) - ID: {station_id}")
    print(f"   Nama Stasiun: {station.stasiun_name}")
    print(f"   Operator: {station.operator}")
    print(f"   Kota di Database: '{station.kota}'")
    print(f"   Next URL: {next_url}")
    print(f"{'='*60}\n")
    
    # ===== AMBIL DAFTAR KOTA DARI DATABASE UNTUK OPERATOR INI =====
    kota_list_from_db = db.session.query(Stasiun.kota)\
        .filter_by(operator=station.operator)\
        .distinct()\
        .order_by(Stasiun.kota)\
        .all()
    kota_list = [k[0] for k in kota_list_from_db if k[0]]
    
    # Jika tidak ada kota di database, gunakan KOTA_LIST sebagai fallback
    if not kota_list:
        kota_list = KOTA_LIST
        print(f"⚠️ Tidak ada kota untuk operator {station.operator}, menggunakan KOTA_LIST default")
    
    # ===== PASTIKAN station.kota ADA DI DALAM kota_list =====
    if station.kota and station.kota not in kota_list:
        print(f"⚠️ Peringatan: Kota '{station.kota}' tidak ada di kota_list!")
        print(f"   Menambahkan kota '{station.kota}' ke dalam dropdown")
        kota_list.append(station.kota)
        kota_list.sort()
    
    print(f"📋 Daftar kota yang tersedia: {kota_list}")
    print(f"📍 Kota yang dipilih: '{station.kota}'")
    print(f"{'='*60}\n")
    
    if request.method == 'POST':
        try:
            # ===== VALIDASI INPUT =====
            stasiun_name = request.form.get('stasiun_name', '').strip()
            kota = request.form.get('kota', '').strip()
            
            print(f"\n📝 POST DATA:")
            print(f"   Nama: {stasiun_name}")
            print(f"   Kota: '{kota}'")
            print(f"{'-'*40}")
            
            if not stasiun_name:
                flash('❌ Nama stasiun harus diisi!', 'error')
                return redirect(url_for('admin_operator_edit_stasiun', station_id=station_id))
            
            # ===== VALIDASI KOTA - PASTIKAN TIDAK KOSONG =====
            if not kota:
                flash('❌ Kota/Kabupaten harus dipilih!', 'error')
                return redirect(url_for('admin_operator_edit_stasiun', station_id=station_id))
            
            # ===== NORMALISASI KOTA =====
            kota_normalized = None
            for k in KOTA_LIST:
                if k.lower() == kota.lower():
                    kota_normalized = k
                    break
            
            # Coba cari yang mirip
            if not kota_normalized:
                for k in KOTA_LIST:
                    if kota.lower() in k.lower() or k.lower() in kota.lower():
                        kota_normalized = k
                        break
            
            if not kota_normalized:
                # Jika tidak ditemukan, gunakan nilai asli dengan peringatan
                kota_normalized = kota
                flash(f'⚠️ Kota "{kota}" tidak ada dalam daftar standar, menggunakan nilai asli', 'warning')
            
            # Catat perubahan untuk pesan
            changes = []
            
            # ===== 1. UPDATE DATA STASIUN UTAMA =====
            if station.stasiun_name != stasiun_name:
                changes.append(f"📝 Nama stasiun: '{station.stasiun_name}' → '{stasiun_name}'")
                station.stasiun_name = stasiun_name
            
            if station.kota != kota_normalized:
                changes.append(f"📍 Kota/Kab: '{station.kota}' → '{kota_normalized}'")
                station.kota = kota_normalized
            
            # ===== 2. AMBIL DATA DARI FORM =====
            lawan_ids = request.form.getlist('lawan_id[]')
            stasiun_lawan_list = request.form.getlist('stasiun_lawan[]')
            freq_tx_list = request.form.getlist('freq_tx[]')
            freq_rx_list = request.form.getlist('freq_rx[]')
            deleted_lawan_ids = request.form.getlist('deleted_lawan_ids[]')
            
            print(f"   Jumlah lawan yang ada: {len(station.stasiun_lawan_list)}")
            print(f"   Jumlah lawan dari form: {len(stasiun_lawan_list)}")
            print(f"   Jumlah lawan yang dihapus: {len(deleted_lawan_ids)}")
            
            # ===== VALIDASI MINIMAL 1 STASIUN LAWAN =====
            valid_lawan = [n for n in stasiun_lawan_list if n.strip()]
            if not valid_lawan:
                flash('❌ Minimal harus ada 1 stasiun lawan!', 'error')
                return redirect(url_for('admin_operator_edit_stasiun', station_id=station_id))
            
            # ===== 3. HAPUS LAWAN YANG DITANDAI =====
            deleted_count = 0
            if deleted_lawan_ids:
                for lawan_id in deleted_lawan_ids:
                    if lawan_id:
                        lawan = StasiunLawan.query.get(int(lawan_id))
                        if lawan and lawan.stasiun_id == station_id:
                            print(f"   🗑️ Menghapus lawan: {lawan.nama_stasiun_lawan} (ID: {lawan.id})")
                            
                            # Putuskan relasi upload, jangan hapus file
                            uploads = UploadGambar.query.filter_by(stasiun_lawan_id=lawan.id).all()
                            for upload in uploads:
                                upload.stasiun_lawan_id = None
                            
                            # Hapus status updates
                            StatusUpdate.query.filter_by(stasiun_lawan_id=lawan.id).delete()
                            
                            changes.append(f"❌ HAPUS: '{lawan.nama_stasiun_lawan}' (status dihapus, gambar tetap ada)")
                            
                            db.session.delete(lawan)
                            deleted_count += 1
            
            # ===== 4. PROSES DATA YANG ADA (UPDATE) DAN DATA BARU (TAMBAH) =====
            updated_count = 0
            new_count = 0
            freq_tx_updated_count = 0
            freq_rx_updated_count = 0
            
            for i, nama_lawan in enumerate(stasiun_lawan_list):
                nama_lawan = nama_lawan.strip()
                if not nama_lawan:
                    continue
                
                freq_tx = freq_tx_list[i].strip() if i < len(freq_tx_list) else ''
                freq_rx = freq_rx_list[i].strip() if i < len(freq_rx_list) else ''
                
                # Cek apakah ini data existing (punya ID) atau baru
                if i < len(lawan_ids) and lawan_ids[i]:
                    # UPDATE DATA EXISTING
                    lawan_id = int(lawan_ids[i])
                    lawan = StasiunLawan.query.get(lawan_id)
                    
                    if lawan and lawan.stasiun_id == station_id:
                        # Cek perubahan nama
                        if lawan.nama_stasiun_lawan != nama_lawan:
                            changes.append(f"✏️ Ubah nama: '{lawan.nama_stasiun_lawan}' → '{nama_lawan}'")
                            lawan.nama_stasiun_lawan = nama_lawan
                            updated_count += 1
                        
                        # Cek perubahan frekuensi TX
                        old_freq_tx = lawan.freq_tx or ''
                        new_freq_tx = freq_tx if freq_tx else ''
                        if old_freq_tx != new_freq_tx:
                            if old_freq_tx and new_freq_tx:
                                changes.append(f"📡 Ubah Freq Tx '{lawan.nama_stasiun_lawan}': '{old_freq_tx}' → '{new_freq_tx}'")
                            elif new_freq_tx and not old_freq_tx:
                                changes.append(f"📡 Tambah Freq Tx '{lawan.nama_stasiun_lawan}': '{new_freq_tx}'")
                            elif old_freq_tx and not new_freq_tx:
                                changes.append(f"📡 Hapus Freq Tx '{lawan.nama_stasiun_lawan}' (sebelumnya '{old_freq_tx}')")
                            
                            lawan.freq_tx = new_freq_tx if new_freq_tx else None
                            freq_tx_updated_count += 1
                        
                        # Cek perubahan frekuensi RX
                        old_freq_rx = lawan.freq_rx or ''
                        new_freq_rx = freq_rx if freq_rx else ''
                        if old_freq_rx != new_freq_rx:
                            if old_freq_rx and new_freq_rx:
                                changes.append(f"📡 Ubah Freq Rx '{lawan.nama_stasiun_lawan}': '{old_freq_rx}' → '{new_freq_rx}'")
                            elif new_freq_rx and not old_freq_rx:
                                changes.append(f"📡 Tambah Freq Rx '{lawan.nama_stasiun_lawan}': '{new_freq_rx}'")
                            elif old_freq_rx and not new_freq_rx:
                                changes.append(f"📡 Hapus Freq Rx '{lawan.nama_stasiun_lawan}' (sebelumnya '{old_freq_rx}')")
                            
                            lawan.freq_rx = new_freq_rx if new_freq_rx else None
                            freq_rx_updated_count += 1
                else:
                    # DATA BARU
                    last_urutan = db.session.query(db.func.max(StasiunLawan.urutan))\
                        .filter_by(stasiun_id=station_id).scalar() or 0
                    
                    new_lawan = StasiunLawan(
                        stasiun_id=station_id,
                        nama_stasiun_lawan=nama_lawan,
                        freq_tx=freq_tx if freq_tx else None,
                        freq_rx=freq_rx if freq_rx else None,
                        group_id=None,
                        urutan=last_urutan + 1
                    )
                    db.session.add(new_lawan)
                    new_count += 1
                    freq_info = []
                    if freq_tx:
                        freq_info.append(f"Tx:{freq_tx}")
                    if freq_rx:
                        freq_info.append(f"Rx:{freq_rx}")
                    freq_str = f" ({', '.join(freq_info)})" if freq_info else ""
                    changes.append(f"✅ TAMBAH: '{nama_lawan}'{freq_str}")
            
            # ===== 5. COMMIT SEMUA PERUBAHAN =====
            db.session.commit()
            
            print(f"\n📊 SUMMARY:")
            print(f"   - Nama diubah: {updated_count}")
            print(f"   - Lawan baru: {new_count}")
            print(f"   - Freq Tx diubah: {freq_tx_updated_count}")
            print(f"   - Freq Rx diubah: {freq_rx_updated_count}")
            print(f"   - Lawan dihapus: {deleted_count}")
            print(f"{'='*60}\n")
            
            # ===== 6. TAMPILKAN RINGKASAN PERUBAHAN =====
            if changes:
                flash('📋 RINCIAN PERUBAHAN:', 'info')
                for change in changes:
                    flash(change, 'debug')
            
            # Ringkasan statistik
            summary = []
            if new_count > 0:
                summary.append(f"✅ {new_count} lawan baru")
            if updated_count > 0:
                summary.append(f"✏️ {updated_count} nama diubah")
            if freq_tx_updated_count > 0:
                summary.append(f"📡 {freq_tx_updated_count} Freq Tx diubah")
            if freq_rx_updated_count > 0:
                summary.append(f"📡 {freq_rx_updated_count} Freq Rx diubah")
            if deleted_count > 0:
                summary.append(f"❌ {deleted_count} lawan dihapus")
            
            if summary:
                flash(' ✓ '.join(summary), 'success')
            else:
                flash('ℹ️ Tidak ada perubahan data.', 'info')
            
            # ===== REDIRECT KE NEXT_URL JIKA ADA =====
            if next_url:
                print(f"✅ Redirecting to next_url: {next_url}")
                return redirect(next_url)
            else:
                print(f"✅ Redirecting to detail page")
                return redirect(url_for('admin_operator_detail_stasiun', station_id=station_id))
            
        except Exception as e:
            db.session.rollback()
            error_msg = str(e)
            flash(f'❌ ERROR: {error_msg}', 'error')
            print(f"\n❌ ERROR in admin_operator_edit_stasiun: {error_msg}")
            import traceback
            traceback.print_exc()
            return redirect(url_for('admin_operator_edit_stasiun', station_id=station_id))
    
    # ===== GET REQUEST - TAMPILKAN FORM EDIT =====
    return render_template('admin_operator/edit_stasiun.html',
                         station=station,
                         kota_list=kota_list)

@app.route('/admin-operator/hapus-stasiun/<int:station_id>')
@login_required
@admin_operator_required
def admin_operator_hapus_stasiun(station_id):
    station = Stasiun.query.get_or_404(station_id)
    
    # Validasi kepemilikan
    if station.operator != current_user.operator_type:
        flash('❌ Stasiun ini bukan milik operator Anda!', 'error')
        return redirect(url_for('admin_operator_dashboard'))
    
    try:
        print(f"\n{'='*60}")
        print(f"🗑️  MENGHAPUS STASIUN: {station.stasiun_name} (ID: {station_id})")
        print(f"   Operator: {station.operator}")
        print(f"{'='*60}")
        
        # ===== 1. AMBIL SEMUA DATA TERKAIT =====
        # Ambil semua grup
        groups = GrupStasiun.query.filter_by(stasiun_id=station_id).all()
        group_ids = [g.id for g in groups]
        print(f"   📦 Grup ditemukan: {len(group_ids)}")
        
        # Ambil semua lawan
        opponents = StasiunLawan.query.filter_by(stasiun_id=station_id).all()
        opponent_ids = [o.id for o in opponents]
        print(f"   👥 Stasiun lawan: {len(opponent_ids)}")
        
        # Ambil semua upload
        all_uploads = UploadGambar.query.filter_by(stasiun_id=station_id).all()
        print(f"   🖼️ Upload gambar: {len(all_uploads)}")
        
        # ===== 2. HAPUS STATUS UPDATES =====
        if opponent_ids:
            deleted_status = StatusUpdate.query.filter(
                StatusUpdate.stasiun_lawan_id.in_(opponent_ids)
            ).delete(synchronize_session=False)
            print(f"   ✅ Status updates dihapus: {deleted_status}")
        
        # ===== 3. HAPUS UPLOAD GAMBAR =====
        cloudinary_deleted = 0
        for upload in all_uploads:
            try:
                # Cek apakah file masih digunakan oleh stasiun lain
                other_refs = UploadGambar.query.filter(
                    UploadGambar.public_id == upload.public_id,
                    UploadGambar.stasiun_id != station_id
                ).count()
                
                if other_refs == 0:
                    # Hapus dari Cloudinary
                    success, msg = delete_from_cloudinary(upload.public_id)
                    if success:
                        cloudinary_deleted += 1
                        print(f"      ✓ Cloudinary: {upload.public_id}")
                    else:
                        print(f"      ⚠ Cloudinary: {msg}")
                
                # Hapus record dari database
                db.session.delete(upload)
                
            except Exception as e:
                print(f"      ✗ Error hapus upload {upload.id}: {e}")
                continue
        
        print(f"   ✅ Upload gambar dihapus: {len(all_uploads)} (Cloudinary: {cloudinary_deleted})")
        
        # ===== 4. HAPUS GRUP STASIUN (URUTAN PENTING!) =====
        if group_ids:
            # PERTAMA: Putuskan relasi lawan ke grup
            for group_id in group_ids:
                updated = StasiunLawan.query.filter_by(
                    stasiun_id=station_id, 
                    grup_id=group_id
                ).update({
                    'group_id': None,
                    'grup_id': None
                })
                print(f"      ⤴️ Relasi grup {group_id} diputus: {updated} lawan")
            
            # KEDUA: Update upload yang terkait grup
            for group_id in group_ids:
                UploadGambar.query.filter_by(
                    stasiun_id=station_id,
                    group_id=group_id
                ).update({'group_id': None})
            
            # KETIGA: Hapus grup
            deleted_groups = GrupStasiun.query.filter_by(stasiun_id=station_id).delete()
            print(f"   ✅ Grup dihapus: {deleted_groups}")
        
        # ===== 5. HAPUS STASIUN LAWAN =====
        if opponent_ids:
            deleted_opponents = StasiunLawan.query.filter_by(stasiun_id=station_id).delete()
            print(f"   ✅ Stasiun lawan dihapus: {deleted_opponents}")
        
        # ===== 6. TERAKHIR, HAPUS STASIUN =====
        db.session.delete(station)
        
        # ===== 7. COMMIT SEMUA PERUBAHAN =====
        db.session.commit()
        
        print(f"\n{'='*60}")
        print(f"✅  BERHASIL MENGHAPUS STASIUN!")
        print(f"{'='*60}")
        
        flash(f'✅ Stasiun "{station.stasiun_name}" berhasil dihapus!', 'success')
        
    except Exception as e:
        db.session.rollback()
        error_msg = f"❌ Error menghapus stasiun: {str(e)}"
        print(f"\n{'='*60}")
        print(f"❌  ERROR: {error_msg}")
        print(f"{'='*60}")
        import traceback
        traceback.print_exc()
        flash(error_msg, 'error')
    
    return redirect(url_for('admin_operator_daftar_stasiun'))

@app.route('/admin-operator/upload-excel', methods=['GET', 'POST'])
@login_required
@admin_operator_required
def admin_operator_upload_excel():
    
    operator = current_user.operator_type
    
    if request.method == 'POST':
        try:
            if 'excel_file' not in request.files:
                flash('Tidak ada file yang dipilih!', 'error')
                return redirect(url_for('admin_operator_upload_excel'))
            
            file = request.files['excel_file']
            
            if file.filename == '':
                flash('Tidak ada file yang dipilih!', 'error')
                return redirect(url_for('admin_operator_upload_excel'))
            
            # Validasi file
            is_valid, error_msg = validate_excel_file(file)
            if not is_valid:
                flash(error_msg, 'error')
                return redirect(url_for('admin_operator_upload_excel'))
            
            kota_default = request.form.get('kota_default')
            if not kota_default:
                kota_default = None
                flash('Kota default tidak dipilih. Sistem akan menggunakan: 1) Kolom KOTA/KAB di Excel, 2) Default: Samarinda', 'info')
            
            # Simpan file sementara
            temp_dir = os.path.join(app.config['UPLOAD_FOLDER'], 'temp')
            os.makedirs(temp_dir, exist_ok=True)
            
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            temp_filename = f"{operator}_{timestamp}_{secure_filename(file.filename)}"
            temp_path = os.path.join(temp_dir, temp_filename)
            
            file.save(temp_path)
            
            # Process file khusus untuk operator
            success_count, error_count, messages = process_excel_upload_admin_operator(
                temp_path, 
                current_user,
                kota_default
            )
            
            # Hapus file temp
            try:
                os.remove(temp_path)
            except:
                pass
            
            # Tampilkan hasil
            if success_count > 0:
                flash(f'Berhasil menambahkan {success_count} stasiun untuk operator {operator.upper()}!', 'success')
            if error_count > 0:
                flash(f'Ada {error_count} error dalam proses upload.', 'warning')
            
            for msg in messages[-5:]:
                flash(msg, 'info')
            
            return redirect(url_for('admin_operator_daftar_stasiun'))
            
        except Exception as e:
            flash(f'Error: {str(e)}', 'error')
            return redirect(url_for('admin_operator_upload_excel'))
    
    # GET request
    total_stations = Stasiun.query.filter_by(operator=operator).count()
    
    total_opponents = 0
    stations = Stasiun.query.filter_by(operator=operator).all()
    for station in stations:
        total_opponents += len(station.stasiun_lawan_list)
    
    kota_list_op = db.session.query(Stasiun.kota)\
        .filter_by(operator=operator)\
        .distinct()\
        .all()
    kota_list_op = [k[0] for k in kota_list_op if k[0]]
    
    for kota in KOTA_LIST:
        if kota not in kota_list_op:
            kota_list_op.append(kota)
    
    return render_template('admin_operator/upload_excel.html',
                         operator=operator,
                         operator_name=operator.upper(),
                         total_stations=total_stations,
                         total_opponents=total_opponents,
                         kota_list=kota_list_op)

@app.route('/admin-operator/download-template')
@login_required
@admin_operator_required
def admin_operator_download_template():
    """Download template Excel untuk admin operator"""
    operator = current_user.operator_type
    return generate_admin_operator_template(operator)

# ============================================================================
# REKAP BULANAN DAN TAHUNAN
# ============================================================================

@app.route('/admin-master/rekap-bulanan')
@login_required
@admin_master_required
def admin_master_rekap_bulanan():
    """Halaman rekap bulanan - HANYA menampilkan form filter dan tombol export"""
    try:
        # Ambil parameter filter dari URL, default ke tahun dan bulan sekarang
        tahun = request.args.get('tahun', datetime.now().year, type=int)
        bulan = request.args.get('bulan', datetime.now().month, type=int)
        operator = request.args.get('operator', 'all')
        kota = request.args.get('kota', 'all')
        
        # Validasi bulan (1-12)
        if bulan < 1 or bulan > 12:
            bulan = datetime.now().month
        
        # Daftar bulan dalam bahasa Indonesia
        bulan_nama = ['Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni',
                     'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember']
        
        # ===== TAMBAHKAN: Validasi operator yang dipilih =====
        if operator != 'all' and operator not in OPERATORS:
            print(f"⚠️ Operator '{operator}' tidak valid, reset ke 'all'")
            operator = 'all'
        
        # ===== PERBAIKAN: Ambil daftar kota case-insensitive =====
        from sqlalchemy import func
        
        kota_raw = db.session.query(
            func.lower(Stasiun.kota).label('kota_lower'),
            Stasiun.kota
        ).distinct(func.lower(Stasiun.kota))\
         .order_by(func.lower(Stasiun.kota))\
         .all()
        
        # Mapping untuk menyimpan format standar
        kota_mapping = {}
        for kota_lower, kota_original in kota_raw:
            if kota_lower not in kota_mapping:
                # Cari format standar dari KOTA_LIST
                found = False
                for std_kota in KOTA_LIST:
                    if std_kota.lower() == kota_lower:
                        kota_mapping[kota_lower] = std_kota
                        found = True
                        break
                if not found:
                    kota_mapping[kota_lower] = kota_original.strip().title()
        
        kota_list = list(kota_mapping.values())
        
        # Jika tidak ada data kota di database, gunakan KOTA_LIST sebagai fallback
        if not kota_list:
            kota_list = KOTA_LIST
            print("⚠️ Tidak ada kota di database, menggunakan KOTA_LIST default")
        
        # Validasi kota yang dipilih
        if kota != 'all':
            kota_found = False
            for k in kota_list:
                if k.lower() == kota.lower():
                    kota = k
                    kota_found = True
                    break
            if not kota_found:
                print(f"⚠️ Kota '{kota}' tidak ada di list, reset ke 'all'")
                kota = 'all'
        
        print(f"\n{'='*60}")
        print(f"📅 HALAMAN REKAP BULANAN")
        print(f"   Tahun: {tahun}, Bulan: {bulan_nama[bulan-1]}, Operator: {operator}, Kota/Kab: {kota}")
        print('='*60)
        
        return render_template('admin_master/rekap_bulanan.html',
                             tahun=tahun,
                             bulan=bulan,
                             bulan_nama=bulan_nama[bulan-1],
                             operator=operator,
                             kota=kota,
                             operators=OPERATORS,
                             kota_list=kota_list)
    except Exception as e:
        flash(f'Error: {str(e)}', 'error')
        print(f"❌ ERROR in rekap bulanan: {str(e)}")
        import traceback
        traceback.print_exc()
        return redirect(url_for('admin_master_dashboard'))
    
@app.route('/admin-master/rekap-bulanan-word-gambar')
@login_required
@admin_master_required
def admin_master_rekap_bulanan_word_gambar():
    """Export rekap bulanan ke Word dengan gambar besar per grup"""
    try:
        # ===== 1. AMBIL DAN VALIDASI PARAMETER =====
        tahun = request.args.get('tahun', datetime.now().year, type=int)
        bulan = request.args.get('bulan', datetime.now().month, type=int)
        operator = request.args.get('operator', 'all')
        kota = request.args.get('kota', 'all')
        
        if bulan < 1 or bulan > 12:
            bulan = datetime.now().month
        
        if operator != 'all' and operator not in OPERATORS:
            operator = 'all'
        
        bulan_nama = ['Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni',
                     'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember']
        
        print(f"\n{'='*80}")
        print(f"📄 EXPORT REKAP BULANAN WORD - ADMIN MASTER")
        print(f"   Tahun: {tahun}, Bulan: {bulan_nama[bulan-1]}")
        print(f"   Operator: {operator if operator != 'all' else 'SEMUA'}")
        print(f"   Kota/Kab: {kota if kota != 'all' else 'SEMUA'}")
        print('='*80)
        
        # ===== 2. QUERY STASIUN DENGAN FILTER (CASE-INSENSITIVE) =====
        query = Stasiun.query
        
        if operator != 'all':
            query = query.filter(Stasiun.operator == operator)
            print(f"✅ Filter operator: {operator}")
        
        if kota != 'all' and kota:
            query = query.filter(Stasiun.kota.ilike(kota))
            print(f"✅ Filter kota dengan ILIKE: {kota}")
        
        all_stations = query.order_by(Stasiun.operator, Stasiun.kota, Stasiun.stasiun_name).all()
        print(f"📊 Total stasiun setelah filter: {len(all_stations)}")
        
        # ===== 3. FILTER STASIUN YANG MEMILIKI DATA DI BULAN INI =====
        stations_with_data = []
        for station in all_stations:
            has_data = db.session.query(StatusUpdate)\
                .join(StasiunLawan)\
                .filter(StasiunLawan.stasiun_id == station.id)\
                .filter(extract('year', StatusUpdate.updated_at) == tahun)\
                .filter(extract('month', StatusUpdate.updated_at) == bulan)\
                .first()
            
            has_upload = UploadGambar.query\
                .filter_by(stasiun_id=station.id)\
                .filter(extract('year', UploadGambar.uploaded_at) == tahun)\
                .filter(extract('month', UploadGambar.uploaded_at) == bulan)\
                .first()
            
            if has_data or has_upload:
                stations_with_data.append(station)
        
        print(f"📊 Stasiun dengan data di bulan {bulan_nama[bulan-1]}: {len(stations_with_data)}")
        
        if not stations_with_data:
            error_msg = f"Tidak ada data untuk periode yang dipilih: {bulan_nama[bulan-1]} {tahun}"
            if operator != 'all':
                error_msg += f" (Operator: {operator.upper()})"
            if kota != 'all':
                error_msg += f" (Kota: {kota})"
            flash(error_msg, 'warning')
            return redirect(url_for('admin_master_rekap_bulanan', tahun=tahun, bulan=bulan, operator=operator, kota=kota))
        
        # ===== 4. BUAT DOKUMEN WORD =====
        from docx import Document
        from docx.shared import Inches, Pt, RGBColor, Cm
        from docx.enum.text import WD_ALIGN_PARAGRAPH
        from docx.enum.table import WD_TABLE_ALIGNMENT
        from docx.enum.section import WD_ORIENT
        from cloudinary.utils import cloudinary_url
        import requests
        from io import BytesIO
        
        doc = Document()
        
        # ===== HALAMAN 1: LANDSCAPE UNTUK RINGKASAN =====
        section = doc.sections[0]
        original_width = section.page_width
        original_height = section.page_height
        
        section.orientation = WD_ORIENT.LANDSCAPE
        section.page_width = original_height
        section.page_height = original_width
        section.top_margin = Cm(1.5)
        section.bottom_margin = Cm(1.5)
        section.left_margin = Cm(1.5)
        section.right_margin = Cm(1.5)
        
        # ===== HALAMAN COVER =====
        title = doc.add_heading('REKAP BULANAN STASIUN MICROWAVE LINK', 0)
        title.alignment = WD_ALIGN_PARAGRAPH.CENTER
        title.runs[0].font.size = Pt(24)
        title.runs[0].font.bold = True
        title.runs[0].font.color.rgb = RGBColor(0, 51, 102)
        
        doc.add_paragraph('\n' * 2)
        
        periode = doc.add_heading(f'{bulan_nama[bulan-1]} {tahun}', level=1)
        periode.alignment = WD_ALIGN_PARAGRAPH.CENTER
        periode.runs[0].font.size = Pt(28)
        periode.runs[0].font.color.rgb = RGBColor(0, 102, 204)
        
        doc.add_paragraph('\n' * 2)
        
        filter_text = f"Operator: {operator.upper() if operator != 'all' else 'SEMUA OPERATOR'} | "
        filter_text += f"Kota/Kab: {kota.title() if kota != 'all' else 'Semua Kota/kab'}"
        filter_para = doc.add_paragraph(filter_text)
        filter_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        filter_para.runs[0].font.size = Pt(14)
        filter_para.runs[0].bold = True
        
        doc.add_paragraph('\n' * 2)
        
        footer = doc.add_paragraph()
        footer.alignment = WD_ALIGN_PARAGRAPH.CENTER
        footer.add_run(f'Dicetak pada: {datetime.now().strftime("%d %B %Y")}').italic = True
        footer.add_run('\n')
        footer.add_run(f'Oleh: {current_user.username}').italic = True
        
        doc.add_page_break()
        
        # ===== RINGKASAN DATA (LANDSCAPE) =====
        doc.add_heading('RINGKASAN DATA STASIUN LAWAN', level=1).alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        all_rows = []
        row_counter = 0
        status_summary = {'aktif': 0, 'tidak_aktif': 0, 'tidak_sesuai': 0, 'belum_ada': 0}
        
        for station in stations_with_data:
            opponents = StasiunLawan.query.filter_by(stasiun_id=station.id).order_by(StasiunLawan.urutan).all()
            
            for opp in opponents:
                latest_status = StatusUpdate.query\
                    .filter_by(stasiun_lawan_id=opp.id)\
                    .filter(extract('year', StatusUpdate.updated_at) == tahun)\
                    .filter(extract('month', StatusUpdate.updated_at) == bulan)\
                    .order_by(StatusUpdate.updated_at.desc())\
                    .first()
                
                if latest_status:
                    row_counter += 1
                    all_rows.append({
                        'no': row_counter,
                        'operator': station.operator.upper(),
                        'stasiun_name': station.stasiun_name,
                        'lawan': opp.nama_stasiun_lawan,
                        'freq_tx': opp.freq_tx if opp.freq_tx else '-',
                        'freq_rx': opp.freq_rx if opp.freq_rx else '-',
                        'kota': station.kota,
                        'status': latest_status.status
                    })
                    
                    if latest_status.status in status_summary:
                        status_summary[latest_status.status] += 1
                else:
                    row_counter += 1
                    all_rows.append({
                        'no': row_counter,
                        'operator': station.operator.upper(),
                        'stasiun_name': station.stasiun_name,
                        'lawan': opp.nama_stasiun_lawan,
                        'freq_tx': opp.freq_tx if opp.freq_tx else '-',
                        'freq_rx': opp.freq_rx if opp.freq_rx else '-',
                        'kota': station.kota,
                        'status': 'belum_ada'
                    })
                    status_summary['belum_ada'] += 1
        
        if all_rows:
            rows_per_page = 35
            total_rows = len(all_rows)
            total_pages = (total_rows + rows_per_page - 1) // rows_per_page
            
            for page in range(total_pages):
                if page > 0:
                    doc.add_page_break()
                
                start_idx = page * rows_per_page
                end_idx = min(start_idx + rows_per_page, total_rows)
                page_data = all_rows[start_idx:end_idx]
                
                if total_pages > 1:
                    page_info = doc.add_paragraph()
                    page_info.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                    page_info.add_run(f'Halaman {page+1} dari {total_pages} (Ringkasan)').italic = True
                    page_info.runs[0].font.size = Pt(9)
                
                table = doc.add_table(rows=1, cols=7)
                table.style = 'Table Grid'
                table.autofit = False
                
                table.columns[0].width = Cm(1.0)
                table.columns[1].width = Cm(2.0)
                table.columns[2].width = Cm(3.0)
                table.columns[3].width = Cm(4.0)
                table.columns[4].width = Cm(2.0)
                table.columns[5].width = Cm(2.0)
                table.columns[6].width = Cm(2.5)
                
                header_cells = table.rows[0].cells
                headers = ['No', 'Operator', 'Stasiun', 'Stasiun Lawan', 'Freq Tx', 'Freq Rx', 'Kota']
                for i, header in enumerate(headers):
                    header_cells[i].text = header
                    header_cells[i].paragraphs[0].runs[0].font.bold = True
                    header_cells[i].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                
                for row in page_data:
                    row_cells = table.add_row().cells
                    row_cells[0].text = str(row['no'])
                    row_cells[0].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                    row_cells[1].text = row['operator']
                    row_cells[1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                    row_cells[2].text = row['stasiun_name']
                    row_cells[2].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.LEFT
                    row_cells[3].text = row['lawan']
                    row_cells[3].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.LEFT
                    row_cells[4].text = row['freq_tx']
                    row_cells[4].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                    row_cells[5].text = row['freq_rx']
                    row_cells[5].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                    row_cells[6].text = row['kota']
                    row_cells[6].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.LEFT
        
        doc.add_paragraph()
        summary_para = doc.add_paragraph()
        summary_para.add_run("REKAPITULASI STATUS:\n").bold = True
        summary_para.add_run(f"Aktif: {status_summary['aktif']} | ")
        summary_para.add_run(f"Tidak Aktif: {status_summary['tidak_aktif']} | ")
        summary_para.add_run(f"Tidak Sesuai: {status_summary['tidak_sesuai']} | ")
        summary_para.add_run(f"Belum Ada Status: {status_summary['belum_ada']}")
        
        doc.add_page_break()
        
        # ===== HALAMAN DOKUMENTASI GAMBAR (PORTRAIT) =====
        new_section = doc.add_section()
        new_section.orientation = WD_ORIENT.PORTRAIT
        new_section.page_width = original_width
        new_section.page_height = original_height
        new_section.top_margin = Cm(2)
        new_section.bottom_margin = Cm(2)
        new_section.left_margin = Cm(2.5)
        new_section.right_margin = Cm(2.5)
        
        doc.add_heading('DOKUMENTASI GAMBAR PER GRUP', level=1).alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        for station in stations_with_data:
            doc.add_heading(f'STASIUN: {station.stasiun_name}', level=2)
            
            info_para = doc.add_paragraph()
            info_para.add_run(f'Operator: {station.operator.upper()}').bold = True
            info_para.add_run(f' | Kota/Kab: {station.kota}').bold = True
            
            groups = GrupStasiun.query.filter_by(stasiun_id=station.id).all()
            groups_dict = {g.id: g.nama_grup for g in groups}
            
            all_opponents = StasiunLawan.query.filter_by(stasiun_id=station.id).order_by(StasiunLawan.urutan).all()
            
            opponents_by_group = {}
            for opp in all_opponents:
                group_key = opp.group_id if opp.group_id is not None else 'ungrouped'
                if group_key not in opponents_by_group:
                    opponents_by_group[group_key] = []
                
                latest_status = StatusUpdate.query\
                    .filter_by(stasiun_lawan_id=opp.id)\
                    .filter(extract('year', StatusUpdate.updated_at) == tahun)\
                    .filter(extract('month', StatusUpdate.updated_at) == bulan)\
                    .order_by(StatusUpdate.updated_at.desc())\
                    .first()
                
                opponents_by_group[group_key].append({
                    'nama': opp.nama_stasiun_lawan,
                    'freq_tx': opp.freq_tx if opp.freq_tx else '-',
                    'freq_rx': opp.freq_rx if opp.freq_rx else '-',
                    'status': latest_status.status if latest_status else 'belum_ada',
                    'status_display': STATUS_OPTIONS.get(latest_status.status, 'Belum Ada') if latest_status else 'Belum Ada',
                    'catatan': latest_status.catatan if latest_status and latest_status.catatan else '-'
                })
            
            # ===== AMBIL SEMUA UPLOAD (TANPA LIMIT) =====
            uploads_in_bulan = UploadGambar.query\
                .filter_by(stasiun_id=station.id)\
                .filter(extract('year', UploadGambar.uploaded_at) == tahun)\
                .filter(extract('month', UploadGambar.uploaded_at) == bulan)\
                .order_by(UploadGambar.uploaded_at.desc())\
                .all()
            
            uploads_by_group = {}
            seen_uploads = set()
            
            for upload in uploads_in_bulan:
                group_key = upload.group_id if upload.group_id is not None else 'ungrouped'
                upload_key = f"{group_key}_{upload.public_id}"
                
                if upload_key not in seen_uploads:
                    seen_uploads.add(upload_key)
                    
                    if group_key not in uploads_by_group:
                        uploads_by_group[group_key] = []
                    
                    uploads_by_group[group_key].append(upload)
            
            all_group_keys = set(list(opponents_by_group.keys()) + list(uploads_by_group.keys()))
            
            sorted_group_keys = []
            if 'ungrouped' in all_group_keys:
                sorted_group_keys.append('ungrouped')
            
            numeric_keys = [k for k in all_group_keys if k != 'ungrouped' and k is not None]
            numeric_keys.sort(key=lambda x: int(x) if str(x).isdigit() else x)
            sorted_group_keys.extend(numeric_keys)
            
            group_counter = 0
            for group_key in sorted_group_keys:
                group_counter += 1
                
                if group_key == 'ungrouped':
                    group_display = 'Tanpa Grup'
                else:
                    group_display = groups_dict.get(int(group_key), f'Grup {group_key}')
                
                doc.add_heading(f'  Grup {group_counter}: {group_display}', level=3)
                
                group_uploads = uploads_by_group.get(group_key, [])
                
                if group_uploads:
                    doc.add_heading('    GAMBAR:', level=4)
                    
                    for upload in group_uploads:
                        try:
                            # ===== PERUBAHAN: Gunakan thumbnail =====
                            thumbnail_url, _ = cloudinary_url(
                                upload.public_id,
                                width=500,
                                height=375,
                                crop="limit",
                                quality="auto:eco"
                            )
                            response = requests.get(thumbnail_url, timeout=30)
                            if response.status_code == 200:
                                img_data = BytesIO(response.content)
                                
                                paragraph = doc.add_paragraph()
                                paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                                run = paragraph.add_run()
                                run.add_picture(img_data, width=Inches(5.0))
                                
                                info_para = doc.add_paragraph()
                                info_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                                info_text = f"File: {upload.original_filename} | "
                                info_text += f"Upload: {upload.uploaded_at.strftime('%d/%m/%Y')}"
                                info_run = info_para.add_run(info_text)
                                info_run.italic = True
                                info_run.font.size = Pt(9)
                                
                                doc.add_paragraph()
                        except Exception as e:
                            print(f"Error download: {e}")
                            doc.add_paragraph(f"    Gambar tidak tersedia: {upload.original_filename}")
                else:
                    doc.add_paragraph("    Tidak ada gambar untuk grup ini")
                
                group_opponents = opponents_by_group.get(group_key, [])
                if group_opponents:
                    doc.add_heading('    DAFTAR STASIUN LAWAN:', level=4)
                    
                    opp_table = doc.add_table(rows=1, cols=6)
                    opp_table.style = 'Table Grid'
                    opp_table.autofit = True
                    
                    opp_headers = ['No', 'Nama Stasiun Lawan', 'Freq Tx', 'Freq Rx', 'Status', 'Catatan']
                    for i, header in enumerate(opp_headers):
                        opp_table.rows[0].cells[i].text = header
                        opp_table.rows[0].cells[i].paragraphs[0].runs[0].font.bold = True
                        opp_table.rows[0].cells[i].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                    
                    for opp_idx, opponent in enumerate(group_opponents, 1):
                        row_cells = opp_table.add_row().cells
                        row_cells[0].text = str(opp_idx)
                        row_cells[0].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                        row_cells[1].text = opponent['nama']
                        row_cells[1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.LEFT
                        row_cells[2].text = opponent['freq_tx']
                        row_cells[2].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                        row_cells[3].text = opponent['freq_rx']
                        row_cells[3].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                        row_cells[4].text = opponent['status_display']
                        row_cells[4].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                        row_cells[5].text = opponent['catatan']
                        row_cells[5].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.LEFT
                    
                    doc.add_paragraph()
                else:
                    doc.add_paragraph("    Tidak ada stasiun lawan dalam grup ini")
                
                doc.add_paragraph()
            
            doc.add_paragraph()
        
        file_stream = BytesIO()
        doc.save(file_stream)
        file_stream.seek(0)
        
        filename = f"Rekap_Bulanan_{bulan_nama[bulan-1]}_{tahun}"
        if operator != 'all':
            filename += f"_{operator.upper()}"
        if kota != 'all':
            filename += f"_{kota.replace(' ', '_')}"
        filename += ".docx"
        
        print(f"\n✅ File berhasil dibuat: {filename}")
        print(f"   Total data: {len(all_rows)}")
        print('='*80)
        
        return send_file(
            file_stream,
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            as_attachment=True,
            download_name=filename
        )
        
    except Exception as e:
        flash(f'Error: {str(e)}', 'error')
        print(f"\n❌ ERROR: {str(e)}")
        import traceback
        traceback.print_exc()
        return redirect(url_for('admin_master_rekap_bulanan', 
                               tahun=request.args.get('tahun', datetime.now().year, type=int),
                               bulan=request.args.get('bulan', datetime.now().month, type=int),
                               operator=request.args.get('operator', 'all'),
                               kota=request.args.get('kota', 'all')))

@app.route('/admin-master/rekap-tahunan')
@login_required
@admin_master_required
def admin_master_rekap_tahunan():
    """Halaman rekap tahunan - HANYA menampilkan form filter dan tombol export"""
    try:
        # Ambil parameter filter dari URL, default ke tahun sekarang
        tahun = request.args.get('tahun', datetime.now().year, type=int)
        operator = request.args.get('operator', 'all')
        kota = request.args.get('kota', 'all')
        
        # ===== TAMBAHKAN: Validasi kota yang dipilih ada di KOTA_LIST =====
        if kota != 'all' and kota not in KOTA_LIST:
            print(f"⚠️ Kota '{kota}' tidak ada di KOTA_LIST, reset ke 'all'")
            kota = 'all'
        
        # ===== TAMBAHKAN: Validasi operator yang dipilih =====
        if operator != 'all' and operator not in OPERATORS:
            print(f"⚠️ Operator '{operator}' tidak valid, reset ke 'all'")
            operator = 'all'
        
        # ===== PERBAIKAN: Ambil daftar kota case-insensitive =====
        from sqlalchemy import func
        
        kota_raw = db.session.query(
            func.lower(Stasiun.kota).label('kota_lower'),
            Stasiun.kota
        ).distinct(func.lower(Stasiun.kota))\
         .order_by(func.lower(Stasiun.kota))\
         .all()
        
        # Mapping untuk menyimpan format standar
        kota_mapping = {}
        for kota_lower, kota_original in kota_raw:
            if kota_lower not in kota_mapping:
                # Cari format standar dari KOTA_LIST
                found = False
                for std_kota in KOTA_LIST:
                    if std_kota.lower() == kota_lower:
                        kota_mapping[kota_lower] = std_kota
                        found = True
                        break
                if not found:
                    kota_mapping[kota_lower] = kota_original.strip().title()
        
        kota_list = list(kota_mapping.values())
        
        # Jika tidak ada data kota di database, gunakan KOTA_LIST sebagai fallback
        if not kota_list:
            kota_list = KOTA_LIST
            print("⚠️ Tidak ada kota di database, menggunakan KOTA_LIST default")
        
        # Validasi kota yang dipilih
        if kota != 'all':
            kota_found = False
            for k in kota_list:
                if k.lower() == kota.lower():
                    kota = k
                    kota_found = True
                    break
            if not kota_found:
                print(f"⚠️ Kota '{kota}' tidak ada di list, reset ke 'all'")
                kota = 'all'
        
        print(f"\n{'='*60}")
        print(f"📅 HALAMAN REKAP TAHUNAN")
        print(f"   Tahun: {tahun}, Operator: {operator}, Kota/Kab: {kota}")
        print('='*60)
        
        return render_template('admin_master/rekap_tahunan.html',
                             tahun=tahun,
                             operator=operator,
                             kota=kota,
                             operators=OPERATORS,
                             kota_list=kota_list)
    except Exception as e:
        flash(f'Error: {str(e)}', 'error')
        print(f"❌ ERROR in rekap tahunan: {str(e)}")
        import traceback
        traceback.print_exc()
        return redirect(url_for('admin_master_dashboard'))                           
    
@app.route('/admin-master/rekap-tahunan-word-gambar')
@login_required
@admin_master_required
def admin_master_rekap_tahunan_word_gambar():
    """Export rekap tahunan ke Word dengan gambar per grup - HANYA bulan dengan data"""
    try:
        # ===== 1. AMBIL DAN VALIDASI PARAMETER =====
        tahun = request.args.get('tahun', datetime.now().year, type=int)
        operator = request.args.get('operator', 'all')
        kota = request.args.get('kota', 'all')
        
        if operator != 'all' and operator not in OPERATORS:
            operator = 'all'
        
        bulan_nama = ['Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni',
                     'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember']
        
        print(f"\n{'='*80}")
        print(f"📄 EXPORT REKAP TAHUNAN WORD - ADMIN MASTER")
        print(f"   Tahun: {tahun}")
        print(f"   Operator: {operator if operator != 'all' else 'SEMUA'}")
        print(f"   Kota/Kab: {kota if kota != 'all' else 'SEMUA'}")
        print('='*80)
        
        # ===== 2. QUERY STASIUN DENGAN FILTER (CASE-INSENSITIVE) =====
        query = Stasiun.query
        
        if operator != 'all':
            query = query.filter(Stasiun.operator == operator)
            print(f"✅ Filter operator: {operator}")
        
        if kota != 'all' and kota:
            query = query.filter(Stasiun.kota.ilike(kota))
            print(f"✅ Filter kota dengan ILIKE: {kota}")
        
        all_stations = query.order_by(Stasiun.operator, Stasiun.kota, Stasiun.stasiun_name).all()
        print(f"📊 Total stasiun setelah filter: {len(all_stations)}")
        
        # ===== 3. FILTER STASIUN YANG MEMILIKI DATA DI TAHUN TERSEBUT =====
        stations_with_data = []
        for station in all_stations:
            has_data = db.session.query(StatusUpdate)\
                .join(StasiunLawan)\
                .filter(StasiunLawan.stasiun_id == station.id)\
                .filter(extract('year', StatusUpdate.updated_at) == tahun)\
                .first()
            
            has_upload = UploadGambar.query\
                .filter_by(stasiun_id=station.id)\
                .filter(extract('year', UploadGambar.uploaded_at) == tahun)\
                .first()
            
            if has_data or has_upload:
                stations_with_data.append(station)
        
        print(f"📊 Stasiun dengan data di tahun {tahun}: {len(stations_with_data)}")
        
        if not stations_with_data:
            error_msg = f"Tidak ada data untuk tahun {tahun}"
            if operator != 'all':
                error_msg += f" (Operator: {operator.upper()})"
            if kota != 'all':
                error_msg += f" (Kota: {kota})"
            flash(error_msg, 'warning')
            return redirect(url_for('admin_master_rekap_tahunan', tahun=tahun, operator=operator, kota=kota))
        
        # ===== 4. KUMPULKAN DATA PER BULAN =====
        bulan_data = {}
        bulan_dengan_data = []
        
        for bulan in range(1, 13):
            bulan_key = f"{tahun}-{bulan:02d}"
            bulan_data[bulan_key] = {
                'nama': bulan_nama[bulan-1],
                'stations': [],
                'total_uploads': 0,
                'has_data': False,
                'status_summary': {'aktif': 0, 'tidak_aktif': 0, 'tidak_sesuai': 0, 'belum_ada': 0}
            }
            
            for station in stations_with_data:
                status_count = db.session.query(StatusUpdate)\
                    .join(StasiunLawan)\
                    .filter(StasiunLawan.stasiun_id == station.id)\
                    .filter(extract('year', StatusUpdate.updated_at) == tahun)\
                    .filter(extract('month', StatusUpdate.updated_at) == bulan)\
                    .count()
                
                upload_count = UploadGambar.query\
                    .filter_by(stasiun_id=station.id)\
                    .filter(extract('year', UploadGambar.uploaded_at) == tahun)\
                    .filter(extract('month', UploadGambar.uploaded_at) == bulan)\
                    .count()
                
                if status_count > 0 or upload_count > 0:
                    bulan_data[bulan_key]['has_data'] = True
                    bulan_data[bulan_key]['stations'].append(station)
                    bulan_data[bulan_key]['total_uploads'] += upload_count
                    
                    opponents = StasiunLawan.query.filter_by(stasiun_id=station.id).all()
                    for opp in opponents:
                        latest_status = StatusUpdate.query\
                            .filter_by(stasiun_lawan_id=opp.id)\
                            .filter(extract('year', StatusUpdate.updated_at) == tahun)\
                            .filter(extract('month', StatusUpdate.updated_at) == bulan)\
                            .order_by(StatusUpdate.updated_at.desc())\
                            .first()
                        
                        if latest_status and latest_status.status in bulan_data[bulan_key]['status_summary']:
                            bulan_data[bulan_key]['status_summary'][latest_status.status] += 1
                        else:
                            bulan_data[bulan_key]['status_summary']['belum_ada'] += 1
            
            if bulan_data[bulan_key]['has_data']:
                bulan_dengan_data.append(bulan)
                print(f"  ✅ {bulan_nama[bulan-1]}: {len(bulan_data[bulan_key]['stations'])} stasiun, "
                      f"{bulan_data[bulan_key]['total_uploads']} uploads")
        
        # ===== 5. BUAT DOKUMEN WORD =====
        from docx import Document
        from docx.shared import Inches, Pt, RGBColor, Cm
        from docx.enum.text import WD_ALIGN_PARAGRAPH
        from docx.enum.table import WD_TABLE_ALIGNMENT
        from docx.enum.section import WD_ORIENT
        from cloudinary.utils import cloudinary_url
        import requests
        from io import BytesIO
        
        doc = Document()
        
        section = doc.sections[0]
        original_width = section.page_width
        original_height = section.page_height
        
        section.orientation = WD_ORIENT.LANDSCAPE
        section.page_width = original_height
        section.page_height = original_width
        section.top_margin = Cm(1.5)
        section.bottom_margin = Cm(1.5)
        section.left_margin = Cm(1.5)
        section.right_margin = Cm(1.5)
        
        title = doc.add_heading('REKAP TAHUNAN STASIUN MICROWAVE LINK', 0)
        title.alignment = WD_ALIGN_PARAGRAPH.CENTER
        title.runs[0].font.size = Pt(24)
        title.runs[0].font.bold = True
        title.runs[0].font.color.rgb = RGBColor(0, 51, 102)
        
        doc.add_paragraph('\n' * 4)
        
        tahun_heading = doc.add_heading(f'TAHUN {tahun}', level=1)
        tahun_heading.alignment = WD_ALIGN_PARAGRAPH.CENTER
        tahun_heading.runs[0].font.size = Pt(28)
        tahun_heading.runs[0].font.color.rgb = RGBColor(0, 102, 204)
        
        doc.add_paragraph('\n' * 4)
        
        filter_text = f"Operator: {operator.upper() if operator != 'all' else 'SEMUA OPERATOR'}"
        filter_para = doc.add_paragraph(filter_text)
        filter_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        filter_para.runs[0].font.size = Pt(14)
        filter_para.runs[0].bold = True
        
        kota_text = f"Kota/Kab: {kota.title() if kota != 'all' else 'Semua Kota/kab'}"
        kota_para = doc.add_paragraph(kota_text)
        kota_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        kota_para.runs[0].font.size = Pt(14)
        kota_para.runs[0].bold = True
        
        doc.add_paragraph('\n' * 2)
        
        footer = doc.add_paragraph()
        footer.alignment = WD_ALIGN_PARAGRAPH.CENTER
        footer.add_run(f'Dicetak pada: {datetime.now().strftime("%d %B %Y")}').italic = True
        footer.add_run('\n')
        footer.add_run(f'Oleh: {current_user.username}').italic = True
        
        doc.add_page_break()
        
        doc.add_heading('RINGKASAN STATISTIK TAHUNAN', level=1).alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        all_rows_data = []
        row_counter = 0
        
        for station in stations_with_data:
            opponents = StasiunLawan.query.filter_by(stasiun_id=station.id).order_by(StasiunLawan.urutan).all()
            
            for opp in opponents:
                row_counter += 1
                all_rows_data.append({
                    'no': row_counter,
                    'operator': station.operator.upper(),
                    'stasiun_name': station.stasiun_name,
                    'stasiun_lawan': opp.nama_stasiun_lawan,
                    'freq_tx': opp.freq_tx if opp.freq_tx else '-',
                    'freq_rx': opp.freq_rx if opp.freq_rx else '-',
                    'kota': station.kota
                })
        
        if all_rows_data:
            rows_per_page = 35
            total_rows = len(all_rows_data)
            total_pages = (total_rows + rows_per_page - 1) // rows_per_page
            
            for page in range(total_pages):
                if page > 0:
                    doc.add_page_break()
                
                start_idx = page * rows_per_page
                end_idx = min(start_idx + rows_per_page, total_rows)
                page_data = all_rows_data[start_idx:end_idx]
                
                if total_pages > 1:
                    page_info = doc.add_paragraph()
                    page_info.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                    page_info.add_run(f'Halaman {page+1} dari {total_pages} (Ringkasan)').italic = True
                    page_info.runs[0].font.size = Pt(9)
                
                main_table = doc.add_table(rows=1, cols=7)
                main_table.style = 'Table Grid'
                main_table.autofit = False
                
                main_table.columns[0].width = Cm(1.0)
                main_table.columns[1].width = Cm(2.0)
                main_table.columns[2].width = Cm(3.0)
                main_table.columns[3].width = Cm(4.0)
                main_table.columns[4].width = Cm(2.0)
                main_table.columns[5].width = Cm(2.0)
                main_table.columns[6].width = Cm(2.5)
                
                header_cells = main_table.rows[0].cells
                headers = ['No', 'Operator', 'Stasiun Name', 'Stasiun Lawan', 'Freq Tx', 'Freq Rx', 'Kota/Kab']
                for i, header in enumerate(headers):
                    header_cells[i].text = header
                    header_cells[i].paragraphs[0].runs[0].font.bold = True
                    header_cells[i].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                
                for data in page_data:
                    row_cells = main_table.add_row().cells
                    row_cells[0].text = str(data['no'])
                    row_cells[0].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                    row_cells[1].text = data['operator']
                    row_cells[1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                    row_cells[2].text = data['stasiun_name']
                    row_cells[2].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.LEFT
                    row_cells[3].text = data['stasiun_lawan']
                    row_cells[3].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.LEFT
                    row_cells[4].text = data['freq_tx']
                    row_cells[4].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                    row_cells[5].text = data['freq_rx']
                    row_cells[5].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                    row_cells[6].text = data['kota']
                    row_cells[6].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.LEFT
        
        doc.add_page_break()
        
        new_section = doc.add_section()
        new_section.orientation = WD_ORIENT.PORTRAIT
        new_section.page_width = original_width
        new_section.page_height = original_height
        new_section.top_margin = Cm(2)
        new_section.bottom_margin = Cm(2)
        new_section.left_margin = Cm(2.5)
        new_section.right_margin = Cm(2.5)
        
        doc.add_heading('DOKUMENTASI GAMBAR PER BULAN', level=1).alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        if not bulan_dengan_data:
            doc.add_paragraph(f"Tidak ada data untuk tahun {tahun}")
        else:
            for bulan in sorted(bulan_dengan_data):
                bulan_key = f"{tahun}-{bulan:02d}"
                bulan_info = bulan_data[bulan_key]
                
                print(f"\n📷 Memproses bulan: {bulan_info['nama']} {tahun}")
                
                doc.add_heading(f'{bulan_info["nama"].upper()} {tahun}', level=2)
                
                status = bulan_info['status_summary']
                summary_text = (f"Rekapitulasi Status: Aktif: {status['aktif']} | "
                               f"Tidak Aktif: {status['tidak_aktif']} | "
                               f"Tidak Sesuai: {status['tidak_sesuai']} | "
                               f"Belum Ada Status: {status['belum_ada']}")
                summary_para = doc.add_paragraph(summary_text)
                summary_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                summary_para.runs[0].italic = True
                
                doc.add_paragraph()
                
                for station in bulan_info['stations']:
                    station_heading = doc.add_heading(f'STASIUN: {station.stasiun_name}', level=3)
                    station_heading.runs[0].font.color.rgb = RGBColor(67, 97, 238)
                    
                    info_para = doc.add_paragraph()
                    info_para.add_run(f'Operator: {station.operator.upper()}').bold = True
                    info_para.add_run(f' | Kota/Kab: {station.kota}').bold = True
                    
                    groups = GrupStasiun.query.filter_by(stasiun_id=station.id).all()
                    groups_dict = {g.id: g.nama_grup for g in groups}
                    
                    all_opponents = StasiunLawan.query.filter_by(stasiun_id=station.id).order_by(StasiunLawan.urutan).all()
                    
                    opponents_by_group = {}
                    for opp in all_opponents:
                        group_key = opp.group_id if opp.group_id is not None else 'ungrouped'
                        if group_key not in opponents_by_group:
                            opponents_by_group[group_key] = []
                        
                        latest_status = StatusUpdate.query\
                            .filter_by(stasiun_lawan_id=opp.id)\
                            .filter(extract('year', StatusUpdate.updated_at) == tahun)\
                            .filter(extract('month', StatusUpdate.updated_at) == bulan)\
                            .order_by(StatusUpdate.updated_at.desc())\
                            .first()
                        
                        opponents_by_group[group_key].append({
                            'nama': opp.nama_stasiun_lawan,
                            'freq_tx': opp.freq_tx if opp.freq_tx else '-',
                            'freq_rx': opp.freq_rx if opp.freq_rx else '-',
                            'status': latest_status.status if latest_status else 'belum_ada',
                            'status_display': STATUS_OPTIONS.get(latest_status.status, 'Belum Ada') if latest_status else 'Belum Ada',
                            'catatan': latest_status.catatan if latest_status and latest_status.catatan else '-'
                        })
                    
                    # ===== AMBIL SEMUA UPLOAD (TANPA LIMIT) =====
                    uploads_in_bulan = UploadGambar.query\
                        .filter_by(stasiun_id=station.id)\
                        .filter(extract('year', UploadGambar.uploaded_at) == tahun)\
                        .filter(extract('month', UploadGambar.uploaded_at) == bulan)\
                        .order_by(UploadGambar.uploaded_at.desc())\
                        .all()
                    
                    uploads_by_group = {}
                    seen_uploads = set()
                    
                    for upload in uploads_in_bulan:
                        group_key = upload.group_id if upload.group_id is not None else 'ungrouped'
                        upload_key = f"{group_key}_{upload.public_id}"
                        
                        if upload_key not in seen_uploads:
                            seen_uploads.add(upload_key)
                            
                            if group_key not in uploads_by_group:
                                uploads_by_group[group_key] = []
                            
                            uploads_by_group[group_key].append(upload)
                    
                    all_group_keys = set(list(opponents_by_group.keys()) + list(uploads_by_group.keys()))
                    
                    sorted_group_keys = []
                    if 'ungrouped' in all_group_keys:
                        sorted_group_keys.append('ungrouped')
                    
                    numeric_keys = [k for k in all_group_keys if k != 'ungrouped' and k is not None]
                    numeric_keys.sort(key=lambda x: int(x) if str(x).isdigit() else x)
                    sorted_group_keys.extend(numeric_keys)
                    
                    group_counter = 0
                    for group_key in sorted_group_keys:
                        group_counter += 1
                        
                        if group_key == 'ungrouped':
                            group_display = 'Tanpa Grup'
                        else:
                            group_display = groups_dict.get(int(group_key), f'Grup {group_key}')
                        
                        print(f"    📁 Grup {group_counter}: {group_display}")
                        
                        doc.add_heading(f'  Grup {group_counter}: {group_display}', level=4)
                        
                        group_uploads = uploads_by_group.get(group_key, [])
                        
                        if group_uploads:
                            doc.add_heading('    GAMBAR:', level=5)
                            
                            for upload in group_uploads:
                                try:
                                    # ===== PERUBAHAN: Gunakan thumbnail =====
                                    thumbnail_url, _ = cloudinary_url(
                                        upload.public_id,
                                        width=500,
                                        height=375,
                                        crop="limit",
                                        quality="auto:eco"
                                    )
                                    response = requests.get(thumbnail_url, timeout=30)
                                    if response.status_code == 200:
                                        img_data = BytesIO(response.content)
                                        
                                        paragraph = doc.add_paragraph()
                                        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                                        run = paragraph.add_run()
                                        run.add_picture(img_data, width=Inches(5.0))
                                        
                                        info_para = doc.add_paragraph()
                                        info_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                                        info_text = f"File: {upload.original_filename} | "
                                        info_text += f"Upload: {upload.uploaded_at.strftime('%d/%m/%Y')}"
                                        info_run = info_para.add_run(info_text)
                                        info_run.italic = True
                                        info_run.font.size = Pt(9)
                                        
                                        doc.add_paragraph()
                                except Exception as e:
                                    print(f"Error download: {e}")
                                    doc.add_paragraph(f"    Gambar tidak tersedia: {upload.original_filename}")
                        else:
                            doc.add_paragraph("    Tidak ada gambar untuk grup ini")
                        
                        group_opponents = opponents_by_group.get(group_key, [])
                        if group_opponents:
                            doc.add_heading('    DAFTAR STASIUN LAWAN:', level=5)
                            
                            opp_table = doc.add_table(rows=1, cols=6)
                            opp_table.style = 'Table Grid'
                            opp_table.autofit = True
                            
                            opp_headers = ['No', 'Nama Stasiun Lawan', 'Freq Tx', 'Freq Rx', 'Status', 'Catatan']
                            for i, header in enumerate(opp_headers):
                                opp_table.rows[0].cells[i].text = header
                                opp_table.rows[0].cells[i].paragraphs[0].runs[0].font.bold = True
                                opp_table.rows[0].cells[i].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                            
                            for opp_idx, opponent in enumerate(group_opponents, 1):
                                row_cells = opp_table.add_row().cells
                                row_cells[0].text = str(opp_idx)
                                row_cells[0].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                                row_cells[1].text = opponent['nama']
                                row_cells[1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.LEFT
                                row_cells[2].text = opponent['freq_tx']
                                row_cells[2].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                                row_cells[3].text = opponent['freq_rx']
                                row_cells[3].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                                row_cells[4].text = opponent['status_display']
                                row_cells[4].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                                row_cells[5].text = opponent['catatan']
                                row_cells[5].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.LEFT
                            
                            doc.add_paragraph()
                        else:
                            doc.add_paragraph("    Tidak ada stasiun lawan dalam grup ini")
                        
                        doc.add_paragraph()
                    
                    doc.add_paragraph()
                
                if bulan != bulan_dengan_data[-1]:
                    doc.add_page_break()
        
        file_stream = BytesIO()
        doc.save(file_stream)
        file_stream.seek(0)
        
        filename = f"Rekap_Tahunan_{tahun}"
        if operator != 'all':
            filename += f"_{operator.upper()}"
        if kota != 'all':
            filename += f"_{kota.replace(' ', '_')}"
        filename += ".docx"
        
        print(f"\n✅ File berhasil dibuat: {filename}")
        print(f"   Tahun: {tahun}")
        print(f"   Bulan dengan data: {', '.join([bulan_nama[b-1] for b in bulan_dengan_data])}")
        print('='*80)
        
        return send_file(
            file_stream,
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            as_attachment=True,
            download_name=filename
        )
        
    except Exception as e:
        flash(f'Error: {str(e)}', 'error')
        print(f"\n❌ ERROR in admin_master_rekap_tahunan_word_gambar: {str(e)}")
        import traceback
        traceback.print_exc()
        return redirect(url_for('admin_master_rekap_tahunan', 
                               tahun=request.args.get('tahun', datetime.now().year, type=int),
                               operator=request.args.get('operator', 'all'),
                               kota=request.args.get('kota', 'all')))

# ============================================================================
# USER OPERATOR ROUTES 
# ============================================================================

@app.route('/user-operator/dashboard')
@login_required
@user_operator_required
def user_operator_dashboard():
    operator = current_user.operator_type
    
    total_stations = Stasiun.query.filter_by(operator=operator).count()
    
    total_opponents = StasiunLawan.query\
        .join(Stasiun)\
        .filter(Stasiun.operator == operator)\
        .count()
    
    uploads_count = db.session.query(db.func.count(db.distinct(UploadGambar.public_id)))\
        .join(Stasiun)\
        .filter(Stasiun.operator == operator)\
        .scalar() or 0
    
    recent_uploads_subquery = db.session.query(
        UploadGambar.public_id,
        db.func.max(UploadGambar.uploaded_at).label('max_uploaded_at')
    ).join(Stasiun)\
     .filter(Stasiun.operator == operator)\
     .group_by(UploadGambar.public_id)\
     .order_by(db.desc('max_uploaded_at'))\
     .limit(5).subquery()
    
    recent_uploads = UploadGambar.query\
        .join(recent_uploads_subquery,
            db.and_(
                UploadGambar.public_id == recent_uploads_subquery.c.public_id,
                UploadGambar.uploaded_at == recent_uploads_subquery.c.max_uploaded_at
            )
        )\
        .order_by(UploadGambar.uploaded_at.desc())\
        .all()
    
    stations_without_upload = []
    all_stations = Stasiun.query.filter_by(operator=operator).all()
    for station in all_stations:
        has_upload = UploadGambar.query.filter_by(stasiun_id=station.id).first()
        if not has_upload:
            stations_without_upload.append(station)
        if len(stations_without_upload) >= 5:
            break
    
    # ===== PERBAIKAN: Ambil daftar kota case-insensitive =====
    from sqlalchemy import func
    
    kota_raw = db.session.query(
        func.lower(Stasiun.kota).label('kota_lower'),
        Stasiun.kota
    ).filter_by(operator=operator)\
     .distinct(func.lower(Stasiun.kota))\
     .order_by(func.lower(Stasiun.kota))\
     .all()
    
    # Mapping untuk menyimpan format standar
    kota_mapping = {}
    for kota_lower, kota_original in kota_raw:
        if kota_lower not in kota_mapping:
            # Cari format standar dari KOTA_LIST
            found = False
            for std_kota in KOTA_LIST:
                if std_kota.lower() == kota_lower:
                    kota_mapping[kota_lower] = std_kota
                    found = True
                    break
            if not found:
                kota_mapping[kota_lower] = kota_original.strip().title()
    
    kota_list = list(kota_mapping.values())
    
    # Jika tidak ada kota di database, gunakan KOTA_LIST sebagai fallback
    if not kota_list:
        kota_list = KOTA_LIST
    
    current_date = datetime.now().strftime('%d %B %Y')
    
    return render_template('user_operator/dashboard.html',
                         operator=operator,
                         operator_name=operator.upper(),
                         current_date=current_date,
                         total_stations=total_stations,
                         total_uploads=uploads_count,
                         total_opponents=total_opponents,
                         recent_uploads=recent_uploads,
                         stations_without_upload=stations_without_upload,
                         kota_list=kota_list,
                         status_options=STATUS_OPTIONS)

def _handle_user_operator_stasiun(template_name='user_operator/stasiun.html'):
    """Daftar stasiun dengan filter kelengkapan status lawan (2 opsi) - OPTIMIZED"""
    operator = current_user.operator_type
    kota = request.args.get('kota', 'all')
    search_stasiun = request.args.get('search_stasiun', '').strip()
    status_lengkap = request.args.get('status_lengkap', 'all')
    
    # ===== PARAMETER PAGINATION =====
    per_page = request.args.get('per_page', 20, type=int)
    if per_page not in [10, 20, 50, 100]:
        per_page = 20
    page = request.args.get('page', 1, type=int)
    
    # ===== QUERY DASAR DENGAN PAGINATION =====
    query = Stasiun.query.filter_by(operator=operator)
    
    if search_stasiun:
        query = query.filter(Stasiun.stasiun_name.ilike(f'%{search_stasiun}%'))
    
    if kota != 'all' and kota:
        query = query.filter(Stasiun.kota.ilike(kota))
    
    # ===== PAGINATION DI DATABASE LEVEL =====
    pagination = query.order_by(Stasiun.stasiun_name).paginate(
        page=page, per_page=per_page, error_out=False
    )
    stations = pagination.items
    
    if not stations:
        # Handle empty case
        from sqlalchemy import func
        kota_raw = db.session.query(
            func.lower(Stasiun.kota).label('kota_lower'),
            Stasiun.kota
        ).filter_by(operator=operator).distinct(func.lower(Stasiun.kota)).order_by(func.lower(Stasiun.kota)).all()
        
        kota_mapping = {}
        for kota_lower, kota_original in kota_raw:
            if kota_lower not in kota_mapping:
                found = False
                for std_kota in KOTA_LIST:
                    if std_kota.lower() == kota_lower:
                        kota_mapping[kota_lower] = std_kota
                        found = True
                        break
                if not found:
                    kota_mapping[kota_lower] = kota_original.strip().title()
        
        kota_list_operator = list(kota_mapping.values())
        if not kota_list_operator:
            kota_list_operator = KOTA_LIST
        
        return render_template(template_name,
                             stations=[],
                             pagination=pagination,
                             operator=operator,
                             operator_name=operator.upper(),
                             kota_list=kota_list_operator,
                             selected_kota=kota,
                             selected_status_lengkap=status_lengkap,
                             status_summary={'lengkap': 0, 'belum_lengkap': 0})
    
    # ===== BATCH QUERY UNTUK SEMUA STASIUN DALAM 1 HALAMAN =====
    station_ids = [s.id for s in stations]
    
    # 1. Batch query semua stasiun lawan (1 query)
    all_opponents = StasiunLawan.query\
        .filter(StasiunLawan.stasiun_id.in_(station_ids))\
        .order_by(StasiunLawan.urutan)\
        .all()
    
    # 2. Batch query semua status terbaru (1 query)
    opponent_ids = [o.id for o in all_opponents]
    latest_status_dict = {}
    
    if opponent_ids:
        from sqlalchemy import func
        subquery = db.session.query(
            StatusUpdate.stasiun_lawan_id,
            func.max(StatusUpdate.updated_at).label('max_updated')
        ).filter(StatusUpdate.stasiun_lawan_id.in_(opponent_ids))\
         .group_by(StatusUpdate.stasiun_lawan_id)\
         .subquery()
        
        latest_statuses = db.session.query(StatusUpdate)\
            .join(subquery,
                  (StatusUpdate.stasiun_lawan_id == subquery.c.stasiun_lawan_id) &
                  (StatusUpdate.updated_at == subquery.c.max_updated))\
            .all()
        
        for status in latest_statuses:
            latest_status_dict[status.stasiun_lawan_id] = status
    
    # 3. Batch query semua grup (1 query)
    all_groups = GrupStasiun.query.filter(GrupStasiun.stasiun_id.in_(station_ids)).all()
    groups_by_station = {}
    for group in all_groups:
        if group.stasiun_id not in groups_by_station:
            groups_by_station[group.stasiun_id] = []
        groups_by_station[group.stasiun_id].append(group)
    
    # 4. Batch query user untuk checked_by (1 query)
    user_ids = [s.checked_by for s in stations if s.checked_by]
    users_dict = {}
    if user_ids:
        users = User.query.filter(User.id.in_(user_ids)).all()
        users_dict = {u.id: u for u in users}
    
    # ===== Kelompokkan lawan per stasiun =====
    opponents_by_station = {}
    for opponent in all_opponents:
        if opponent.stasiun_id not in opponents_by_station:
            opponents_by_station[opponent.stasiun_id] = []
        opponents_by_station[opponent.stasiun_id].append(opponent)
    
    # ===== Proses setiap stasiun =====
    for station in stations:
        station.stasiun_lawan_list = opponents_by_station.get(station.id, [])
        station.grup_list = groups_by_station.get(station.id, [])
        station.checked_by_user = users_dict.get(station.checked_by)
        
        # Hitung status untuk setiap lawan
        total_lawan = len(station.stasiun_lawan_list)
        lawan_tanpa_status = 0
        
        for lawan in station.stasiun_lawan_list:
            latest_status = latest_status_dict.get(lawan.id)
            lawan.latest_status = latest_status
            # Untuk template, kita perlu status_updates sebagai list (untuk sort)
            if latest_status:
                lawan.status_updates = [latest_status]
            else:
                lawan.status_updates = []
            
            if not latest_status:
                lawan_tanpa_status += 1
        
        station.has_lawan_without_status = lawan_tanpa_status > 0
    
    # ===== Filter berdasarkan status_lengkap =====
    if status_lengkap != 'all':
        filtered_stations = []
        for station in stations:
            total_lawan = len(station.stasiun_lawan_list)
            lawan_tanpa_status = sum(1 for l in station.stasiun_lawan_list if not l.latest_status)
            is_lengkap = (lawan_tanpa_status == 0 and total_lawan > 0)
            
            if status_lengkap == 'lengkap' and is_lengkap:
                filtered_stations.append(station)
            elif status_lengkap == 'belum_lengkap' and not is_lengkap and total_lawan > 0:
                filtered_stations.append(station)
        
        stations = filtered_stations
        pagination.items = stations
        pagination.total = len(filtered_stations)
    
    # ===== Hitung status summary untuk badge =====
    status_summary = {'lengkap': 0, 'belum_lengkap': 0}
    for station in stations:
        total_lawan = len(station.stasiun_lawan_list)
        lawan_tanpa_status = sum(1 for l in station.stasiun_lawan_list if not l.latest_status)
        if lawan_tanpa_status == 0 and total_lawan > 0:
            status_summary['lengkap'] += 1
        elif total_lawan > 0:
            status_summary['belum_lengkap'] += 1
    
    # ===== Ambil daftar kota untuk filter =====
    from sqlalchemy import func
    
    kota_raw = db.session.query(
        func.lower(Stasiun.kota).label('kota_lower'),
        Stasiun.kota
    ).filter_by(operator=operator)\
     .distinct(func.lower(Stasiun.kota))\
     .order_by(func.lower(Stasiun.kota))\
     .all()
    
    kota_mapping = {}
    for kota_lower, kota_original in kota_raw:
        if kota_lower not in kota_mapping:
            found = False
            for std_kota in KOTA_LIST:
                if std_kota.lower() == kota_lower:
                    kota_mapping[kota_lower] = std_kota
                    found = True
                    break
            if not found:
                kota_mapping[kota_lower] = kota_original.strip().title()
    
    kota_list_operator = list(kota_mapping.values())
    if not kota_list_operator:
        kota_list_operator = KOTA_LIST
    
    selected_kota_normalized = kota
    if kota != 'all':
        found = False
        for k in kota_list_operator:
            if k.lower() == kota.lower():
                selected_kota_normalized = k
                found = True
                break
        if not found:
            selected_kota_normalized = 'all'
    
    return render_template(template_name,
                         stations=stations,
                         pagination=pagination,
                         operator=operator,
                         operator_name=operator.upper(),
                         kota_list=kota_list_operator,
                         selected_kota=selected_kota_normalized,
                         selected_status_lengkap=status_lengkap,
                         status_summary=status_summary)

@app.route('/user-operator/stasiun')
@login_required
@user_operator_required
def user_operator_stasiun():
    return _handle_user_operator_stasiun('user_operator/stasiun.html')

@app.route('/user-operator/stasiun-v2')
@app.route('/user-operator/daftar-stasiun-v2')
@login_required
@user_operator_required
def user_operator_stasiun_v2():
    return _handle_user_operator_stasiun('user_operator/stasiun_v2.html')

@app.route('/user-operator/stasiun/<int:station_id>')
@login_required
@user_operator_required
def user_operator_detail_stasiun(station_id):
    """Detail stasiun dengan tampilan per grup dan status per lawan"""
    station = Stasiun.query.get_or_404(station_id)
    
    if station.operator != current_user.operator_type:
        flash('Stasiun ini bukan milik operator Anda!', 'error')
        return redirect(url_for('user_operator_dashboard'))
    
    # Ambil parameter next untuk kembali ke halaman sebelumnya
    next_url = request.args.get('next', '')
    
    all_lawans = StasiunLawan.query\
        .filter_by(stasiun_id=station_id)\
        .order_by(StasiunLawan.urutan)\
        .all()
    
    # AMBIL SEMUA UPLOAD UNIK (group by public_id)
    uploads_query = db.session.query(
        UploadGambar.group_id,
        UploadGambar.public_id,
        db.func.max(UploadGambar.id).label('latest_id')
    ).filter_by(stasiun_id=station_id)\
     .group_by(UploadGambar.group_id, UploadGambar.public_id)\
     .subquery()
    
    all_uploads = UploadGambar.query\
        .join(uploads_query, UploadGambar.id == uploads_query.c.latest_id)\
        .order_by(UploadGambar.uploaded_at.desc())\
        .all()
    
    all_groups = GrupStasiun.query.filter_by(stasiun_id=station_id).all()
    groups_dict = {g.id: g.nama_grup for g in all_groups}
    
    groups_data = {}
    
    # === PROSES LAWAN PER GRUP ===
    for lawan in all_lawans:
        group_key = lawan.group_id if lawan.group_id is not None else 'ungrouped'
        
        if group_key not in groups_data:
            if group_key == 'ungrouped':
                label = 'Tanpa Grup'
            else:
                label = groups_dict.get(int(group_key), f'Grup {group_key}')
                
            groups_data[group_key] = {
                'group_id': lawan.group_id,
                'label': label,
                'lawans': [],
                'uploads': [],
                'status_summary': {'aktif': 0, 'tidak_aktif': 0, 'tidak_sesuai': 0}
            }
        
        latest_status = StatusUpdate.query\
            .filter_by(stasiun_lawan_id=lawan.id)\
            .order_by(StatusUpdate.updated_at.desc())\
            .first()
        
        # Ambil informasi updater
        updater_info = None
        if latest_status and latest_status.updated_by:
            updater = db.session.get(User, latest_status.updated_by)
            updater_info = updater.username if updater else None
        
        lawan_data = build_stasiun_lawan_dict(
            lawan, station, latest_status, updater_info=updater_info
        )
        
        groups_data[group_key]['lawans'].append(lawan_data)
        
        if latest_status and latest_status.status in groups_data[group_key]['status_summary']:
            groups_data[group_key]['status_summary'][latest_status.status] += 1
    
    # === PROSES UPLOAD PER GRUP ===
    for upload in all_uploads:
        group_key = upload.group_id if upload.group_id is not None else 'ungrouped'
        
        if group_key not in groups_data:
            if group_key == 'ungrouped':
                label = 'Tanpa Grup'
            else:
                label = groups_dict.get(int(group_key), f'Grup {group_key}')
                
            groups_data[group_key] = {
                'group_id': upload.group_id,
                'label': label,
                'lawans': [],
                'uploads': [],
                'status_summary': {'aktif': 0, 'tidak_aktif': 0, 'tidak_sesuai': 0}
            }
        
        # Cek duplikasi public_id dalam grup yang sama
        existing_public_ids = [u.get('public_id') for u in groups_data[group_key]['uploads']]
        if upload.public_id not in existing_public_ids:
            uploader = db.session.get(User, upload.uploaded_by) if upload.uploaded_by else None
            
            upload_data = {
                'id': upload.id,
                'public_id': upload.public_id,
                'cloudinary_url': upload.cloudinary_url,
                'original_filename': upload.original_filename,
                'uploaded_at': upload.uploaded_at,
                'uploaded_by': uploader.username if uploader else 'System',
                'status': upload.status,
                'format': upload.format,
                'width': upload.width,
                'height': upload.height,
                'bytes_size': upload.bytes_size,
                'size_mb': round(upload.bytes_size / (1024 * 1024), 2) if upload.bytes_size else None
            }
            groups_data[group_key]['uploads'].append(upload_data)
    
    # Urutkan grup: ungrouped dulu, lalu grup lainnya
    sorted_groups = {}
    if 'ungrouped' in groups_data:
        sorted_groups['ungrouped'] = groups_data.pop('ungrouped')
    
    numeric_groups = {k: v for k, v in groups_data.items() if k != 'ungrouped'}
    for group_key in sorted(numeric_groups.keys(), key=lambda x: int(x) if str(x).isdigit() else x):
        sorted_groups[group_key] = numeric_groups[group_key]
    
    # Hitung status counts total
    status_counts = {'aktif': 0, 'tidak_aktif': 0, 'tidak_sesuai': 0}
    for lawan in all_lawans:
        latest_status = StatusUpdate.query\
            .filter_by(stasiun_lawan_id=lawan.id)\
            .order_by(StatusUpdate.updated_at.desc())\
            .first()
        
        if latest_status and latest_status.status in status_counts:
            status_counts[latest_status.status] += 1
    
    # Hitung statistik upload
    total_uploads = len(all_uploads)
    total_upload_size = sum([u.bytes_size for u in all_uploads if u.bytes_size]) if all_uploads else 0
    total_upload_size_mb = round(total_upload_size / (1024 * 1024), 2) if total_upload_size else 0
    
    # Hitung statistik grup
    total_groups = len(sorted_groups)
    groups_with_uploads = sum(1 for g in sorted_groups.values() if g['uploads'])
    groups_without_uploads = total_groups - groups_with_uploads
    
    return render_template('user_operator/detail_stasiun.html',
                         station=station,
                         groups=sorted_groups,
                         total_opponents=len(all_lawans),
                         status_counts=status_counts,
                         all_uploads=all_uploads,
                         total_uploads=total_uploads,
                         total_upload_size_mb=total_upload_size_mb,
                         total_groups=total_groups,
                         groups_with_uploads=groups_with_uploads,
                         groups_without_uploads=groups_without_uploads,
                         next_url=next_url,
                         STATUS_OPTIONS=STATUS_OPTIONS)

@app.route('/user-operator/edit-group/<int:station_id>', methods=['POST'])
@login_required
@user_operator_required
def user_operator_edit_group(station_id):
    """Edit grup"""
    station = Stasiun.query.get_or_404(station_id)
    
    if station.operator != current_user.operator_type:
        flash('Akses ditolak!', 'error')
        return redirect(url_for('user_operator_grup_stasiun', station_id=station_id))
    
    try:
        group_id = request.form.get('group_id')
        new_group_name = request.form.get('new_group_name')
        add_opponents = request.form.getlist('add_opponents[]')
        move_opponents = request.form.getlist('move_opponents[]')
        
        for opponent_id in add_opponents:
            lawan = StasiunLawan.query.get(opponent_id)
            if lawan and lawan.stasiun_id == station_id:
                lawan.group_id = int(group_id)
        
        for opponent_id in move_opponents:
            lawan = StasiunLawan.query.get(opponent_id)
            if lawan and lawan.stasiun_id == station_id:
                new_group_key = request.form.get(f'new_group_{opponent_id}')
                if new_group_key == 'ungrouped':
                    lawan.group_id = None
                elif new_group_key:
                    lawan.group_id = int(new_group_key)
        
        db.session.commit()
        flash('Perubahan grup berhasil disimpan!', 'success')
        
    except Exception as e:
        db.session.rollback()
        flash(f'Error: {str(e)}', 'error')
    
    return redirect(url_for('user_operator_grup_stasiun', station_id=station_id))

@app.route('/user-operator/edit-lawan-single/<int:station_id>', methods=['POST'])
@login_required
@user_operator_required
def user_operator_edit_lawan_single(station_id):
    """Edit satu stasiun lawan"""
    station = Stasiun.query.get_or_404(station_id)
    
    if station.operator != current_user.operator_type:
        flash('Akses ditolak!', 'error')
        return redirect(url_for('user_operator_grup_stasiun', station_id=station_id))
    
    try:
        lawan_id = request.form.get('lawan_id')
        nama_lawan = request.form.get('nama_lawan')
        new_group_id = request.form.get('new_group_id')
        
        lawan = StasiunLawan.query.get(lawan_id)
        if lawan and lawan.stasiun_id == station_id:
            lawan.nama_stasiun_lawan = nama_lawan.strip()
            if new_group_id:
                lawan.group_id = int(new_group_id) if new_group_id != '' else None
        
        db.session.commit()
        flash('Data lawan berhasil diperbarui!', 'success')
        
    except Exception as e:
        db.session.rollback()
        flash(f'Error: {str(e)}', 'error')
    
    return redirect(url_for('user_operator_grup_stasiun', station_id=station_id))

@app.route('/user-operator/delete-lawan/<int:station_id>', methods=['POST'])
@login_required
@user_operator_required
def user_operator_delete_lawan(station_id):
    """Hapus stasiun lawan (untuk yang tanpa grup)"""
    station = Stasiun.query.get_or_404(station_id)
    
    if station.operator != current_user.operator_type:
        flash('Akses ditolak!', 'error')
        return redirect(url_for('user_operator_grup_stasiun', station_id=station_id))
    
    try:
        lawan_id = request.form.get('lawan_id')
        lawan = StasiunLawan.query.get(lawan_id)
        
        if lawan and lawan.stasiun_id == station_id:
            StatusUpdate.query.filter_by(stasiun_lawan_id=lawan.id).delete()
            
            uploads = UploadGambar.query.filter_by(stasiun_lawan_id=lawan.id).all()
            for upload in uploads:
                delete_from_cloudinary(upload.public_id)
                db.session.delete(upload)
            
            db.session.delete(lawan)
            db.session.commit()
            
            flash(f'Stasiun lawan "{lawan.nama_stasiun_lawan}" berhasil dihapus!', 'success')
        else:
            flash('Stasiun lawan tidak ditemukan', 'error')
            
    except Exception as e:
        db.session.rollback()
        flash(f'Error: {str(e)}', 'error')
    
    return redirect(url_for('user_operator_grup_stasiun', station_id=station_id))

@app.route('/user-operator/delete-lawan-from-group/<int:station_id>', methods=['POST'])
@login_required
@user_operator_required
def user_operator_delete_lawan_from_group(station_id):
    """
    KELUARKAN stasiun lawan dari grup - PINDAHKAN KE TANPA GRUP
    BUKAN DIHAPUS PERMANEN!
    """
    station = Stasiun.query.get_or_404(station_id)
    
    if station.operator != current_user.operator_type:
        flash('Akses ditolak!', 'error')
        return redirect(url_for('user_operator_grup_stasiun', station_id=station_id))
    
    try:
        lawan_id = request.form.get('lawan_id')
        lawan = StasiunLawan.query.get(lawan_id)
        
        if lawan and lawan.stasiun_id == station_id:
            nama_lawan = lawan.nama_stasiun_lawan
            lawan.group_id = None
            
            UploadGambar.query.filter_by(
                stasiun_lawan_id=lawan.id,
                group_id=lawan.group_id
            ).update({'group_id': None})
            
            db.session.commit()
            
            flash(f'Stasiun lawan "{nama_lawan}" berhasil dikeluarkan dari grup dan dipindahkan ke "Tanpa Grup"!', 'success')
        else:
            flash('Stasiun lawan tidak ditemukan', 'error')
            
    except Exception as e:
        db.session.rollback()
        flash(f'Error: {str(e)}', 'error')
    
    return redirect(url_for('user_operator_grup_stasiun', station_id=station_id))

@app.route('/user-operator/grup-stasiun/<int:station_id>', methods=['GET', 'POST'])
@login_required
@user_operator_required
def user_operator_grup_stasiun(station_id):
    station = Stasiun.query.get_or_404(station_id)
    
    # Ambil parameter next untuk kembali ke halaman sebelumnya
    next_url = request.args.get('next', '')
    
    if station.operator != current_user.operator_type:
        flash('Stasiun ini bukan milik operator Anda!', 'error')
        return redirect(url_for('user_operator_dashboard'))
    
    if request.method == 'POST':
        try:
            action = request.form.get('action')
            
            # ===== 1. CREATE GROUP =====
            if action == 'create_group':
                group_name = request.form.get('group_name')
                selected_opponents = request.form.getlist('selected_opponents[]')
                
                if not group_name or not selected_opponents:
                    flash('Nama grup dan stasiun lawan harus dipilih!', 'error')
                    return redirect(url_for('user_operator_grup_stasiun', station_id=station_id))
                
                new_grup = GrupStasiun(
                    stasiun_id=station_id,
                    nama_grup=group_name
                )
                db.session.add(new_grup)
                db.session.flush()
                
                for opponent_id in selected_opponents:
                    lawan = StasiunLawan.query.get(opponent_id)
                    if lawan and lawan.stasiun_id == station_id:
                        lawan.group_id = new_grup.id
                        lawan.grup_id = new_grup.id
                
                db.session.commit()
                flash(f'Grup "{group_name}" berhasil dibuat dengan {len(selected_opponents)} lawan!', 'success')
                # TETAP DI HALAMAN GRUP STASIUN
                return redirect(url_for('user_operator_grup_stasiun', station_id=station_id))
                
            # ===== 2. DELETE GROUP =====
            elif action == 'delete_group':
                group_id = request.form.get('group_id')
                if group_id:
                    grup = GrupStasiun.query.get(group_id)
                    if grup and grup.stasiun_id == station_id:
                        # Pindahkan semua lawan ke ungrouped
                        StasiunLawan.query\
                            .filter_by(stasiun_id=station_id, grup_id=group_id)\
                            .update({
                                'group_id': None,
                                'grup_id': None
                            })
                        
                        # Update upload yang terkait
                        UploadGambar.query\
                            .filter_by(stasiun_id=station_id, group_id=int(group_id))\
                            .update({'group_id': None})
                        
                        db.session.delete(grup)
                        db.session.commit()
                        flash('Grup berhasil dihapus! Semua anggota dipindahkan ke "Tanpa Grup".', 'success')
                # TETAP DI HALAMAN GRUP STASIUN
                return redirect(url_for('user_operator_grup_stasiun', station_id=station_id))
            
            # ===== 3. REMOVE FROM GROUP (KELUARKAN DARI GRUP) =====
            elif action == 'remove_from_group':
                opponent_id = request.form.get('opponent_id')
                if opponent_id:
                    lawan = StasiunLawan.query.get(opponent_id)
                    if lawan and lawan.stasiun_id == station_id:
                        nama_lawan = lawan.nama_stasiun_lawan
                        
                        # Keluarkan dari grup
                        lawan.group_id = None
                        lawan.grup_id = None
                        
                        # Update upload yang terkait
                        UploadGambar.query.filter_by(
                            stasiun_lawan_id=lawan.id
                        ).update({'group_id': None})
                        
                        db.session.commit()
                        flash(f'Stasiun lawan "{nama_lawan}" berhasil dikeluarkan dari grup!', 'success')
                    else:
                        flash('Stasiun lawan tidak ditemukan!', 'error')
                # TETAP DI HALAMAN GRUP STASIUN
                return redirect(url_for('user_operator_grup_stasiun', station_id=station_id))
            
            # ===== 4. EDIT SINGLE OPPONENT (PINDAHKAN GRUP) =====
            elif action == 'edit_lawan':
                lawan_id = request.form.get('lawan_id')
                nama_lawan = request.form.get('nama_lawan')
                freq_tx = request.form.get('freq_tx', '').strip()
                freq_rx = request.form.get('freq_rx', '').strip()
                new_group_id = request.form.get('new_group_id')
                
                # Validasi input
                if not lawan_id or not nama_lawan:
                    flash('Data tidak lengkap!', 'error')
                    return redirect(url_for('user_operator_grup_stasiun', station_id=station_id))
                
                # Cari lawan
                lawan = StasiunLawan.query.get(lawan_id)
                if not lawan or lawan.stasiun_id != station_id:
                    flash('Stasiun lawan tidak ditemukan!', 'error')
                    return redirect(url_for('user_operator_grup_stasiun', station_id=station_id))
                
                # Catat perubahan untuk pesan
                changes = []
                
                # Update nama
                if lawan.nama_stasiun_lawan != nama_lawan:
                    changes.append(f"Nama: '{lawan.nama_stasiun_lawan}' → '{nama_lawan}'")
                    lawan.nama_stasiun_lawan = nama_lawan
                
                # Update freq_tx (boleh kosong)
                old_freq_tx = lawan.freq_tx or ''
                new_freq_tx = freq_tx if freq_tx else None
                if old_freq_tx != (new_freq_tx or ''):
                    if old_freq_tx and new_freq_tx:
                        changes.append(f"Freq Tx: '{old_freq_tx}' → '{new_freq_tx}'")
                    elif new_freq_tx and not old_freq_tx:
                        changes.append(f"Freq Tx ditambahkan: '{new_freq_tx}'")
                    elif old_freq_tx and not new_freq_tx:
                        changes.append(f"Freq Tx dihapus (sebelumnya '{old_freq_tx}')")
                    lawan.freq_tx = new_freq_tx
                
                # Update freq_rx (boleh kosong)
                old_freq_rx = lawan.freq_rx or ''
                new_freq_rx = freq_rx if freq_rx else None
                if old_freq_rx != (new_freq_rx or ''):
                    if old_freq_rx and new_freq_rx:
                        changes.append(f"Freq Rx: '{old_freq_rx}' → '{new_freq_rx}'")
                    elif new_freq_rx and not old_freq_rx:
                        changes.append(f"Freq Rx ditambahkan: '{new_freq_rx}'")
                    elif old_freq_rx and not new_freq_rx:
                        changes.append(f"Freq Rx dihapus (sebelumnya '{old_freq_rx}')")
                    lawan.freq_rx = new_freq_rx
                
                # Update grup
                old_group = lawan.group_id
                if new_group_id and new_group_id.strip():
                    new_group_int = int(new_group_id)
                    if old_group != new_group_int:
                        # Cari nama grup baru
                        grup_baru = GrupStasiun.query.get(new_group_int)
                        grup_baru_nama = grup_baru.nama_grup if grup_baru else f'Grup {new_group_int}'
                        
                        if old_group:
                            grup_lama = GrupStasiun.query.get(old_group)
                            grup_lama_nama = grup_lama.nama_grup if grup_lama else f'Grup {old_group}'
                            changes.append(f"Grup: '{grup_lama_nama}' → '{grup_baru_nama}'")
                        else:
                            changes.append(f"Grup: 'Tanpa Grup' → '{grup_baru_nama}'")
                        
                        lawan.group_id = new_group_int
                        lawan.grup_id = new_group_int
                else:
                    if old_group is not None:
                        grup_lama = GrupStasiun.query.get(old_group)
                        grup_lama_nama = grup_lama.nama_grup if grup_lama else f'Grup {old_group}'
                        changes.append(f"Grup: '{grup_lama_nama}' → 'Tanpa Grup'")
                        lawan.group_id = None
                        lawan.grup_id = None
                
                # Update upload yang terkait dengan group_id baru
                if 'group_id' in str(changes) or 'grup' in str(changes):
                    UploadGambar.query.filter_by(
                        stasiun_lawan_id=lawan.id
                    ).update({'group_id': lawan.group_id})
                
                # Commit perubahan
                db.session.commit()
                
                # Tampilkan pesan
                if changes:
                    flash(f'✅ Stasiun lawan "{nama_lawan}" berhasil diperbarui!', 'success')
                    for change in changes:
                        flash(f'  • {change}', 'info')
                else:
                    flash(f'ℹ️ Tidak ada perubahan pada "{nama_lawan}".', 'info')
                
                # TETAP DI HALAMAN GRUP STASIUN
                return redirect(url_for('user_operator_grup_stasiun', station_id=station_id))
            
            # Jika ada next_url, redirect ke sana
            if next_url:
                return redirect(next_url)
            else:
                return redirect(url_for('user_operator_grup_stasiun', station_id=station_id))
            
        except Exception as e:
            db.session.rollback()
            flash(f'❌ Error: {str(e)}', 'error')
            print(f"ERROR in user_operator_grup_stasiun: {str(e)}")
            import traceback
            traceback.print_exc()
            return redirect(url_for('user_operator_grup_stasiun', station_id=station_id))
    
    # ===== GET REQUEST - TAMPILKAN HALAMAN =====
    stasiun_lawan = StasiunLawan.query\
        .filter_by(stasiun_id=station_id)\
        .order_by(StasiunLawan.urutan)\
        .all()
    
    all_groups = GrupStasiun.query.filter_by(stasiun_id=station_id).all()
    groups_dict = {g.id: g for g in all_groups}
    
    groups_with_names = {}
    
    for lawan in stasiun_lawan:
        if lawan.grup_id:
            group_key = lawan.grup_id
        else:
            group_key = 'ungrouped'
        
        if group_key not in groups_with_names:
            if group_key == 'ungrouped':
                groups_with_names[group_key] = {
                    'nama': 'Tanpa Grup',
                    'lawans': []
                }
            else:
                grup_obj = groups_dict.get(group_key)
                groups_with_names[group_key] = {
                    'nama': grup_obj.nama_grup if grup_obj else f'Grup {group_key}',
                    'lawans': []
                }
        
        lawan_data = {
            'id': lawan.id,
            'nama_stasiun_lawan': lawan.nama_stasiun_lawan,
            'freq_tx': lawan.freq_tx,
            'freq_rx': lawan.freq_rx,
            'group_id': lawan.group_id
        }
        groups_with_names[group_key]['lawans'].append(lawan_data)
    
    ungrouped_lawans = StasiunLawan.query\
        .filter_by(stasiun_id=station_id, grup_id=None)\
        .order_by(StasiunLawan.urutan)\
        .all()
    
    return render_template('user_operator/grup_stasiun.html',
                         station=station,
                         stasiun_lawan=stasiun_lawan,
                         groups=groups_with_names,
                         ungrouped_lawans=ungrouped_lawans,
                         next_url=next_url)

@app.route('/user-operator/upload/<int:station_id>', methods=['GET', 'POST'])
@login_required
@user_operator_required
def user_operator_upload(station_id):
    station = Stasiun.query.get_or_404(station_id)
    
    if station.operator != current_user.operator_type:
        flash('Akses ditolak!', 'error')
        return redirect(url_for('user_operator_dashboard'))
    
    # ========== GET REQUEST - TAMPILKAN FORM ==========
    if request.method == 'GET':
        all_groups = GrupStasiun.query.filter_by(stasiun_id=station_id).all()
        station.grup_list = all_groups
        
        stasiun_lawan = StasiunLawan.query\
            .filter_by(stasiun_id=station_id)\
            .order_by(StasiunLawan.urutan)\
            .all()
        
        groups = {}
        for lawan in stasiun_lawan:
            group_key = lawan.group_id if lawan.group_id is not None else 'ungrouped'
            if group_key not in groups:
                groups[group_key] = []
            
            latest_status = StatusUpdate.query\
                .filter_by(stasiun_lawan_id=lawan.id)\
                .order_by(StatusUpdate.updated_at.desc())\
                .first()
            
            lawan_data = {
                'id': lawan.id,
                'nama': lawan.nama_stasiun_lawan,
                'freq_tx': lawan.freq_tx,
                'freq_rx': lawan.freq_rx,
                'latest_status': latest_status,
                'catatan': latest_status.catatan if latest_status else ''
            }
            groups[group_key].append(lawan_data)
        
        # ===== TAMBAHKAN: Cek status upload per grup =====
        group_uploads_status = get_group_upload_status(station_id)
        
        return render_template('user_operator/upload.html',
                             station=station,
                             groups=groups,
                             group_uploads_status=group_uploads_status,
                             status_options=STATUS_OPTIONS)
    
    # ========== POST REQUEST - UPLOAD PER GRUP ==========
    if request.method == 'POST':
        try:
            print("=" * 80)
            print("🚀 UPLOAD PER GRUP")
            print("=" * 80)
            
            # Ambil group_key dari form
            group_key = request.form.get('group_key')
            print(f"📁 Group key: {group_key}")
            
            if not group_key:
                flash('Group key tidak ditemukan!', 'error')
                return redirect(url_for('user_operator_upload', station_id=station_id))
            
            # Tentukan group_id
            if group_key == 'ungrouped':
                group_id = None
                group_label = 'Tanpa Grup'
            else:
                try:
                    group_id = int(group_key)
                    grup = GrupStasiun.query.filter_by(id=group_id, stasiun_id=station_id).first()
                    group_label = grup.nama_grup if grup else f'Grup {group_id}'
                except:
                    flash('Group key tidak valid!', 'error')
                    return redirect(url_for('user_operator_upload', station_id=station_id))
            
            # Ambil lawan dalam grup ini
            if group_id is None:
                lawans_in_group = StasiunLawan.query.filter_by(stasiun_id=station_id, group_id=None).all()
            else:
                lawans_in_group = StasiunLawan.query.filter_by(stasiun_id=station_id, group_id=group_id).all()
            
            if not lawans_in_group:
                flash(f'Tidak ada stasiun lawan dalam grup {group_label}', 'warning')
                return redirect(url_for('user_operator_upload', station_id=station_id))
            
            print(f"👥 Lawan dalam grup: {len(lawans_in_group)}")
            
            # ===== 1. AMBIL FILE YANG DIUPLOAD =====
            uploaded_files = []
            file_input_key = f'gambar_{group_key}'
            
            if file_input_key in request.files:
                files = request.files.getlist(file_input_key)
                print(f"📸 File yang diupload: {len(files)} file")
                
                for file in files:
                    if file and file.filename:
                        print(f"  - {file.filename}")
                        
                        # Validasi file
                        if not allowed_file(file.filename):
                            flash(f'File {file.filename} format tidak didukung!', 'warning')
                            continue
                        
                        # Validasi ukuran
                        file_content = file.read()
                        file_size = len(file_content)
                        file.seek(0)
                        
                        if file_size > 5 * 1024 * 1024:
                            flash(f'File {file.filename} terlalu besar (maks 5MB)!', 'warning')
                            continue
                        
                        # Upload ke Cloudinary
                        result, error = upload_to_cloudinary(file, station_id, None, group_id)
                        
                        if result:
                            uploaded_files.append({
                                'public_id': result['public_id'],
                                'url': result['url'],
                                'original_filename': result['original_filename'],
                                'width': result['width'],
                                'height': result['height'],
                                'format': result['format'],
                                'bytes': result['bytes']
                            })
                            print(f"  ✅ Uploaded: {result['public_id']}")
                        else:
                            flash(f'Gagal upload {file.filename}: {error}', 'error')
            else:
                print(f"⚠️ Tidak ada file yang diupload untuk grup {group_key}")
            
            if not uploaded_files:
                flash('Tidak ada file yang berhasil diupload!', 'warning')
                return redirect(url_for('user_operator_upload', station_id=station_id))
            
            # ===== 2. AMBIL STATUS DARI FORM =====
            status_map = {}
            for lawan in lawans_in_group:
                status_key = f'status_{group_key}_{lawan.id}'
                catatan_key = f'catatan_{group_key}_{lawan.id}'
                
                status_value = request.form.get(status_key)
                catatan_value = request.form.get(catatan_key, '')
                
                if status_value:
                    status_map[lawan.id] = {
                        'status': status_value,
                        'catatan': catatan_value
                    }
                    print(f"  Status for lawan {lawan.id}: {status_value}")
            
            # ===== 3. SIMPAN DATA =====
            upload_count = 0
            status_count = 0
            
            for lawan in lawans_in_group:
                status_info = status_map.get(lawan.id)
                
                if not status_info:
                    print(f"⚠️ No status for lawan {lawan.id}, skipping...")
                    continue
                
                # Simpan status update
                new_status = StatusUpdate(
                    stasiun_lawan_id=lawan.id,
                    status=status_info['status'],
                    updated_by=current_user.id,
                    catatan=status_info['catatan']
                )
                db.session.add(new_status)
                status_count += 1
                
                # Simpan upload records untuk setiap file
                for file_data in uploaded_files:
                    upload = UploadGambar(
                        stasiun_id=station_id,
                        stasiun_lawan_id=lawan.id,
                        group_id=group_id,
                        # Cloudinary data
                        public_id=file_data['public_id'],
                        cloudinary_url=file_data['url'],
                        original_filename=file_data['original_filename'],
                        width=file_data['width'],
                        height=file_data['height'],
                        format=file_data['format'],
                        bytes_size=file_data['bytes'],
                        # Other data
                        status=status_info['status'],
                        uploaded_by=current_user.id,
                        is_checked=True
                    )
                    db.session.add(upload)
                    upload_count += 1
                
                # Update status di semua upload yang sudah ada
                UploadGambar.query.filter_by(
                    stasiun_lawan_id=lawan.id
                ).update({'status': status_info['status']})
            
            db.session.commit()
            
            print(f"\n📊 SUMMARY:")
            print(f"  - Uploaded files: {len(uploaded_files)}")
            print(f"  - Upload records: {upload_count}")
            print(f"  - Status updates: {status_count}")
            print("=" * 80)
            
            flash(f'✅ Berhasil upload {len(uploaded_files)} file untuk grup {group_label}! {status_count} status diperbarui.', 'success')
            
        except Exception as e:
            db.session.rollback()
            print(f"\n❌ ERROR: {str(e)}")
            import traceback
            traceback.print_exc()
            flash(f'Error: {str(e)}', 'error')
        
        return redirect(url_for('user_operator_detail_stasiun', station_id=station_id))

@app.route('/user-operator/delete-opponent/<int:station_id>/<int:opponent_id>', methods=['POST'])
@login_required
@user_operator_required
def user_operator_delete_opponent(station_id, opponent_id):
    """Hapus stasiun lawan dari database"""
    station = Stasiun.query.get_or_404(station_id)
    
    if station.operator != current_user.operator_type:
        flash('Akses ditolak!', 'error')
        return redirect(url_for('user_operator_dashboard'))
    
    try:
        lawan = StasiunLawan.query.get_or_404(opponent_id)
        
        StatusUpdate.query.filter_by(stasiun_lawan_id=opponent_id).delete()
        
        uploads = UploadGambar.query.filter_by(stasiun_lawan_id=opponent_id).all()
        for upload in uploads:
            delete_from_cloudinary(upload.public_id)
            db.session.delete(upload)
        
        db.session.delete(lawan)
        db.session.commit()
        
        flash('Stasiun lawan berhasil dihapus!', 'success')
    except Exception as e:
        db.session.rollback()
        flash(f'Error: {str(e)}', 'error')
    
    return redirect(url_for('user_operator_grup_stasiun', station_id=station_id))

@app.route('/user-operator/edit-grup/<int:station_id>/<group_key>', methods=['GET', 'POST'])
@login_required
@user_operator_required
def user_operator_edit_grup(station_id, group_key):
    """Edit grup - UPDATE UNTUK CLOUDINARY"""
    station = Stasiun.query.get_or_404(station_id)
    
    if station.operator != current_user.operator_type:
        flash('Akses ditolak!', 'error')
        return redirect(url_for('user_operator_dashboard'))
    
    if group_key == 'ungrouped':
        group_id = None
        group_label = 'Tanpa Grup'
    else:
        try:
            group_id = int(group_key)
            grup = GrupStasiun.query.filter_by(id=group_id, stasiun_id=station_id).first()
            if grup:
                group_label = grup.nama_grup
            else:
                group_label = f'Grup {group_id}'
        except:
            flash('Group key tidak valid!', 'error')
            return redirect(url_for('user_operator_detail_stasiun', station_id=station_id))
    
    if group_id is None:
        lawans_in_group = StasiunLawan.query.filter_by(stasiun_id=station_id, group_id=None).all()
    else:
        lawans_in_group = StasiunLawan.query.filter_by(stasiun_id=station_id, group_id=group_id).all()
    
    if group_id is None:
        uploads_sub = db.session.query(
            UploadGambar.public_id, 
            db.func.max(UploadGambar.id).label('id')
        ).filter(
            UploadGambar.stasiun_id == station_id, 
            UploadGambar.group_id.is_(None)
        ).group_by(UploadGambar.public_id).subquery()
    else:
        uploads_sub = db.session.query(
            UploadGambar.public_id, 
            db.func.max(UploadGambar.id).label('id')
        ).filter_by(
            stasiun_id=station_id, 
            group_id=group_id
        ).group_by(UploadGambar.public_id).subquery()
    
    current_uploads = UploadGambar.query.join(
        uploads_sub, UploadGambar.id == uploads_sub.c.id
    ).order_by(UploadGambar.uploaded_at.desc()).all()
    
    if request.method == 'POST':
        try:
            print("\n" + "="*80)
            print("🚀 EDIT GRUP - CLOUDINARY")
            print("="*80)
            
            # ===== DEBUG: CEK SEMUA FILE DI REQUEST =====
            print(f"\n📋 REQUEST FILES:")
            print(f"  request.files keys: {list(request.files.keys())}")
            
            files = request.files.getlist('new_files[]')
            print(f"  files from 'new_files[]': {len(files)}")
            
            for i, file in enumerate(files):
                if file and file.filename:
                    print(f"    File {i}: {file.filename} - {file.content_type} - {file.content_length} bytes")
            
            # ===== DEBUG: CEK SEMUA FORM DATA =====
            print(f"\n📋 FORM DATA:")
            for key, value in request.form.items():
                if not key.startswith('status_') and not key.startswith('catatan_'):
                    print(f"  {key}: {value}")
            
            delete_files = request.form.getlist('delete_files[]')
            print(f"\n🗑️ Files to delete: {delete_files}")
            
            # ===== 1. HAPUS FILE =====
            deleted_count = 0
            for upload_id in delete_files:
                try:
                    upload_id_int = int(upload_id)
                    upload = UploadGambar.query.get(upload_id_int)
                    
                    if upload and upload.stasiun_id == station_id:
                        public_id = upload.public_id
                        print(f"  Deleting: {public_id}")
                        
                        # Cari semua record dengan public_id yang sama dalam grup ini
                        deleted_records = UploadGambar.query.filter_by(
                            stasiun_id=station_id,
                            public_id=public_id,
                            group_id=group_id
                        ).delete(synchronize_session=False)
                        
                        print(f"    - {deleted_records} records deleted")
                        
                        # Cek apakah masih ada record lain dengan public_id yang sama
                        other_refs = UploadGambar.query.filter_by(public_id=public_id).count()
                        if other_refs == 0:
                            success, msg = delete_from_cloudinary(public_id)
                            if success:
                                print(f"    - Cloudinary file deleted")
                        
                        deleted_count += deleted_records
                except Exception as e:
                    print(f"    - Error: {e}")
            
            # ===== 2. KOLEKSI STATUS DARI FORM =====
            status_map = {}
            for lawan in lawans_in_group:
                status_value = request.form.get(f'status_{lawan.id}')
                catatan_value = request.form.get(f'catatan_{lawan.id}', '')
                if status_value:
                    status_map[lawan.id] = {
                        'status': status_value,
                        'catatan': catatan_value
                    }
                    print(f"  Status for lawan {lawan.id}: {status_value}")
            
            # ===== 3. UPLOAD FILE BARU - CARA ALTERNATIF =====
            uploaded_count = 0
            
            # Coba cari file dengan cara langsung
            for key in request.files:
                file_list = request.files.getlist(key)
                print(f"  Processing key '{key}' with {len(file_list)} files")
                
                for file in file_list:
                    if file and file.filename:
                        print(f"\n  📤 Processing: {file.filename}")
                        
                        # Validasi file
                        if not allowed_file(file.filename):
                            print(f"    - ❌ Format tidak didukung: {file.filename}")
                            flash(f'File {file.filename} format tidak didukung!', 'warning')
                            continue
                        
                        # Baca file untuk validasi ukuran
                        file_content = file.read()
                        file_size = len(file_content)
                        file.seek(0)  # Reset pointer
                        
                        if file_size > 5 * 1024 * 1024:
                            print(f"    - ❌ File terlalu besar: {file_size} bytes")
                            flash(f'File {file.filename} terlalu besar (maks 5MB)!', 'warning')
                            continue
                        
                        print(f"    - ✅ File valid: {file.filename}, size: {file_size} bytes")
                        
                        # Upload ke Cloudinary
                        result, error = upload_to_cloudinary(file, station_id, None, group_id)
                        
                        if result:
                            print(f"    - ✅ File uploaded ke Cloudinary: {result['public_id']}")
                            
                            # Untuk setiap lawan dalam grup, buat record upload
                            for lawan in lawans_in_group:
                                # Ambil status lawan ini dari form
                                status_info = status_map.get(lawan.id, {})
                                current_status = status_info.get('status')
                                current_catatan = status_info.get('catatan', '')
                                
                                upload = UploadGambar(
                                    stasiun_id=station_id,
                                    stasiun_lawan_id=lawan.id,
                                    group_id=group_id,
                                    # Data Cloudinary
                                    public_id=result['public_id'],
                                    cloudinary_url=result['url'],
                                    original_filename=result['original_filename'],
                                    width=result['width'],
                                    height=result['height'],
                                    format=result['format'],
                                    bytes_size=result['bytes'],
                                    # Data lain
                                    status=current_status,
                                    uploaded_by=current_user.id,
                                    is_checked=True
                                )
                                db.session.add(upload)
                                uploaded_count += 1
                                print(f"    - Created upload for lawan {lawan.id} with status '{current_status}'")
                        else:
                            print(f"    - ❌ Failed to upload: {error}")
                            flash(f'Gagal upload {file.filename}: {error}', 'error')
            
            # ===== 4. UPDATE STATUS =====
            status_count = 0
            for lawan in lawans_in_group:
                status_info = status_map.get(lawan.id)
                
                if status_info:
                    print(f"\n  Updating status for lawan {lawan.id}: {status_info['status']}")
                    
                    # Buat status update baru
                    status_update = StatusUpdate(
                        stasiun_lawan_id=lawan.id,
                        status=status_info['status'],
                        updated_by=current_user.id,
                        catatan=status_info['catatan']
                    )
                    db.session.add(status_update)
                    status_count += 1
                    
                    # Update status di SEMUA upload yang terkait dengan lawan ini
                    updated = UploadGambar.query.filter_by(
                        stasiun_lawan_id=lawan.id
                    ).update({'status': status_info['status']})
                    print(f"    - Updated {updated} existing uploads")
            
            # ===== 5. COMMIT SEMUA PERUBAHAN =====
            db.session.commit()
            
            print("\n" + "="*80)
            print(f"✅ SUMMARY:")
            print(f"  - Deleted: {deleted_count} records")
            print(f"  - Uploaded: {uploaded_count} files")
            print(f"  - Status updated: {status_count}")
            print("="*80)
            
            if status_count > 0:
                flash(f'✅ {status_count} status diperbarui', 'success')
            
        except Exception as e:
            db.session.rollback()
            print(f"\n❌ ERROR: {str(e)}")
            import traceback
            traceback.print_exc()
            flash(f'Error: {str(e)}', 'error')
        
        return redirect(url_for('user_operator_detail_stasiun', station_id=station_id))
    
    # GET REQUEST - TAMPILKAN FORM
    opponents_data = []
    for lawan in lawans_in_group:
        latest = StatusUpdate.query.filter_by(
            stasiun_lawan_id=lawan.id
        ).order_by(StatusUpdate.updated_at.desc()).first()
        
        opponents_data.append({
            'id': lawan.id,
            'nama': lawan.nama_stasiun_lawan,
            'freq_tx': lawan.freq_tx,
            'freq_rx': lawan.freq_rx,
            'current_status': latest.status if latest else None,
            'status_display': STATUS_OPTIONS.get(latest.status, 'Belum Ada') if latest else 'Belum Ada',
            'catatan': latest.catatan if latest else ''
        })
    
    return render_template('user_operator/edit_grup.html',
                         station=station,
                         group_key=group_key,
                         group_label=group_label,
                         group_id=group_id,
                         current_uploads=current_uploads,
                         opponents=opponents_data,
                         status_options=STATUS_OPTIONS)

@app.route('/user-operator/hapus-upload/<int:upload_id>')
@login_required
@user_operator_required
def user_operator_hapus_upload(upload_id):
    """Hapus upload dan semua duplikatnya dari Cloudinary"""
    upload = UploadGambar.query.get_or_404(upload_id)
    station = Stasiun.query.get(upload.stasiun_id)
    
    if station.operator != current_user.operator_type:
        flash('Akses ditolak!', 'error')
        return redirect(url_for('user_operator_dashboard'))
    
    try:
        # Cari semua record dengan public_id yang sama
        same_file_uploads = UploadGambar.query.filter_by(
            stasiun_id=upload.stasiun_id,
            public_id=upload.public_id,
            group_id=upload.group_id
        ).all()
        
        # Cek apakah ada record lain dengan public_id yang sama (untuk stasiun lain)
        other_uploads = UploadGambar.query\
            .filter(UploadGambar.public_id == upload.public_id,
                   UploadGambar.id != upload.id).count()
        
        # Hapus dari Cloudinary (hanya jika ini adalah satu-satunya record)
        if other_uploads == 0:
            success, message = delete_from_cloudinary(upload.public_id)
            if success:
                print(f"✅ Deleted from Cloudinary: {upload.public_id}")
            else:
                print(f"⚠️ Cloudinary delete issue: {message}")
        
        # Hapus record dari database
        for u in same_file_uploads:
            db.session.delete(u)
        
        db.session.commit()
        
        flash(f'✅ Berhasil menghapus {len(same_file_uploads)} gambar dari Cloudinary!', 'success')
        
    except Exception as e:
        db.session.rollback()
        flash(f'❌ Error: {str(e)}', 'error')
    
    return redirect(url_for('user_operator_detail_stasiun', station_id=upload.stasiun_id))

@app.route('/user-operator/cleanup-files/<int:station_id>')
@login_required
@user_operator_required
def user_operator_cleanup_files(station_id):
    """Bersihkan file fisik yang tidak terpakai - TIDAK PERLU UNTUK CLOUDINARY"""
    flash('Fitur cleanup tidak diperlukan karena menggunakan Cloudinary', 'info')
    return redirect(url_for('user_operator_detail_stasiun', station_id=station_id))

@app.route('/api/delete-group-uploads/<int:station_id>', methods=['POST'])
@login_required
@user_operator_required
def api_delete_group_uploads(station_id):
    """Delete semua upload dalam satu grup dari Cloudinary"""
    station = Stasiun.query.get_or_404(station_id)
    
    if station.operator != current_user.operator_type:
        return jsonify({'success': False, 'message': 'Akses ditolak!'}), 403
    
    try:
        data = request.get_json()
        group_id = data.get('group_id')
        
        if not group_id:
            return jsonify({'success': False, 'message': 'Group ID harus diisi!'})
        
        # Ambil semua upload dalam grup
        if group_id == 'ungrouped':
            uploads = UploadGambar.query.filter(
                UploadGambar.stasiun_id == station_id,
                UploadGambar.group_id.is_(None)
            ).all()
        else:
            uploads = UploadGambar.query.filter_by(
                stasiun_id=station_id,
                group_id=group_id
            ).all()
        
        # Koleksi public_id unik
        unique_public_ids = set()
        deleted_count = 0
        
        # Hapus dari Cloudinary (sekali per public_id unik)
        for upload in uploads:
            if upload.public_id not in unique_public_ids:
                # Cek apakah public_id ini dipakai di tempat lain
                other_refs = UploadGambar.query\
                    .filter(UploadGambar.public_id == upload.public_id,
                           UploadGambar.id != upload.id).count()
                
                if other_refs == 0:
                    delete_from_cloudinary(upload.public_id)
                
                unique_public_ids.add(upload.public_id)
            
            db.session.delete(upload)
            deleted_count += 1
        
        # Hapus status updates untuk lawan dalam grup
        if group_id == 'ungrouped':
            for upload in uploads:
                if upload.stasiun_lawan_id:
                    StatusUpdate.query.filter_by(stasiun_lawan_id=upload.stasiun_lawan_id).delete()
        else:
            lawans_in_group = StasiunLawan.query.filter_by(
                stasiun_id=station_id,
                group_id=group_id
            ).all()
            
            for lawan in lawans_in_group:
                StatusUpdate.query.filter_by(stasiun_lawan_id=lawan.id).delete()
        
        db.session.commit()
        
        return jsonify({
            'success': True, 
            'message': f'{deleted_count} gambar dalam grup berhasil dihapus dari Cloudinary!'
        })
        
    except Exception as e:
        db.session.rollback()
        return jsonify({'success': False, 'message': str(e)})

@app.route('/user-operator/fix-group-ids/<int:station_id>')
@login_required
@user_operator_required
def user_operator_fix_group_ids(station_id):
    """Route untuk memperbaiki group_id di upload_gambar"""
    station = Stasiun.query.get_or_404(station_id)
    
    if station.operator != current_user.operator_type:
        flash('Akses ditolak!', 'error')
        return redirect(url_for('user_operator_dashboard'))
    
    try:
        uploads = UploadGambar.query.filter_by(stasiun_id=station_id).all()
        
        for upload in uploads:
            if upload.stasiun_lawan_id:
                lawan = StasiunLawan.query.get(upload.stasiun_lawan_id)
                if lawan:
                    upload.group_id = lawan.group_id
        
        db.session.commit()
        flash('Group IDs berhasil diperbaiki!', 'success')
    except Exception as e:
        db.session.rollback()
        flash(f'Error: {str(e)}', 'error')
    
    return redirect(url_for('user_operator_detail_stasiun', station_id=station_id))



@app.route('/user-operator/manage-groups/<int:station_id>')
@login_required
@user_operator_required
def user_operator_manage_groups(station_id):
    return redirect(url_for('user_operator_grup_stasiun', station_id=station_id))

@app.route('/user-operator/upload-gambar/<int:station_id>', methods=['GET', 'POST'])
@login_required
@user_operator_required
def user_operator_upload_gambar(station_id):
    return redirect(url_for('user_operator_upload', station_id=station_id))

@app.route('/user-operator/manage-uploads/<int:station_id>')
@login_required
@user_operator_required
def user_operator_manage_uploads(station_id):
    return redirect(url_for('user_operator_detail_stasiun', station_id=station_id))

@app.route('/user-operator/card-view')
@login_required
@user_operator_required
def user_operator_card_view():
    return redirect(url_for('user_operator_stasiun'))

# ============================================================================
# USER OPERATOR - HASIL LAPORAN
# ============================================================================
@app.route('/user-operator/hasil-laporan')
@login_required
@user_operator_required
def user_operator_hasil_laporan():
    operator = current_user.operator_type
    page = request.args.get('page', 1, type=int)
    per_page = request.args.get('per_page', 10, type=int)
    
    # ===== TAMBAHKAN SEMUA PARAMETER FILTER =====
    filter_kota = request.args.get('filter_kota', 'all')
    search_nomor_ba = request.args.get('search_nomor_ba', '').strip()
    search_nomor_spt = request.args.get('search_nomor_spt', '').strip()
    tanggal_mulai = request.args.get('tanggal_mulai', '').strip()
    tanggal_sampai = request.args.get('tanggal_sampai', '').strip()
    
    # Query hanya untuk operator yang sama
    query = HasilLaporan.query.filter(HasilLaporan.operator == operator)
    
    # Filter berdasarkan kota/kab
    if filter_kota != 'all':
        query = query.filter(HasilLaporan.kota_kab == filter_kota)
    
    # Filter pencarian Nomor BA (partial match)
    if search_nomor_ba:
        query = query.filter(HasilLaporan.nomor_ba.ilike(f'%{search_nomor_ba}%'))
    
    # Filter pencarian Nomor SPT (partial match)
    if search_nomor_spt:
        query = query.filter(HasilLaporan.nomor_spt.ilike(f'%{search_nomor_spt}%'))
    
    # Filter rentang tanggal mulai
    if tanggal_mulai:
        try:
            tanggal_mulai_date = datetime.strptime(tanggal_mulai, '%Y-%m-%d').date()
            query = query.filter(HasilLaporan.tanggal_ba >= tanggal_mulai_date)
        except ValueError:
            pass
    
    # Filter rentang tanggal sampai
    if tanggal_sampai:
        try:
            tanggal_sampai_date = datetime.strptime(tanggal_sampai, '%Y-%m-%d').date()
            query = query.filter(HasilLaporan.tanggal_ba <= tanggal_sampai_date)
        except ValueError:
            pass
    
    # Urutkan berdasarkan tanggal dibuat terbaru
    query = query.order_by(HasilLaporan.created_at.desc())
    
    # Pagination
    pagination = query.paginate(page=page, per_page=per_page, error_out=False)
    laporan_list = pagination.items
    
    # Ambil daftar kota yang pernah ada dalam laporan untuk operator ini
    kota_with_laporan = db.session.query(HasilLaporan.kota_kab)\
        .filter(HasilLaporan.operator == operator)\
        .distinct().all()
    kota_with_laporan = [k[0] for k in kota_with_laporan if k[0]]
    
    return render_template('user_operator/hasil_laporan.html',
                         laporan_list=laporan_list,
                         pagination=pagination,
                         operator=operator,
                         operator_name=operator.upper(),
                         kota_list=KOTA_LIST,
                         kota_with_laporan=kota_with_laporan,
                         filter_kota=filter_kota,
                         search_nomor_ba=search_nomor_ba,
                         search_nomor_spt=search_nomor_spt,
                         tanggal_mulai=tanggal_mulai,
                         tanggal_sampai=tanggal_sampai)

# ============================================================================
# API ENDPOINTS - UPDATE UNTUK CLOUDINARY
# ============================================================================

@app.route('/api/upload/<int:upload_id>')
@login_required
def api_get_upload_detail(upload_id):
    upload = UploadGambar.query.get_or_404(upload_id)
    
    station = Stasiun.query.get(upload.stasiun_id)
    if current_user.role == 'user_operator' and station.operator != current_user.operator_type:
        return jsonify({'error': 'Akses ditolak'}), 403
    
    catatan = ''
    if upload.stasiun_lawan_id:
        status_update = StatusUpdate.query\
            .filter_by(stasiun_lawan_id=upload.stasiun_lawan_id)\
            .order_by(StatusUpdate.updated_at.desc())\
            .first()
        if status_update:
            catatan = status_update.catatan or ''
    
    # Dapatkan thumbnail URL
    thumbnail_url = None
    if upload.public_id:
        thumbnail_url, _ = cloudinary_url(
            upload.public_id, 
            width=200, 
            height=200, 
            crop="fill", 
            quality="auto"
        )
    
    return jsonify({
        'success': True,
        'upload': {
            'id': upload.id,
            'public_id': upload.public_id,
            'url': upload.cloudinary_url,
            'thumbnail_url': thumbnail_url,
            'original_filename': upload.original_filename,
            'status': upload.status,
            'width': upload.width,
            'height': upload.height,
            'format': upload.format,
            'size_mb': round(upload.bytes_size / (1024 * 1024), 2) if upload.bytes_size else None,
            'uploaded_at': upload.uploaded_at.strftime('%Y-%m-%d %H:%M:%S')
        },
        'catatan': catatan
    })

@app.route('/api/update_upload_status/<int:upload_id>', methods=['PUT'])
@login_required
def api_update_upload_status(upload_id):
    upload = UploadGambar.query.get_or_404(upload_id)
    
    station = Stasiun.query.get(upload.stasiun_id)
    if current_user.role == 'user_operator' and station.operator != current_user.operator_type:
        return jsonify({'error': 'Akses ditolak'}), 403
    
    try:
        data = request.get_json()
        status = data.get('status')
        catatan = data.get('catatan', '')
        
        if not status:
            return jsonify({'success': False, 'message': 'Status harus diisi!'})
        
        upload.status = status
        
        if upload.stasiun_lawan_id:
            status_update = StatusUpdate.query\
                .filter_by(stasiun_lawan_id=upload.stasiun_lawan_id)\
                .order_by(StatusUpdate.updated_at.desc())\
                .first()
            
            if not status_update:
                status_update = StatusUpdate(
                    stasiun_lawan_id=upload.stasiun_lawan_id,
                    status=status,
                    updated_by=current_user.id,
                    catatan=catatan
                )
                db.session.add(status_update)
            else:
                status_update.status = status
                status_update.updated_by = current_user.id
                status_update.catatan = catatan
                status_update.updated_at = datetime.utcnow()
        
        db.session.commit()
        
        return jsonify({
            'success': True,
            'message': 'Status berhasil diperbarui!'
        })
        
    except Exception as e:
        db.session.rollback()
        return jsonify({'success': False, 'message': str(e)})

@app.route('/api/delete_upload/<int:upload_id>', methods=['DELETE'])
@login_required
def api_delete_upload(upload_id):
    upload = UploadGambar.query.get_or_404(upload_id)
    
    station = Stasiun.query.get(upload.stasiun_id)
    if current_user.role == 'user_operator' and station.operator != current_user.operator_type:
        return jsonify({'error': 'Akses ditolak'}), 403
    
    try:
        # Hapus dari Cloudinary
        success, message = delete_from_cloudinary(upload.public_id)
        
        # Hapus dari database
        db.session.delete(upload)
        db.session.commit()
        
        if success:
            return jsonify({
                'success': True,
                'message': 'Gambar berhasil dihapus dari Cloudinary!'
            })
        else:
            return jsonify({
                'success': True,
                'message': f'Gambar dihapus dari database, tapi mungkin masih di Cloudinary: {message}'
            })
            
    except Exception as e:
        db.session.rollback()
        return jsonify({'success': False, 'message': str(e)})

@app.route('/api/stasiun/<int:station_id>/uploads-detail')
@login_required
def api_get_uploads_detail(station_id):
    station = Stasiun.query.get_or_404(station_id)
    
    if current_user.role == 'user_operator' and station.operator != current_user.operator_type:
        return jsonify({'error': 'Akses ditolak'}), 403
    
    uploads = UploadGambar.query.filter_by(stasiun_id=station_id).all()
    
    result = []
    for upload in uploads:
        opponent_name = None
        if upload.stasiun_lawan_id:
            lawan = StasiunLawan.query.get(upload.stasiun_lawan_id)
            opponent_name = lawan.nama_stasiun_lawan if lawan else None
        
        # Dapatkan thumbnail URL
        thumbnail_url = None
        if upload.public_id:
            thumbnail_url, _ = cloudinary_url(
                upload.public_id, 
                width=200, 
                height=200, 
                crop="fill", 
                quality="auto"
            )
        
        result.append({
            'id': upload.id,
            'public_id': upload.public_id,
            'url': upload.cloudinary_url,
            'thumbnail_url': thumbnail_url,
            'original_filename': upload.original_filename,
            'status': upload.status,
            'opponent': opponent_name,
            'date': upload.uploaded_at.strftime('%d/%m/%Y'),
            'format': upload.format,
            'width': upload.width,
            'height': upload.height,
            'size_mb': round(upload.bytes_size / (1024 * 1024), 2) if upload.bytes_size else None
        })
    
    return jsonify(result)

@app.route('/api/stasiun-lawan/<int:station_id>/groups')
@login_required
def api_get_stasiun_lawan_groups(station_id):
    station = Stasiun.query.get_or_404(station_id)
    
    if current_user.role == 'user_operator' and station.operator != current_user.operator_type:
        return jsonify({'error': 'Akses ditolak'}), 403
    
    stasiun_lawan = StasiunLawan.query.filter_by(stasiun_id=station_id).order_by(StasiunLawan.urutan).all()
    
    groups = {}
    for lawan in stasiun_lawan:
        group_key = lawan.group_id if lawan.group_id is not None else 'no_group'
        if group_key not in groups:
            groups[group_key] = {
                'group_id': lawan.group_id,
                'opponents': []
            }
        groups[group_key]['opponents'].append({
            'id': lawan.id,
            'nama': lawan.nama_stasiun_lawan,
            'freq_tx': lawan.freq_tx,
            'freq_rx': lawan.freq_rx,
            'group_id': lawan.group_id,
            'urutan': lawan.urutan
        })
    
    result = []
    for group_key, group_data in groups.items():
        result.append(group_data)
    
    return jsonify(result)

# ============================================================================
# ROUTE UNTUK MELIHAT GAMBAR (TIDAK DIGUNAKAN LAGI)
# ============================================================================

@app.route('/uploads/<path:filename>')
def serve_uploaded_file(filename):
    """Tidak digunakan lagi karena pakai Cloudinary"""
    return "Gambar menggunakan Cloudinary", 404

# ============================================================================
# INITIALIZATION
# ============================================================================

def init_database():
    """Inisialisasi database dengan data default"""
    with app.app_context():
        try:
            db.create_all()
            
            if User.query.count() == 0:
                users_data = [
                    {
                        'username': 'admin_master',
                        'password': 'adminmasterbalmon2026',
                        'role': 'admin_master',
                        'operator_type': None
                    },
                    {
                        'username': 'pic_telkom',
                        'password': 'pictelkombalmon2026',
                        'role': 'admin_operator',
                        'operator_type': 'telkom'
                    },
                    {
                        'username': 'pic_telkomsel',
                        'password': 'pictelkomselbalmon2026',
                        'role': 'admin_operator',
                        'operator_type': 'telkomsel'
                    },
                    {
                        'username': 'pic_indosat',
                        'password': 'picindosatbalmon2026',
                        'role': 'admin_operator',
                        'operator_type': 'indosat'
                    },
                    {
                        'username': 'pic_xl',
                        'password': 'picxlbalmon2026',
                        'role': 'admin_operator',
                        'operator_type': 'xl'
                    },
                    {
                        'username': 'telkom',
                        'password': 'operatortelkombalmon01',
                        'role': 'user_operator',
                        'operator_type': 'telkom'
                    },
                    {
                        'username': 'telkomsel',
                        'password': 'operatortelkomselbalmon02',
                        'role': 'user_operator',
                        'operator_type': 'telkomsel'
                    },
                    {
                        'username': 'ioh',
                        'password': 'operatorindosatbalmon010',
                        'role': 'user_operator',
                        'operator_type': 'ioh'
                    },
                    {
                        'username': 'xl_smart',
                        'password': 'operatorxlbalmon20',
                        'role': 'user_operator',
                        'operator_type': 'xl smart'
                    }
                ]
                
                for user_data in users_data:
                    hashed_password = generate_password_hash(user_data['password'], method='pbkdf2:sha256')
                    user = User(
                        username=user_data['username'],
                        password=hashed_password,
                        role=user_data['role'],
                        operator_type=user_data['operator_type']
                    )
                    db.session.add(user)
                
                db.session.commit()
                print("=" * 60)
                print("DATABASE BERHASIL DIINISIALISASI")
                print("=" * 60)
                print("User login telah dibuat sesuai dengan data di database.")
                print("Silakan gunakan kredensial yang telah diberikan.")
                print("=" * 60)
            else:
                print("Database sudah ada, melewati inisialisasi...")
            
            create_upload_folders()
            print("Folder upload berhasil dibuat/dicek.")
            print("=" * 60)
                
        except Exception as e:
            print(f"Error initializing database: {e}")
            db.session.rollback()        

# ============================================================================
# MAIN EXECUTION
# ============================================================================


# ============================================================================
# DETAIL STASIUN V2 ROUTES & API PERIKSA PARAMETER
# ============================================================================

@app.route('/admin-master/daftar-stasiun-v2/detail/<int:station_id>')
@login_required
@admin_master_required
def admin_master_detail_stasiun_v2(station_id):
    return admin_master_detail_stasiun(station_id)

@app.route('/admin-operator/daftar-stasiun-v2/detail/<int:station_id>')
@login_required
@admin_operator_required
def admin_operator_detail_stasiun_v2(station_id):
    return admin_operator_detail_stasiun(station_id)

@app.route('/user-operator/daftar-stasiun-v2/detail/<int:station_id>')
@login_required
@user_operator_required
def user_operator_detail_stasiun_v2(station_id):
    return user_operator_detail_stasiun(station_id)

@app.route('/api/stasiun-lawan/<int:lawan_id>/update-parameter', methods=['POST'])
@login_required
def api_update_stasiun_lawan_parameter(lawan_id):
    try:
        data = request.get_json() or {}
        lawan = StasiunLawan.query.get_or_404(lawan_id)
        station = Stasiun.query.get(lawan.stasiun_id)
        
        # Validasi akses untuk operator
        if current_user.role in ['admin_operator', 'user_operator']:
            if station.operator != current_user.operator_type:
                return jsonify({'success': False, 'message': 'Akses ditolak: bukan stasiun milik operator Anda'}), 403

        # Update parameter stasiun lawan
        if 'nomor_isr' in data:
            lawan.nomor_isr = data['nomor_isr'].strip() if data['nomor_isr'] else None
        if 'link_id' in data:
            lawan.link_id = data['link_id'].strip() if data['link_id'] else None
        if 'id_sims' in data:
            lawan.id_sims = data['id_sims'].strip() if data['id_sims'] else None
        if 'client_operator' in data:
            lawan.client_operator = data['client_operator'].strip() if data['client_operator'] else None
        if 'hasil_pemeriksaan' in data and data['hasil_pemeriksaan']:
            lawan.hasil_pemeriksaan = data['hasil_pemeriksaan'].strip()
        if 'metode_pemeriksaan' in data:
            lawan.metode_pemeriksaan = data['metode_pemeriksaan'].strip() if data['metode_pemeriksaan'] else None
        if 'sertifikat_perangkat' in data:
            lawan.sertifikat_perangkat = data['sertifikat_perangkat'].strip() if data['sertifikat_perangkat'] else None
        if 'mulai_beroperasi' in data:
            lawan.mulai_beroperasi = data['mulai_beroperasi'].strip() if data['mulai_beroperasi'] else None
        if 'tanggal_pemeriksaan' in data and data['tanggal_pemeriksaan']:
            try:
                lawan.tanggal_pemeriksaan = datetime.strptime(data['tanggal_pemeriksaan'], '%Y-%m-%d').date()
            except ValueError:
                pass

        # Hasil pemeriksaan parameter teknis
        if 'hasil_freq_tx' in data:
            lawan.hasil_freq_tx = data['hasil_freq_tx'].strip() if data['hasil_freq_tx'] else None
        if 'hasil_freq_rx' in data:
            lawan.hasil_freq_rx = data['hasil_freq_rx'].strip() if data['hasil_freq_rx'] else None
        if 'hasil_bw_mhz' in data:
            lawan.hasil_bw_mhz = data['hasil_bw_mhz'].strip() if data['hasil_bw_mhz'] else None
        if 'hasil_tipe_perangkat' in data:
            lawan.hasil_tipe_perangkat = data['hasil_tipe_perangkat'].strip() if data['hasil_tipe_perangkat'] else None
        if 'hasil_longitude' in data:
            lawan.hasil_longitude = data['hasil_longitude'].strip() if data['hasil_longitude'] else None
        if 'hasil_latitude' in data:
            lawan.hasil_latitude = data['hasil_latitude'].strip() if data['hasil_latitude'] else None

        if 'catatan_pemeriksaan' in data:
            lawan.catatan_pemeriksaan = data['catatan_pemeriksaan'].strip() if data['catatan_pemeriksaan'] else None
        
        lawan.updated_at = datetime.utcnow()

        # Update status operasional jika diubah
        status_val = data.get('status_lawan')
        if status_val and status_val in ['aktif', 'tidak_aktif', 'tidak_sesuai', 'belum_ada']:
            catatan = data.get('catatan_pemeriksaan') or f"Status diperbarui oleh {current_user.username}"
            new_update = StatusUpdate(
                stasiun_lawan_id=lawan.id,
                status=status_val,
                updated_by=current_user.id,
                updated_at=datetime.utcnow(),
                catatan=catatan
            )
            db.session.add(new_update)

        # Update status stasiun induk jika diperiksa
        if station and data.get('hasil_pemeriksaan') in ['Sesuai ISR', 'Tidak Sesuai Parameter Teknis']:
            station.status_pemeriksaan = 'sudah_diperiksa'
            station.checked_at = datetime.utcnow()
            station.checked_by = current_user.id

        db.session.commit()

        return jsonify({
            'success': True,
            'message': 'Parameter pemeriksaan stasiun lawan berhasil diperbarui!',
            'data': {
                'id': lawan.id,
                'nomor_isr': lawan.nomor_isr,
                'hasil_pemeriksaan': lawan.hasil_pemeriksaan,
                'freq_tx': lawan.freq_tx,
                'freq_rx': lawan.freq_rx,
                'bw_mhz': lawan.bw_mhz,
                'tipe_perangkat': lawan.tipe_perangkat,
                'status': status_val or 'aktif'
            }
        })
    except Exception as e:
        db.session.rollback()
        return jsonify({'success': False, 'message': str(e)}), 500


_STATION_OCR_CACHE = {}

@app.route('/api/stasiun/<int:station_id>/ocr-scan-uploads', methods=['POST'])
@login_required
def api_ocr_scan_station_uploads(station_id):
    """Jalankan RapidOCR langsung pada gambar-gambar capture terupload stasiun ini dengan optimasi Cloudinary"""
    try:
        data = request.get_json(silent=True) or {}
        force_rescan = request.args.get('force') == '1' or data.get('force', False)
        
        station = Stasiun.query.get_or_404(station_id)
        # Ambil gambar capture terupload stasiun ini (urutkan ID asc, ambil max 4 untuk kecepatan & hemat RAM)
        uploads = UploadGambar.query.filter_by(stasiun_id=station_id).order_by(UploadGambar.id.asc()).limit(4).all()
        if not uploads:
            return jsonify({'status': 'error', 'message': 'Belum ada capturean yang diunggah untuk stasiun ini pada menu galeri stasiun.'}), 404

        upload_ids = [u.id for u in uploads]

        # Cek cache server
        if not force_rescan and station_id in _STATION_OCR_CACHE:
            cached = _STATION_OCR_CACHE[station_id]
            if cached.get('upload_ids') == upload_ids and cached.get('links'):
                return jsonify({
                    'status': 'success',
                    'links': cached['links'],
                    'confidence': cached.get('confidence', 95.0),
                    'text': cached.get('text', ''),
                    'cached': True
                })

        import tempfile, shutil, requests, gc
        from mwlink_ocr_engine import process_microwave_ocr

        temp_dir = tempfile.mkdtemp()
        temp_paths = []
        try:
            for idx, up in enumerate(uploads):
                if up.cloudinary_url:
                    img_url = up.cloudinary_url
                    # Optimasi Cloudinary: ubah resolusi ke max width 1200 & compress ke jpeg agar unduhan super cepat (< 100KB)
                    if '/image/upload/' in img_url and '/w_' not in img_url:
                        img_url = img_url.replace('/image/upload/', '/image/upload/w_1200,c_limit,q_auto,f_jpg/')
                    
                    fpath = os.path.join(temp_dir, f"upload_{idx+1}.jpg")
                    try:
                        r_img = requests.get(img_url, timeout=12, headers={'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)'})
                        if r_img.status_code == 200:
                            with open(fpath, 'wb') as f_out:
                                f_out.write(r_img.content)
                            temp_paths.append(fpath)
                    except Exception as err_dl:
                        print(f"Peringatan: Gagal unduh {img_url}: {err_dl}")

            if not temp_paths:
                return jsonify({'status': 'error', 'message': 'Tidak dapat mengunduh gambar capture terupload dari Cloudinary.'}), 400

            result_data = process_microwave_ocr(temp_paths, station_name_hint=station.stasiun_name)
            links = result_data.get('links', [])
            confidence = result_data.get('confidence', 95.0)

            for l in links:
                if 'confidence' not in l:
                    l['confidence'] = confidence

            _STATION_OCR_CACHE[station_id] = {
                'upload_ids': upload_ids,
                'links': links,
                'confidence': confidence,
                'text': result_data.get('text', '')
            }

            gc.collect()

            return jsonify({
                'status': 'success',
                'links': links,
                'confidence': confidence,
                'text': result_data.get('text', ''),
                'cached': False
            })
        finally:
            shutil.rmtree(temp_dir, ignore_errors=True)
    except Exception as e:
        import traceback
        print(f"Error in api_ocr_scan_station_uploads: {traceback.format_exc()}")
        return jsonify({'status': 'error', 'message': f"Gagal membaca capturean stasiun: {str(e)}"}), 500

if __name__ == '__main__':
    print("Starting Station Monitoring System...")
    print("Initializing database...")
    init_database()
    print("Application is running on http://localhost:5000")
    
    # Otomatis buka browser ke tampilan web
    if os.environ.get('WERKZEUG_RUN_MAIN') == 'true' or not app.debug:
        import threading
        import webbrowser
        threading.Timer(1.5, lambda: webbrowser.open('http://localhost:5000')).start()
        
    app.run(debug=True, host='0.0.0.0', port=5000)