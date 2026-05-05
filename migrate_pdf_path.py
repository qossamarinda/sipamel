import psycopg2
import os
from dotenv import load_dotenv

load_dotenv()

# Ambil DATABASE_URL
DATABASE_URL = os.environ.get('DATABASE_URL')

if not DATABASE_URL:
    DATABASE_URL = 'postgresql://flask_user:rahasia123@localhost:5432/station_monitoring'

print(f"📌 Menghubungkan ke database...")

try:
    conn = psycopg2.connect(DATABASE_URL)
    cur = conn.cursor()
    
    # Cek apakah kolom sudah ada
    cur.execute("""
        SELECT column_name 
        FROM information_schema.columns 
        WHERE table_name='hasil_laporan' AND column_name='file_pdf_path'
    """)
    
    if not cur.fetchone():
        print("📝 Menambahkan kolom file_pdf_path...")
        cur.execute("ALTER TABLE hasil_laporan ADD COLUMN file_pdf_path VARCHAR(500)")
        conn.commit()
        print("✅ Kolom file_pdf_path berhasil ditambahkan!")
    else:
        print("✅ Kolom file_pdf_path sudah ada, tidak perlu migrasi!")
    
    cur.close()
    conn.close()
    
except Exception as e:
    print(f"❌ Error: {e}")