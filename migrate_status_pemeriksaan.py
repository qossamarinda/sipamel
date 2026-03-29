# migrate_status_pemeriksaan.py
import sys
import os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from app import app, db
from sqlalchemy import text

def migrate():
    with app.app_context():
        try:
            # Cek apakah kolom sudah ada
            inspector = db.inspect(db.engine)
            columns = [col['name'] for col in inspector.get_columns('stasiun')]
            
            print("=" * 60)
            print("MIGRASI DATABASE - STATUS PEMERIKSAAN")
            print("=" * 60)
            print(f"Kolom yang ada: {columns}")
            print("-" * 60)
            
            # Tambah kolom status_pemeriksaan
            if 'status_pemeriksaan' not in columns:
                db.session.execute(text(
                    "ALTER TABLE stasiun ADD COLUMN status_pemeriksaan VARCHAR(20) DEFAULT 'belum_diperiksa' NOT NULL"
                ))
                print("✅ Kolom 'status_pemeriksaan' berhasil ditambahkan")
            else:
                print("ℹ️ Kolom 'status_pemeriksaan' sudah ada")
            
            # Tambah kolom checked_at
            if 'checked_at' not in columns:
                db.session.execute(text(
                    "ALTER TABLE stasiun ADD COLUMN checked_at TIMESTAMP DEFAULT NULL"
                ))
                print("✅ Kolom 'checked_at' berhasil ditambahkan")
            else:
                print("ℹ️ Kolom 'checked_at' sudah ada")
            
            # Tambah kolom checked_by
            if 'checked_by' not in columns:
                db.session.execute(text(
                    "ALTER TABLE stasiun ADD COLUMN checked_by INTEGER REFERENCES users(id) DEFAULT NULL"
                ))
                print("✅ Kolom 'checked_by' berhasil ditambahkan")
            else:
                print("ℹ️ Kolom 'checked_by' sudah ada")
            
            db.session.commit()
            
            print("-" * 60)
            print("✅ MIGRASI SELESAI!")
            print("=" * 60)
            
            # Tampilkan statistik
            total_stations = db.session.execute(text("SELECT COUNT(*) FROM stasiun")).scalar()
            sudah_diperiksa = db.session.execute(text(
                "SELECT COUNT(*) FROM stasiun WHERE status_pemeriksaan = 'sudah_diperiksa'"
            )).scalar()
            
            print(f"\n📊 STATISTIK:")
            print(f"   Total stasiun: {total_stations}")
            print(f"   Sudah diperiksa: {sudah_diperiksa}")
            print(f"   Belum diperiksa: {total_stations - sudah_diperiksa}")
            
        except Exception as e:
            db.session.rollback()
            print(f"\n❌ ERROR: {e}")
            import traceback
            traceback.print_exc()

if __name__ == '__main__':
    migrate()