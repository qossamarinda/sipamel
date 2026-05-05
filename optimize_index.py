# optimize_index.py
import sys
import os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from app import app, db
from sqlalchemy import text

def add_indexes():
    with app.app_context():
        print("=" * 60)
        print("MENAMBAH INDEX DATABASE")
        print("=" * 60)
        
        indexes = [
            ("idx_stasiun_operator", "stasiun", "operator"),
            ("idx_stasiun_kota", "stasiun", "kota"),
            ("idx_stasiun_name", "stasiun", "stasiun_name"),
            ("idx_stasiun_status", "stasiun", "status_pemeriksaan"),
            ("idx_stasiun_operator_kota", "stasiun", "operator, kota"),
            ("idx_stasiun_lawan_stasiun_id", "stasiun_lawan", "stasiun_id"),
            ("idx_status_updates_lawan_id", "status_updates", "stasiun_lawan_id"),
        ]
        
        for idx_name, table, columns in indexes:
            try:
                db.session.execute(text(
                    f"CREATE INDEX IF NOT EXISTS {idx_name} ON {table}({columns})"
                ))
                print(f"✅ Index {idx_name} berhasil ditambahkan")
            except Exception as e:
                print(f"⚠️ Index {idx_name} gagal: {e}")
        
        db.session.commit()
        print("-" * 60)
        print("✅ SEMUA INDEX BERHASIL DITAMBAHKAN!")
        print("=" * 60)

if __name__ == '__main__':
    add_indexes()