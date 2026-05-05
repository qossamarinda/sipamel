# migrate.py
from app import app, db
from app import HasilLaporan

def migrate():
    with app.app_context():
        print("=" * 60)
        print("🚀 MIGRASI DATABASE - MENAMBAH TABEL HASIL LAPORAN")
        print("=" * 60)
        
        # Cek tabel yang sudah ada
        from sqlalchemy import inspect
        inspector = inspect(db.engine)
        existing_tables = inspector.get_table_names()
        
        print(f"\n📋 Tabel yang sudah ada: {', '.join(existing_tables)}")
        
        if 'hasil_laporan' in existing_tables:
            print("\n⚠️ Tabel 'hasil_laporan' sudah ada!")
            confirm = input("Apakah ingin menghapus dan membuat ulang? (y/n): ")
            if confirm.lower() == 'y':
                print("🗑️ Menghapus tabel lama...")
                HasilLaporan.__table__.drop(db.engine)
                print("✅ Tabel lama dihapus")
                db.create_all()
                print("✅ Tabel baru dibuat")
            else:
                print("✅ Melewati migrasi")
        else:
            # Buat tabel baru
            print("\n📝 Membuat tabel baru 'hasil_laporan'...")
            db.create_all()
            print("✅ Tabel 'hasil_laporan' berhasil dibuat!")
        
        # Verifikasi
        inspector = inspect(db.engine)
        tables = inspector.get_table_names()
        print(f"\n📋 Daftar tabel setelah migrasi: {', '.join(tables)}")
        
        if 'hasil_laporan' in tables:
            columns = inspector.get_columns('hasil_laporan')
            print("\n📊 Struktur tabel hasil_laporan:")
            for col in columns:
                print(f"   - {col['name']}: {col['type']}")
        
        print("\n" + "=" * 60)
        print("✅ MIGRASI SELESAI!")
        print("=" * 60)

if __name__ == '__main__':
    migrate()