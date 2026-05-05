# safe_migrate_operators.py
from app import app, db, User, Stasiun
from datetime import datetime

def safe_migrate():
    with app.app_context():
        print("=" * 70)
        print("MIGRASI OPERATOR YANG AMAN (DATA TETAP ADA)")
        print("=" * 70)
        
        try:
            # 1. Cek data sebelum migrasi
            print("\n📊 DATA SEBELUM MIGRASI:")
            users_indosat = User.query.filter_by(operator_type='indosat').all()
            users_xl = User.query.filter_by(operator_type='xl').all()
            stations_indosat = Stasiun.query.filter_by(operator='indosat').all()
            stations_xl = Stasiun.query.filter_by(operator='xl').all()
            
            print(f"   Users (indosat): {len(users_indosat)} orang")
            for u in users_indosat:
                print(f"     - {u.username}")
            
            print(f"   Users (xl): {len(users_xl)} orang")
            for u in users_xl:
                print(f"     - {u.username}")
            
            print(f"   Stasiun (indosat): {len(stations_indosat)} stasiun")
            print(f"   Stasiun (xl): {len(stations_xl)} stasiun")
            
            # 2. Konfirmasi
            print("\n⚠️  PERINGATAN: Data di atas akan diupdate!")
            print("   Operator 'indosat' → 'ioh'")
            print("   Operator 'xl' → 'xl smart'")
            print("   SEMUA DATA LAIN TETAP ADA (nama stasiun, gambar, status, dll)")
            
            confirm = input("\nLanjutkan? (y/n): ")
            if confirm.lower() != 'y':
                print("❌ Migrasi dibatalkan.")
                return
            
            # 3. Lakukan migrasi
            print("\n🔄 MENGUPDATE DATA...")
            
            # Update users
            for user in users_indosat:
                old = user.operator_type
                user.operator_type = 'ioh'
                print(f"   ✓ User {user.username}: {old} → ioh")
            
            for user in users_xl:
                old = user.operator_type
                user.operator_type = 'xl smart'
                print(f"   ✓ User {user.username}: {old} → xl smart")
            
            # Update stasiun
            for station in stations_indosat:
                old = station.operator
                station.operator = 'ioh'
                print(f"   ✓ Stasiun {station.stasiun_name}: {old} → ioh")
            
            for station in stations_xl:
                old = station.operator
                station.operator = 'xl smart'
                print(f"   ✓ Stasiun {station.stasiun_name}: {old} → xl smart")
            
            # Commit ke database
            db.session.commit()
            
            # 4. Verifikasi hasil
            print("\n📊 DATA SETELAH MIGRASI:")
            users_ioh = User.query.filter_by(operator_type='ioh').count()
            users_xl_smart = User.query.filter_by(operator_type='xl smart').count()
            stations_ioh = Stasiun.query.filter_by(operator='ioh').count()
            stations_xl_smart = Stasiun.query.filter_by(operator='xl smart').count()
            
            print(f"   Users (ioh): {users_ioh} orang")
            print(f"   Users (xl smart): {users_xl_smart} orang")
            print(f"   Stasiun (ioh): {stations_ioh} stasiun")
            print(f"   Stasiun (xl smart): {stations_xl_smart} stasiun")
            
            print("\n✅ MIGRASI SELESAI!")
            print("   Semua data tetap utuh, hanya nama operator yang berubah.")
            
        except Exception as e:
            db.session.rollback()
            print(f"\n❌ ERROR: {str(e)}")
            print("   Migrasi dibatalkan, data tidak berubah.")
            import traceback
            traceback.print_exc()

if __name__ == '__main__':
    safe_migrate()