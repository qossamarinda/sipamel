# check_migration.py
from app import app, db, User, Stasiun

with app.app_context():
    print("=" * 70)
    print("VERIFIKASI DATA SETELAH MIGRASI")
    print("=" * 70)
    
    # Cek operator di stasiun
    print("\n📊 OPERATOR DI TABEL STASIUN:")
    operators = db.session.query(Stasiun.operator).distinct().all()
    for op in operators:
        count = Stasiun.query.filter_by(operator=op[0]).count()
        print(f"   {op[0]}: {count} stasiun")
    
    # Cek operator_type di users
    print("\n👥 OPERATOR DI TABEL USERS:")
    operator_types = db.session.query(User.operator_type).distinct().all()
    for op_type in operator_types:
        if op_type[0]:
            count = User.query.filter_by(operator_type=op_type[0]).count()
            print(f"   {op_type[0]}: {count} user")
    
    # Cek apakah masih ada indosat/XL lama
    print("\n🔍 CEK DATA LAMA:")
    old_indosat_stations = Stasiun.query.filter_by(operator='indosat').count()
    old_xl_stations = Stasiun.query.filter_by(operator='xl').count()
    old_indosat_users = User.query.filter_by(operator_type='indosat').count()
    old_xl_users = User.query.filter_by(operator_type='xl').count()
    
    print(f"   Stasiun dengan operator 'indosat': {old_indosat_stations}")
    print(f"   Stasiun dengan operator 'xl': {old_xl_stations}")
    print(f"   Users dengan operator_type 'indosat': {old_indosat_users}")
    print(f"   Users dengan operator_type 'xl': {old_xl_users}")
    
    if old_indosat_stations == 0 and old_xl_stations == 0:
        print("\n✅ Data stasiun sudah sesuai!")
    else:
        print("\n⚠️ Masih ada stasiun dengan operator lama!")
    
    if old_indosat_users == 0 and old_xl_users == 0:
        print("✅ Data users sudah sesuai!")
    else:
        print("⚠️ Masih ada users dengan operator_type lama!")