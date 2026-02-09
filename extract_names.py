import pandas as pd
import json

def extract_factory2_data(file_path):
    """
    Ekstrak data dari file Factory 2.xlsx
    """
    print("📂 Membaca Factory 2.xlsx...")
    
    try:
        df = pd.read_excel(file_path, sheet_name="Factory 2", header=None)
        print(f"✅ Berhasil membaca. Shape: {df.shape}")
        
        shift_a_names = []
        shift_b_names = []
        
        # Ambil Shift A dari kolom A
        for i in range(1, len(df)):
            name = df.iat[i, 0]
            if pd.notna(name) and isinstance(name, str) and name.strip():
                shift_a_names.append(name.strip())
        
        # Ambil Shift 2 dari kolom C (Shift 2 = Shift B)
        for i in range(1, len(df)):
            name = df.iat[i, 2]
            if pd.notna(name) and isinstance(name, str) and name.strip():
                # Filter #VALUE! dan empty
                if name != "#VALUE!" and name.strip():
                    shift_b_names.append(name.strip())
        
        print(f"📊 Shift A: {len(shift_a_names)} nama")
        print(f"📊 Shift B: {len(shift_b_names)} nama")
        
        # Tampilkan sample
        print("\n🎯 Sample Shift A (5 pertama):")
        for name in shift_a_names[:5]:
            print(f"   • {name}")
        
        print("\n🎯 Sample Shift B (5 pertama):")
        for name in shift_b_names[:5]:
            print(f"   • {name}")
        
        return shift_a_names, shift_b_names
        
    except Exception as e:
        print(f"❌ Error: {e}")
        return [], []

def create_complete_json():
    """
    Buat file JSON lengkap dengan semua factory
    """
    print("\n" + "="*50)
    print("BUAT FILE JSON LENGKAP")
    print("="*50)
    
    # Data Factory 2 dari file baru
    f2_shift_a, f2_shift_b = extract_factory2_data("Factory 2.xlsx")
    
    # Data untuk factory lain (default)
    default_names = [f"Operator {i+1}" for i in range(24)]
    default_staff = [f"Staff {i+1}" for i in range(24)]
    
    # Struktur lengkap
    data = {
        "factory1": {
            "shift_a": default_names,
            "shift_b": default_staff
        },
        "factory2": {
            "shift_a": f2_shift_a[:48],  # Maks 48
            "shift_b": f2_shift_b[:48]
        },
        "factory3": {
            "shift_a": [f"F3-Operator-{i+1}" for i in range(24)],
            "shift_b": [f"F3-Staff-{i+1}" for i in range(24)]
        },
        "factory4": {
            "shift_a": [f"F4-Karyawan-{i+1}" for i in range(24)],
            "shift_b": [f"F4-Pegawai-{i+1}" for i in range(24)]
        }
    }
    
    # Simpan ke file
    output_file = "member_data.json"
    with open(output_file, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    
    print(f"\n✅ File JSON berhasil dibuat: {output_file}")
    
    # Ringkasan
    print("\n📋 RINGKASAN DATA:")
    for factory in ["factory1", "factory2", "factory3", "factory4"]:
        print(f"\n{factory.upper()}:")
        print(f"  Shift A: {len(data[factory]['shift_a'])} nama")
        print(f"  Shift B: {len(data[factory]['shift_b'])} nama")
        print(f"  Sample A: {data[factory]['shift_a'][:3]}")
        print(f"  Sample B: {data[factory]['shift_b'][:3]}")

if __name__ == "__main__":
    create_complete_json()