import pandas as pd
import numpy as np
from pathlib import Path
import glob
import re

# Dosyayı bul
desktop = Path(r"C:\Users\Hira\OneDrive\Masaüstü")
files = glob.glob(str(desktop / 'FR.KP*'))
if not files:
    raise SystemExit('FR.KP dosyası bulunamadı')
excel_path = files[0]
print(f"Kullanılan dosya: {excel_path}\n")

# ===== ODA KOKULARI VERİSİ =====
print("="*80)
print("1. ODA KOKULARI VERİLERİ İŞLENİYOR...")
print("="*80)

df_oda_raw = pd.read_excel(excel_path, sheet_name='ODA KOKULARI', header=None)
# Satır 1'de asıl başlıklar var
headers_row1 = df_oda_raw.iloc[1].values.tolist()
# Satır 0'dan eksik başlıkları tamamla
headers_row0 = df_oda_raw.iloc[0].values.tolist()
# Birleştir - Satır 0 NaN ise Satır 1'i kullan
headers = []
for i in range(len(headers_row0)):
    if pd.isna(headers_row0[i]):
        headers.append(headers_row1[i] if not pd.isna(headers_row1[i]) else f'Col_{i}')
    else:
        headers.append(headers_row0[i])

# Gerçek veriler satır 2'den başlıyor
df_oda = df_oda_raw.iloc[2:].copy()
df_oda.columns = headers
df_oda.columns = df_oda.columns.str.strip()  # Sütun adlarındaki boşlukları temizle
df_oda = df_oda.dropna(subset=['PARTİ NO'])  # Boş satırları çıkar
df_oda = df_oda.reset_index(drop=True)

print(f"\nTemizlenen ODA KOKULARI verisi:")
print(f"Satır sayısı: {len(df_oda)}")
print(f"Sütunlar: {list(df_oda.columns)}")
print(f"\nİlk 3 satır:")
print(df_oda.head(3))

# ===== SPEKT LİSTESİ (STANDARTLAR) =====
print("\n" + "="*80)
print("2. SPEKT LİSTESİ (STANDARTLAR) İŞLENİYOR...")
print("="*80)

df_spekt_raw = pd.read_excel(excel_path, sheet_name='SPEKT LİSTESİ', header=None)

# Standartları ayıkla
# Satır 0: başlıklar, satır 3: ODA KOKULARI, satır 5: NICHE, satır 7: SPREY, satır 9: OTO
standards = {}

# Sütun indeksleri (başlık satırına göre)
col_map = {
    'urun_tipi': 0,
    'alkol': 3,
    'yogunluk': 5,
    'kirilma': 8,
    'ph': 11,
    'sicaklik': 13
}

# ODA KOKULARI standartları (satır 3)
row = 3
standards['ODA KOKULARI'] = {
    'Alkol': df_spekt_raw.iloc[row, col_map['alkol']],
    'Yoğunluk': df_spekt_raw.iloc[row, col_map['yogunluk']],
    'Kırılma İndisi': df_spekt_raw.iloc[row, col_map['kirilma']],
    'pH': df_spekt_raw.iloc[row, col_map['ph']],
    'Sıcaklık': df_spekt_raw.iloc[row, col_map['sicaklik']]
}

# NICHE ODA KOKULARI (satır 5)
row = 5
standards['NICHE ODA KOKULARI'] = {
    'Alkol': df_spekt_raw.iloc[row, col_map['alkol']],
    'Yoğunluk': df_spekt_raw.iloc[row, col_map['yogunluk']],
    'Kırılma İndisi': df_spekt_raw.iloc[row, col_map['kirilma']],
    'pH': df_spekt_raw.iloc[row, col_map['ph']],
    'Sıcaklık': df_spekt_raw.iloc[row, col_map['sicaklik']]
}

print("\nStandartlar:")
for urun_tipi, specs in standards.items():
    print(f"\n{urun_tipi}:")
    for param, value in specs.items():
        print(f"  {param}: {value}")

# ===== KARŞILAŞTIRMA =====
print("\n" + "="*80)
print("3. KALITE KONTROLÜ - STANDARTLARLA KARŞILAŞTIRMA")
print("="*80)

# Numrik değerlerin aralıklarını parse et
def parse_range(range_str):
    """'80-95' -> (80, 95), '2,0-6,0' -> (2.0, 6.0)"""
    if pd.isna(range_str) or not isinstance(range_str, str):
        return None, None
    
    range_str = str(range_str).replace(',', '.')
    parts = range_str.split('-')
    if len(parts) == 2:
        try:
            return float(parts[0].strip()), float(parts[1].strip())
        except:
            return None, None
    return None, None

# ODA KOKULARI ürünleri için kontrol
print("\nODA KOKULARI ürünlerinin kalite kontrolü:")
print("-" * 80)

spec = standards['ODA KOKULARI']
alkol_min, alkol_max = parse_range(spec['Alkol'])
yogunluk_min, yogunluk_max = parse_range(spec['Yoğunluk'])
kirilma_min, kirilma_max = parse_range(spec['Kırılma İndisi'])
ph_min, ph_max = parse_range(spec['pH'])

print(f"\nBeklenilen Aralıklar:")
print(f"  Alkol: {alkol_min} - {alkol_max}")
print(f"  Yoğunluk: {yogunluk_min} - {yogunluk_max}")
print(f"  Kırılma İndisi: {kirilma_min} - {kirilma_max}")
print(f"  pH: {ph_min} - {ph_max}")

# Kontrol sütunları ekle
results = []

for idx, row in df_oda.iterrows():
    try:
        party = row['PARTİ NO']
        urun = row['ÜRÜN ADI']
        
        # Numrik verileri dönüştür
        alkol_val = float(row['Alkol Derecesi']) if pd.notna(row['Alkol Derecesi']) else None
        yogunluk_val = float(str(row['Yoğunluk, 25ºC']).replace(',', '.')) if pd.notna(row['Yoğunluk, 25ºC']) else None
        kirilma_val = float(str(row['Kırılma İndisi']).replace(',', '.')) if pd.notna(row['Kırılma İndisi']) else None
        ph_val = float(str(row['pH ,25ºC']).replace(',', '.')) if pd.notna(row['pH ,25ºC']) else None
        
        # Uygunluk kontrolleri
        alkol_ok = alkol_min <= alkol_val <= alkol_max if alkol_val is not None and alkol_min is not None else None
        yogunluk_ok = yogunluk_min <= yogunluk_val <= yogunluk_max if yogunluk_val is not None and yogunluk_min is not None else None
        kirilma_ok = kirilma_min <= kirilma_val <= kirilma_max if kirilma_val is not None and kirilma_min is not None else None
        ph_ok = ph_min <= ph_val <= ph_max if ph_val is not None and ph_min is not None else None
        
        # Genel uygunluk
        checks = [alkol_ok, yogunluk_ok, kirilma_ok, ph_ok]
        checks = [c for c in checks if c is not None]
        overall_ok = all(checks) if checks else None
        
        results.append({
            'Parti No': party,
            'Ürün Adı': urun,
            'Alkol': alkol_val,
            'Alkol OK': alkol_ok,
            'Yoğunluk': yogunluk_val,
            'Yoğunluk OK': yogunluk_ok,
            'Kırılma': kirilma_val,
            'Kırılma OK': kirilma_ok,
            'pH': ph_val,
            'pH OK': ph_ok,
            'Genel Durum': 'UYGUN' if overall_ok else 'UYGUN DEĞİL' if overall_ok is False else 'VERİ HATASI'
        })
    except Exception as e:
        print(f"Satır {idx} işleme hatası: {e}")

df_results = pd.DataFrame(results)

print(f"\n\nKONTROL SONUÇLARI:")
print(df_results[['Parti No', 'Ürün Adı', 'Genel Durum']].to_string(index=False))

# İstatistikler
print("\n" + "="*80)
print("İSTATİSTİKLER:")
print("="*80)
print(f"Toplam ürün: {len(df_results)}")
print(f"UYGUN: {(df_results['Genel Durum'] == 'UYGUN').sum()}")
print(f"UYGUN DEĞİL: {(df_results['Genel Durum'] == 'UYGUN DEĞİL').sum()}")
print(f"VERİ HATASI: {(df_results['Genel Durum'] == 'VERİ HATASI').sum()}")

# Detaylı sonuçları dosyaya kaydet
output_file = r"C:\Users\Hira\OneDrive\Masaüstü\kalite_kontrol_sonuclari.xlsx"
df_results.to_excel(output_file, index=False)
print(f"\n✓ Sonuçlar kaydedildi: {output_file}")

