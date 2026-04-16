import pandas as pd
import numpy as np

# Sonuçları oku ve göster
results_file = r"C:\Users\Hira\OneDrive\Masaüstü\ml_model_results.xlsx"

xls = pd.ExcelFile(results_file)

print("="*80)
print("XGBoost MAKİNE ÖĞRENMESİ - TAHMİN SONUÇLARI")
print("="*80)

# Özet bilgileri oku
summary = pd.read_excel(results_file, sheet_name='Özet')
print("\n📊 MODELPERFORMANSI:")
print(summary.to_string(index=False))

# Sınıf dağılımı
class_dist = pd.read_excel(results_file, sheet_name='Sınıf Dağılımı')
print("\n📈 SINIF DAĞILIMI (Test Seti):")
print(class_dist.to_string(index=False))

# Özellik önemi
feature_imp = pd.read_excel(results_file, sheet_name='Özellik Önemi')
print("\n⭐ ÖZELLİK ÖNEMLİLİĞİ:")
print(feature_imp.to_string(index=False))

# Tüm tahminleri oku
all_predictions = pd.read_excel(r"C:\Users\Hira\OneDrive\Masaüstü\tum_tahminler.xlsx")
print("\n" + "="*80)
print("KALİTE TAHMİNLERİ ÖZETİ")
print("="*80)
print(f"\nToplam ürün: {len(all_predictions)}")
print(f"UYGUN tahmin edilen: {(all_predictions['TAHMİN'] == 'UYGUN').sum()}")
print(f"UYGUN DEĞİL tahmin edilen: {(all_predictions['TAHMİN'] == 'UYGUN DEĞİL').sum()}")

# Gerçek vs Tahmin uyumu
print(f"\n" + "="*80)
print("GERÇEK vs TAHMİN KARŞILAŞTIRMASI")
print("="*80)
match = all_predictions['KALITE'] == all_predictions['TAHMİN']
print(f"Doğru tahmin: {match.sum()} ({match.sum()/len(all_predictions)*100:.1f}%)")
print(f"Yanlış tahmin: {(~match).sum()} ({(~match).sum()/len(all_predictions)*100:.1f}%)")

# İlk 15 örnek
print(f"\nİlk 15 Ürün Tahmin Sonuçları:")
print("-" * 80)
sample = all_predictions.head(15)[['PARTİ NO', 'ÜRÜN ADI', 'KALITE', 'TAHMİN']]
for idx, row in sample.iterrows():
    match_symbol = "✓" if row['KALITE'] == row['TAHMİN'] else "✗"
    print(f"{match_symbol} {row['PARTİ NO']:>6} | {row['ÜRÜN ADI'][:40]:40} | {row['KALITE']:15} -> {row['TAHMİN']}")

print("\n" + "="*80)
print("✓ Tüm sonuçlar Masaüstü'nde Excel dosyalarına kaydedildi")
print("="*80)
