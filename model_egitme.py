import pandas as pd
import numpy as np
from pathlib import Path
import glob
from sklearn.model_selection import train_test_split
from sklearn.preprocessing import LabelEncoder
from sklearn.metrics import classification_report, confusion_matrix, accuracy_score
import xgboost as xgb
import warnings
warnings.filterwarnings('ignore')

# Dosyayı bul
desktop = Path(r"C:\Users\Hira\OneDrive\Masaüstü")
files = glob.glob(str(desktop / 'FR.KP*'))
excel_path = files[0]

print("="*80)
print("ODA KOKUSU KALİTE KONTROL - MAKİNE ÖĞRENMESİ (XGBoost)")
print("="*80)

# ===== VERİ YÜKLEME =====
df_raw = pd.read_excel(excel_path, sheet_name='ODA KOKULARI', header=None)

# Sütun konumları (indeks ile)
party_idx = 0
date_idx = 1
stock_idx = 2
product_idx = 3
appearance_idx = 4
color_idx = 5
aroma_idx = 6
ph_idx = 7
alcohol_idx = 8
refractive_idx = 9
density_idx = 10
comment_idx = 11

# Başlıkları al (satır 1'den)
headers_row1 = df_raw.iloc[1].values
headers = [
    'PARTİ NO',
    'TARİH',
    'STOK KODU',
    'ÜRÜN ADI',
    'GÖRÜNÜŞ',
    'RENKLİ RENKSİZ',
    'KOKU',
    'pH',
    'ALKOL',
    'KIRILMA',
    'YOGUNLUK',
    'AÇIKLAMA'
]

# Veri yükleme (satır 2'den)
df = df_raw.iloc[2:].copy()
df.columns = headers
df = df[df['PARTİ NO'].notna()].reset_index(drop=True)

print(f"\nToplam veri: {len(df)} satır")
print(f"Sütunlar: {list(df.columns)}\n")

# ===== VERİ TEMIZLEME VE ÖZELLİK ÇIKARMA =====
df_clean = df.copy()

# Numrik sütunları dönüştür
numric_cols = {'pH': 'pH', 'ALKOL': 'Alkol', 'KIRILMA': 'Kirilma', 'YOGUNLUK': 'Yogunluk'}
for old_col, new_col in numric_cols.items():
    df_clean[new_col] = pd.to_numeric(
        df_clean[old_col].astype(str).str.replace(',', '.'), 
        errors='coerce'
    )

# Kategorik özellikler
df_clean['GÖRÜŞ_KATEGORİ'] = df_clean['GÖRÜNÜŞ'].astype(str).str.strip()
df_clean['KOKU_KATEGORİ'] = df_clean['KOKU'].astype(str).str.strip()
df_clean['RENKKATEGORİ'] = df_clean['RENKLİ RENKSİZ'].astype(str).str.strip()

# Eksik verileri kontrol et
print("Eksik verileri doldurma...")
for col in ['pH', 'Alkol', 'Kirilma', 'Yogunluk']:
    missing = df_clean[col].isna().sum()
    if missing > 0:
        df_clean[col].fillna(df_clean[col].median(), inplace=True)
        print(f"  {col}: {missing} eksik değer düzeltildi")

# ===== KALITE ETIKETI (TARGET) OLUŞTURMA =====
# Spekt listesinden standartlar
standards = {
    'Alkol': (80, 95),
    'Yogunluk': (0.8, 0.9),
    'Kirilma': (1.37, 1.39),
    'pH': (2.0, 6.0)
}

def kalite_kontrolu(row):
    """Her parametreyi kontrol et ve genel kaliteyi belirle"""
    checks = []
    for param, (min_val, max_val) in standards.items():
        if pd.notna(row[param]):
            checks.append(min_val <= row[param] <= max_val)
    
    if not checks:
        return 'VERİ HATASI'
    return 'UYGUN' if all(checks) else 'UYGUN DEĞİL'

df_clean['KALITE'] = df_clean.apply(kalite_kontrolu, axis=1)

print(f"\nKalite Dağılımı:")
print(df_clean['KALITE'].value_counts())

# ===== MAKİNE ÖĞRENMESİ HAZIRLIĞI =====
# Sadece geçerli verileri kullan
df_ml = df_clean[df_clean['KALITE'] != 'VERİ HATASI'].copy()

print(f"\nMakine öğrenme veri seti: {len(df_ml)} satır")

# Features
features = ['pH', 'Alkol', 'Kirilma', 'Yogunluk']
X = df_ml[features].copy()

# Target
y = df_ml['KALITE'].copy()

# Label encoding for target
le = LabelEncoder()
y_encoded = le.fit_transform(y)

print(f"\nÖzellikleri kullanılan parametreler:")
for feat in features:
    print(f"  - {feat} (min: {X[feat].min():.3f}, max: {X[feat].max():.3f})")

print(f"\nTarget sınıfları: {dict(zip(le.classes_, np.unique(y_encoded)))}")

# ===== TRAIN-TEST BÖLÜMÜ =====
X_train, X_test, y_train, y_test = train_test_split(
    X, y_encoded, test_size=0.2, random_state=42, stratify=y_encoded
)

# Test verisinin indekslerini kaydet (ürün kodu ve adı için)
test_indices = X_test.index

print(f"\nTrain set: {len(X_train)} satır")
print(f"Test set: {len(X_test)} satır")

# ===== XGBoost MODELİ =====
print("\n" + "="*80)
print("XGBoost Modeli Eğitiliyor...")
print("="*80)

model = xgb.XGBClassifier(
    n_estimators=100,
    max_depth=5,
    learning_rate=0.1,
    subsample=0.8,
    colsample_bytree=0.8,
    random_state=42,
    verbosity=0
)

model.fit(X_train, y_train)

# ===== DEĞERLENDİRME =====
y_pred = model.predict(X_test)
accuracy = accuracy_score(y_test, y_pred)

print(f"\n✓ Model başarıyla eğitildi!")
print(f"Doğruluk (Accuracy): {accuracy:.4f} ({accuracy*100:.2f}%)")

print(f"\nDetaylı Sonuçlar:")
print(classification_report(y_test, y_pred, target_names=le.classes_))

print(f"\nKarmaşa Matrisi (Confusion Matrix):")
cm = confusion_matrix(y_test, y_pred)
print(cm)

# ===== ÖZELLİK ÖNEMLİLİĞİ =====
print(f"\nÖzellik Önem Sıralaması:")
feature_importance = pd.DataFrame({
    'Özellik': features,
    'Önem': model.feature_importances_
}).sort_values('Önem', ascending=False)
print(feature_importance.to_string(index=False))

# ===== SONUÇLARI KAYDET =====
output_file = r"C:\Users\Hira\OneDrive\Masaüstü\ml_model_results.xlsx"

with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
    # Test sonuçları - ürün bilgileri ile
    test_info = df_ml.loc[test_indices, ['PARTİ NO', 'ÜRÜN ADI']].reset_index(drop=True)
    results_df = pd.DataFrame({
        'PARTİ NO': test_info['PARTİ NO'].values,
        'ÜRÜN ADI': test_info['ÜRÜN ADI'].values,
        'Gerçek': le.inverse_transform(y_test),
        'Tahmin': le.inverse_transform(y_pred),
        'pH': X_test['pH'].values,
        'Alkol': X_test['Alkol'].values,
        'Kirilma': X_test['Kirilma'].values,
        'Yogunluk': X_test['Yogunluk'].values
    }).reset_index(drop=True)
    results_df.to_excel(writer, sheet_name='Test Sonuçları', index=False)
    
    # Özet istatistikler
    summary_df = pd.DataFrame({
        'Metrik': ['Toplam Test Verisi', 'Doğru Tahmin', 'Yanlış Tahmin', 'Doğruluk (%)'],
        'Değer': [len(y_test), (y_pred == y_test).sum(), (y_pred != y_test).sum(), f'{accuracy*100:.2f}']
    })
    summary_df.to_excel(writer, sheet_name='Özet', index=False)
    
    # Özellik önemi
    feature_importance.to_excel(writer, sheet_name='Özellik Önemi', index=False)
    
    # Sınıf dağılımı
    class_dist = pd.DataFrame({
        'Sınıf': le.classes_,
        'Sayı': [np.sum(y_test == i) for i in range(len(le.classes_))]
    })
    class_dist.to_excel(writer, sheet_name='Sınıf Dağılımı', index=False)

print(f"\n✓ Sonuçlar kaydedildi: {output_file}")

# ===== TÜM VERİ ÜZERİNDE TAHMİN =====
print("\n" + "="*80)
print("Tüm Veri Seti Üzerinde Tahmin Yapılıyor...")
print("="*80)

X_all = df_ml[features].copy()
y_all_pred = model.predict(X_all)
y_all_pred_proba = model.predict_proba(X_all)

df_ml['TAHMİN'] = le.inverse_transform(y_all_pred)
for idx, class_name in enumerate(le.classes_):
    df_ml[f'{class_name}_OLASILIK'] = y_all_pred_proba[:, idx]

# Gerçek vs Tahmin karşılaştırması
comparison_df = df_ml[[
    'PARTİ NO', 'ÜRÜN ADI', 'KALITE', 'TAHMİN', 
    'pH', 'Alkol', 'Kirilma', 'Yogunluk'
]].copy()

# Sonuçları kaydet
output_predictions = r"C:\Users\Hira\OneDrive\Masaüstü\tum_tahminler.xlsx"
comparison_df.to_excel(output_predictions, index=False)
print(f"✓ Tüm tahminler kaydedildi: {output_predictions}")

print(f"\nİlk 10 Tahmin Sonucu:")
print(comparison_df.head(10).to_string(index=False))

print("\n" + "="*80)
print("İŞLEM TAMAMLANDI")
print("="*80)
