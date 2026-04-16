# karsilastirma.py / kalite_kontrol_sonuclari.xlsx
Üretim verilerini spekt limitleriyle kıyaslayarak ürünleri "Uygun" veya "Uygun Değil" şeklinde etiketledim. Bu işlemi yaparak modeli eğitmek için gereken cevap anahtarını oluşturdum. Sonuç olarak her partinin kalite durumunu belirten Excel dosyasını aldım.

# model_egitme.py / ml_model_results.xlsx
Etiketlediğim bu verileri kullanarak XGBoost modelini eğittim. Modelin geçmiş verilerdeki kalite desenini öğrenmesini sağladım. Sonuçta modelin %86.6 doğrulukla tahmin yapabildiğini ve kaliteyi en çok "Yoğunluk" ile "Alkol" parametrelerinin etkilediğini saptadım.

# sonuc_dosyasi.py / tum_tahminler.xlsx
Eğittiğim modeli tüm veri seti üzerinde çalıştırarak analiz yapmasını istedim. Modelin her bir ürün için "Uygun" veya "Uygun Değil" tahminlerini yapmasını amaçladım. Sonuçta gerçek değerlerle model tahminlerini yan yana getiren kapsamlı bir excel raporu elde ettim.

# Özetle: 
Manuel kalite kontrol sürecini makine öğrenmesine öğreterek, ürün hatalarını önceden tespit edebilen dijital bir sistem kurmuş oldum.
# bitirme
