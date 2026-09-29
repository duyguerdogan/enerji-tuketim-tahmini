## Makine Bazlı Enerji Tüketim Tahmini

Bu proje, geçmiş enerji tüketim verilerini analiz ederek makinelerin aylık enerji tüketimini tahmin etmek amacıyla geliştirilmiştir. Farklı tahmin yöntemleri geçmiş dönemler üzerinde karşılaştırılır ve her makine için en başarılı yaklaşım otomatik olarak seçilir.

Sistem; mevsimsel değişimleri, yakın dönem eğilimlerini, geçen yılın aynı dönemindeki tüketimi ve makinelerin çalışma davranışlarını birlikte değerlendirir. Üretilen tahminler MySQL veritabanına kaydedilir ve gerçek tüketim değerleriyle karşılaştırmalı grafikler oluşturulur.

## Projenin Amacı

Projenin temel amacı, makine bazındaki enerji tüketimini önceden tahmin ederek enerji kullanımının izlenmesini ve planlanmasını kolaylaştırmaktır. Geçen yılın aynı ayına ait verilerin de modele dahil edilebilmesi için geçmiş veri aralığı mevsimsel karşılaştırmaya uygun şekilde ele alınmıştır.

Bu çalışma, gerçek üretim veya tesis verileriyle kullanılabilecek ve yeni verilerle geliştirilebilecek bir enerji tüketimi tahmin prototipidir.

## Kullanılan Teknolojiler

- **Python:** Veri işleme, modelleme ve uygulama akışı
- **Pandas ve NumPy:** Veri temizleme, dönüştürme ve sayısal işlemler
- **XGBoost:** Gecikmeli tüketim değerleri ve zaman özellikleri üzerinden makine öğrenmesi tabanlı tahmin
- **Prophet:** Trend ve mevsimsellik bileşenlerini kullanan zaman serisi tahmini
- **Scikit-learn:** Model değerlendirme, hata metrikleri ve yardımcı makine öğrenmesi işlemleri
- **Matplotlib:** Tahmin ve gerçek tüketim değerlerinin görselleştirilmesi
- **SQLAlchemy ve PyMySQL:** Tahmin sonuçlarının MySQL veritabanına aktarılması
- **OpenPyXL:** Excel veri dosyalarının okunması
- **python-dotenv:** Veritabanı ve çalışma ayarlarının `.env` dosyasından alınması

## Tahmin Yaklaşımları

Uygulama tek bir modele bağlı kalmak yerine birden fazla yaklaşımı karşılaştırır:

- Prophet zaman serisi modeli
- XGBoost regresyon modeli
- Geçen yılın aynı dönemini referans alan tahmin
- Birden fazla modelin sonuçlarını birleştiren ağırlıklı tahmin
- Kesintili çalışan makineler için ON/OFF sınıflandırma ve regresyon yaklaşımı

Modeller geçmiş aylar üzerinde backtest yöntemiyle değerlendirilir. Hata oranlarına göre her makine için en uygun tahmin stratejisi seçilir.

## Veri Yapısı

Uygulama, Excel dosyasındaki tarih-saat ve makine bazlı enerji tüketim sütunlarını kullanır. Veri dosyasının konumu `.env` içerisindeki `DATA_FILE_PATH` değişkeniyle belirtilir.

Beklenen veri yapısı:

- Tarih ve saat bilgisi
- Her makine veya sayaç için ayrı enerji tüketim sütunu
- Düzenli zaman aralıklarında kaydedilmiş tüketim değerleri

Veri seti boyut veya gizlilik nedeniyle repoya eklenmemiştir. Kod, aynı yapıya sahip yeni bir veri setine uyarlanabilir.

## Kurulum

Öncelikle gerekli Python paketlerini yükleyin:

```bash
pip install -r requirements.txt
```

Ardından `.env.example` dosyasını örnek alarak proje klasöründe bir `.env` dosyası oluşturun:

```env
MYSQL_HOST=localhost
MYSQL_PORT=3306
MYSQL_USER=root
MYSQL_PASSWORD=your_password
MYSQL_DATABASE=enerji_tahmin

DATA_FILE_PATH=subat25-mayis26.xlsx
BACKTEST_AYLAR=2
TAHMIN_AYI=2026-04
```

## Çalıştırma

```bash
python energy_forecast.py
```

Çalışma tamamlandığında tahmin sonuçları MySQL veritabanına kaydedilir. Ayrıca makinelerin tahmin edilen ve gerçekleşen tüketim değerlerini gösteren karşılaştırma grafikleri oluşturulur.

## Değerlendirme

Model performansı MAE, RMSE ve MAPE gibi hata metrikleriyle ölçülür. Tahmin ayına ait gerçek veriler mevcut olduğunda tahminler bu değerlerle karşılaştırılarak model başarısı değerlendirilir.
