# Enerji Tuketim Tahmini

Bu proje, gecmis enerji tuketim verilerini kullanarak makine bazinda aylik tuketim tahmini yapmak icin gelistirilmis bir Python uygulamasidir. Sistem, farkli tahmin stratejilerini backtest sonuclariyla karsilastirir ve her makine icin en uygun yaklasimi secerek tahminleri MySQL veritabanina kaydeder.

## Projenin Amaci

Enerji tuketimindeki donemsel davranislari analiz ederek gelecek ay icin makine bazli tahmin uretmek hedeflenir. Kod, ozellikle ayni ayin gecen yilki tuketim davranisini, yakin donem trendlerini ve makine bazli kullanim aliskanliklarini birlikte degerlendirir.

Bu calisma, gercek uretim veya tesis verileriyle genisletilebilecek bir enerji tahmini prototipi olarak dusunulmustur.

## Kullanilan Yontemler

- Prophet ile zaman serisi tahmini
- XGBoost ile ozellik tabanli tahmin
- Gecen yil ayni zaman araligi referansi
- Rolling blend yaklasimi
- Acik/kapali calisma davranisi icin ON/OFF stratejisi
- Backtest ile makine bazinda otomatik strateji secimi

## Veri Yapisi

Uygulama Excel dosyasindan tarih/saat ve makine bazli tuketim verilerini okur. Veri dosyasi `.env` icindeki `DATA_FILE_PATH` degiskeniyle belirtilir.

Beklenen genel yapi:

- Tarih/saat kolonu
- Makine veya sayac bazli enerji tuketim kolonlari
- Duzenli zaman araliklariyla olculmus tuketim degerleri

## Kurulum

```bash
pip install -r requirements.txt
```

`.env.example` dosyasini temel alarak proje klasorunde `.env` dosyasi olusturun:

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

## Calistirma

```bash
python energy_forecast.py
```

Calisma sonunda tahmin sonuclari MySQL veritabanina yazilir ve tahmin/gercek tuketim karsilastirma grafikleri uretilir.

## Notlar

- `.env` dosyasi sifre icerebilecegi icin GitHub'a yuklenmemelidir.
- Veri dosyalari boyut ve gizlilik nedeniyle repoya eklenmeyebilir; ornek veri veya veri formati aciklamasi yeterlidir.
- Backtest sonuclari, tahmin yapilan ayin gercek verisi varsa hesaplanir. Henuz gerceklesmemis bir ay icin calistirilacaksa gelecege donuk tarih araligi uretimi ayrica eklenebilir.
