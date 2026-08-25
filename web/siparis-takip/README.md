# Sipariş Takip Web Uygulaması

FastAPI + openpyxl tabanlı, web arayüzünden veri girip Excel'e kaydeden sistem.

## Çalıştırma
```bash
pip install -r requirements.txt
cd "G:\Drive'ım\SİPARİŞ TAKİP"
python main.py
# Tarayıcı: http://127.0.0.1:8000
```

## Özellikler
- Web'den ekle / düzenle / sil
- Veriler `siparis.xlsx` içinde saklanır (arka planda openpyxl)
- KALAN PARA = VERİLEN - (BİRİM×ADET + ÇANTA×ADET)
- KALAN ÖDEME = VERİLEN - ALINAN ÖDEME
- Canlı özet (toplam sipariş, verilen/alınan ödeme, kalanlar)

## Not
Gerçek veri `G:\Drive'ım\SİPARİŞ TAKİP\siparis.xlsx` içindedir (web uygulaması orada çalışır).
