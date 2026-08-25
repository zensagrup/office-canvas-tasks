---
name: excel-monthly-ops-report
description: "Excel'de Power Query + tek makro ile aylik operasyon raporunu otomatikleştir (RefreshAll + PDF)."
version: 1.0.0
author: office-canvas-tasks
license: MIT
---

# Excel — Aylık Operasyon Raporu Otomasyonu

## Trigger
Kullanıcı "her ay aynı raporu manuel birleştiriyorum" / "aylık ops raporunu otomatikleştir" der.

## Request (ajanın vereceği görev)
Aylık operasyon raporunu manuel kopyala-yapıştır yapmadan üret; kaynak klasördeki dosyalar değiştikçe tek tıkla yenilenebilir olsun.

## Setup (başlangıç koşulu)
- Ham veri CSV/klasör halinde geliyor.
- Sunum katmanı ayrı bir sayfada; ham ve hesap mantığı aynı sayfada değil.

## Beklenen davranış (expected)
1. **Power Query**: Data > Get Data > From Folder → klasörü bağla; temizleme adımlarını (başlık yükselt, boş satır sil, tarih tipi) BİR KEZ tanımla.
2. **Katman ayrımı**: RawData (gizli, salt-okunur) / Transform / Report ayrı sayfalar.
3. **Makro**: `Sub RefreshAndExport()` → `ThisWorkbook.RefreshAll`, bekle, rapor sayfasını `ExportAsFixedFormat` ile tarihli PDF yap.
4. Hata yönetimi: eksik dosya/başarısız yenileme için `MsgBox`.

## Uygulama adımları (APP-apply.md bakınız)
## Kaynak
excels.uk, technoexcel
