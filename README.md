# Office & Canvas Görev Senaryoları

Excel, Word ve Canvas (Canva / Microsoft 365 Copilot Pages / Microsoft Whiteboard) için
pratik, çalıştırılabilir görev senaryoları. Her senaryo **request / setup / expected** yapısındadır.

## Yapı
```
office-canvas-tasks/
├── excel/scenarios.json     # Power Query / VBA / Copilot
├── word/scenarios.json       # Mail Merge (mektup, sertifika, etiket, e-posta, sözleşme)
└── canvas/scenarios.json     # Canva + M365 Copilot Pages + Whiteboard
```

## Senaryo şeması
```json
{
  "id": "benzersiz-id",
  "title": "Başlık",
  "request": "Yapılması istenen görev",
  "setup": "Başlangıç koşulu / girdi",
  "expected": "Doğru sonuç / beklenen davranış",
  "source": "Kaynak (internet araştırması)"
}
```

## Kullanım
Bu liste bir **eval veya checklist** olarak kullanılabilir:
- Her senaryoyu gerçek araçta (Excel/Word/Canva/Whiteboard) deneyin.
- `expected` alanı "doğru davranış" kriteridir; sapma varsa not alın.
- `setup` alanı, ajan/senaryo çalıştırıcısına verilmemesi gereken (gizli) kriterleri içermez;
  burada yalnızca gerçekçi bir başlangıç durumu tanımlanır.

## Lisans
MIT
