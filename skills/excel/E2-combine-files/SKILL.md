---
name: excel-combine-files
description: "Power Query 'From Folder' ile çoklu Excel/CSV dosyasını tek tıkla birleştir."
version: 1.0.0
author: office-canvas-tasks
license: MIT
---

# Excel — Çoklu Dosyayı Birleştirme

## Trigger
"12 bölgesel dosyayı tek master'a birleştiriyorum" / "klasördeki dosyaları topla".

## Request
12 bölgesel dosyayı tek master tabloya birleştir; yeni dosya eklenince yeniden çalışsın.

## Setup
Aynı şemalı (.xlsx/.csv) dosyalar bir klasörde; sütunlar aynı sırada/isimde.

## Expected
Power Query `From Folder → Combine Files` ile içe aktar; sonraki aylarda yalnızca **Refresh**. VBA gerekmez.

## Kaynak
codeandcrumbs
