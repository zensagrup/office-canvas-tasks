---
name: excel-weekly-kpi-copilot
description: "Excel'de Power Query hazırlar, VBA PDF basar, Copilot 3 maddelik özet yazar."
version: 1.0.0
author: office-canvas-tasks
license: MIT
---

# Excel — Haftalık KPI + AI Özet

## Trigger
"Haftalık KPI raporunu ve üst yönetim özeti otomatikleştir."

## Request
Haftalık KPI'ları hesapla ve üst yönetim için 3 maddelik özet üret.

## Setup
RawData / WeeklyReport / Summary / Settings sayfaları ayrı; KPI'lar Excel formüllerinde.

## Expected
1. Power Query veriyi hazırlar.
2. VBA `RefreshAll` + pivot + PDF export.
3. Copilot (veya AI): "ana trendler, anomaliler, en iyi bölgeler" → 3 madde.

## Kaynak
Hakky Handbook
