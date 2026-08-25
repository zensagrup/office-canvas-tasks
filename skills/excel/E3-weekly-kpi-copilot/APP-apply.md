# Uygulama — E3 Haftalık KPI + Copilot

## Sayfa rolleri
| Sayfa | Rol |
|---|---|
| RawData | Gizli, Power Query kaynağı (değiştirme!) |
| WeeklyReport | Pivot/Chart/KPI görünümü |
| Summary | Copilot çıktısının yapıştırıldığı sabit alan |
| Settings | Yollar, e-posta, sorgu adları |

## VBA (WeeklyReportAutomation)
```vba
Sub WeeklyReportAutomation()
    ThisWorkbook.RefreshAll
    DoEvents
    Dim f As String: f = ThisWorkbook.Path & "\Weekly_" & Format(Date, "YYYYMMDD") & ".pdf"
    Sheets("WeeklyReport").ExportAsFixedFormat xlTypePDF, f
    MsgBox "Haftalık rapor basıldı.", vbInformation
End Sub
```

## Copilot özet promptu
> Bu haftalık KPI tablosunu özetle: (1) ana trendler, (2) anomali sayısı, (3) öne çıkan bölgeler. Ham sayıları ekle.

Çıktıyı Summary!A1 / A4 / A7'e yapıştır.
