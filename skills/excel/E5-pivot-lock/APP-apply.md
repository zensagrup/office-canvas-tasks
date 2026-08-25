# Uygulama — E5 Pivot Kilitleme

1. Pivot'u seç → **PivotTable Analyze** > **Options** > **Total & Filters** sekmesi.
2. **AutoSort** kutusunu işaretsiz yap (alanlar kendiliğinden sıralanmasın).
3. **Data** sekmesi > **Refresh data when opening the file** işaretle.
4. Power Query çıktısını pivotun veri kaynağı yap: Pivot'a sağ tık → **Change Data Source** → RawData tablosu.
5. VBA ile pivot yenileme:
```vba
Sub RefreshPivots()
    Dim ws As Worksheet, pt As PivotTable
    For Each ws In ThisWorkbook.Worksheets("WeeklyReport").Parent.Worksheets
        For Each pt In ws.PivotTables
            pt.RefreshTable
        Next pt
    Next ws
End Sub
```
