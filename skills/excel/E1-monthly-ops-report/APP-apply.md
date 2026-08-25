# Uygulama — E1 Aylık Operasyon Raporu

## Adım 1: Power Query bağla
1. Excel > **Data** sekmesi > **Get Data** > **From File** > **From Folder**.
2. Kaynak klasörü seç (aylık export'lar burada birikmeli, standart isimlendirme).
3. **Combine** > **Combine & Transform Data** → Power Query Editor açılır.
4. Adımlar: `Promote Headers` → `Remove Blank Rows` → sütun tiplerini ayarla (Tarih=Date, Tutar=Decimal).
5. **Close & Load** → `RawData` sayfasına (gizli, salt-okunur).

## Adım 2: Katman sayfaları
- `RawData` (gizli) | `Transform` (formüller) | `Report` (yönetim görünümü) | `Settings` (kontrol hücresi).

## Adım 3: Makro (VBA)
`Alt+F11` → Module1'e yapıştır:

```vba
Sub RefreshAndExport()
    On Error GoTo ErrH
    Application.ScreenUpdating = False
    ThisWorkbook.RefreshAll
    ' Async sorgular bitene kadar bekle
    DoEvents
    Dim periad As String: periad = Format(Date, "YYYY-MM")
    Dim dosya As String
    dosya = ThisWorkbook.Path & "\OpsRapor_" & periad & ".pdf"
    Sheets("Report").ExportAsFixedFormat _
        Type:=xlTypePDF, Filename:=dosya, _
        Quality:=xlQualityStandard, IncludeDocProperties:=True, _
        IgnorePrintAreas:=False
    MsgBox "Rapor güncellendi: " & dosya, vbInformation
    Exit Sub
ErrH:
    MsgBox "Yenileme başarısız: " & Err.Description, vbCritical
End Sub
```

## Adım 4: Buton
Insert > Shapes → düğme çiz > `RefreshAndExport` makrosunu ata.

## Doğrulama
Yeni ay dosyasını klasöre at → düğmeye bas → PDF güncellendi mi? RawData'da sütun kayması var mı?
