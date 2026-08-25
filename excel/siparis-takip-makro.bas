Attribute VB_Name = "SiparisTakipMakro"
Option Explicit

' ============================================================
' Sipariş Takip Otomasyonu — Yenile + PDF Export
' Açıklama: "Siparişler" sayfasındaki formülleri yeniler,
'           "Özet" sayfasını PDF olarak dışa aktarır.
' ============================================================

Sub YenileVeExport()
    On Error GoTo ErrH
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationAutomatic
    
    ' 1) Tüm sayfaları yenile (formüller güncellensin)
    ThisWorkbook.Sheets("Siparişler").Calculate
    ThisWorkbook.Sheets("Özet").Calculate
    ThisWorkbook.RefreshAll
    
    ' 2) Özet sayfasını PDF yap
    Dim dosya As String
    Dim periad As String
    periad = Format(Date, "YYYY-MM-DD")
    dosya = ThisWorkbook.Path & "\Siparis_Ozet_" & periad & ".pdf"
    
    ThisWorkbook.Sheets("Özet").ExportAsFixedFormat _
        Type:=xlTypePDF, _
        Filename:=dosya, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, _
        IgnorePrintAreas:=False
    
    Application.ScreenUpdating = True
    MsgBox "Yenileme tamam. Özet PDF:" & vbNewLine & dosya, vbInformation, "Sipariş Takip"
    Exit Sub

ErrH:
    Application.ScreenUpdating = True
    MsgBox "Hata: " & Err.Description, vbCritical, "Sipariş Takip"
End Sub

' ------------------------------------------------------------
' Yeni sipariş satırı ekle (formülleri otomatik kopyalar)
' ------------------------------------------------------------
Sub YeniSiparisEkle()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Siparişler")
    
    Dim sonSatir As Long
    sonSatir = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    Dim yeniSatir As Long
    yeniSatir = sonSatir + 1
    
    ' Başlıkları referans alarak formülleri kopyala (KALAN PARA ve KALAN ÖDEME)
    ws.Cells(yeniSatir, 10).Formula = "=I" & yeniSatir & "-(G" & yeniSatir & "*F" & yeniSatir & "+H" & yeniSatir & "*F" & yeniSatir & ")"
    ws.Cells(yeniSatir, 12).Formula = "=I" & yeniSatir & "-K" & yeniSatir & ""
    
    ' Para formatı
    Dim rng As Range
    Set rng = ws.Range(ws.Cells(yeniSatir, 7), ws.Cells(yeniSatir, 12))
    rng.NumberFormat = "#,##0.00 ₺"
    ws.Cells(yeniSatir, 2).NumberFormat = "DD.MM.YYYY"
    
    ws.Activate
    ws.Cells(yeniSatir, 1).Select
    MsgBox "Yeni satır eklendi: " & yeniSatir & vbNewLine & _
           "KALAN PARA ve KALAN ÖDEME formülleri otomatik kopyalandı.", _
           vbInformation, "Sipariş Takip"
End Sub
