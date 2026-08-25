# Uygulama — E6 Doğrulama Kapısı

`Settings` sayfasına kontrol hücreleri:
| Hücre | Anlamı | Örnek |
|---|---|---|
| Kontrol_Satir | Bu ay satır sayısı | =COUNTA(RawData!A:A) |
| Kontrol_Tarih | En yeni veri tarihi | =MAX(RawData!Tarih) |
| Esik_Satir | Beklenen min | 100 |
| Esik_Varyans | % sapma | 0.3 |

```vba
Sub KontrolVeGonder()
    Dim satir, esik, varyans
    satir = Range("Kontrol_Satir").Value
    esik = Range("Esik_Satir").Value
    If satir < esik Then
        MsgBox "Satır sayısı eşik altında (" & satir & "<" & esik & "). Gönderim iptal.", vbCritical
        Exit Sub
    End If
    ' Tarih 30 günden eski mi?
    If Date - Range("Kontrol_Tarih").Value > 30 Then
        MsgBox "Veri bayat (>30 gün). Gönderim iptal.", vbCritical
        Exit Sub
    End If
    MsgBox "Kontrol geçti, rapor gönderilebilir.", vbInformation
End Sub
```
> "Kör otomasyon" yerine: hata varsa DUR, gönderme yapılmaz.
