# Uygulama — E4 Koşullu Birleştirme (VBA)

```vba
Sub MergeLatestByCustomer()
    Dim fso As Object, folder As Object, file As Object
    Dim wbS As Workbook, wsS As Worksheet, wsD As Worksheet
    Dim lr As Long, sr As Long, dr As Long, key, dt
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder("C:\Data\Sources")
    Set wsD = ThisWorkbook.Sheets("Master")
    dr = wsD.Cells(wsD.Rows.Count, 1).End(xlUp).Row + 1
    Application.ScreenUpdating = False
    For Each file In folder.Files
        If LCase(Right(file.Name, 5)) = ".xlsx" Then
            Set wbS = Workbooks.Open(file.Path, ReadOnly:=True)
            Set wsS = wbS.Sheets(1)
            lr = wsS.Cells(wsS.Rows.Count, 1).End(xlUp).Row
            For sr = 2 To lr
                key = wsS.Cells(sr, 1).Value
                dt = wsS.Cells(sr, 2).Value
                ' Master'da yoksa veya daha yeniyse yaz
                Dim cur
                cur = Application.XLookup(key, wsD.Columns(1), wsD.Columns(2), "YOK", 0, -1)
                If cur = "YOK" Or dt > cur Then
                    wsD.Cells(dr, 1).Value = key
                    wsD.Cells(dr, 2).Value = dt
                    dr = dr + 1
                End If
            Next sr
            wbS.Close False
        End If
    Next file
    Application.ScreenUpdating = True
    MsgBox "Birleştirme tamam.", vbInformation
End Sub
```
> Yol: `C:\Data\Sources` ve sütun sırasını kendi verine göre düzelt.
