Sub FilterSheet2BySheet1Left10()
    Dim ws1 As Worksheet, ws2 As Worksheet
    Dim lastRow1 As Long, lastRow2 As Long, i As Long
    Dim dict As Object
    Dim val As Variant, strKey As String
    
    Set ws1 = ActiveWorkbook.Sheets("1")
    Set ws2 = ActiveWorkbook.Sheets("2")
    
    Set dict = CreateObject("Scripting.Dictionary")
    
    lastRow1 = ws1.Cells(ws1.Rows.Count, "C").End(xlUp).Row
    For i = 1 To lastRow1
        val = ws1.Cells(i, "C").Value
        If Not IsError(val) Then
            strKey = Left(Trim(CStr(val)), 10)
            If strKey <> "" Then
                dict(strKey) = True
            End If
        End If
    Next i
    
    lastRow2 = ws2.Cells(ws2.Rows.Count, "D").End(xlUp).Row
    For i = lastRow2 To 1 Step -1
        val = ws2.Cells(i, "D").Value
        If IsError(val) Then
            ws2.Rows(i).Delete
        Else
            strKey = Left(Trim(CStr(val)), 10)
            If Not dict.Exists(strKey) Then
                ws2.Rows(i).Delete
            End If
        End If
    Next i
End Sub
