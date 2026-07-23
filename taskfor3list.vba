Sub FilterSheet2AndMoveMismatches()
    Dim wb As Workbook
    Dim ws1 As Worksheet, ws2 As Worksheet, ws3 As Worksheet
    Dim lastRow1 As Long, lastRow2 As Long, i As Long, nextRow3 As Long
    Dim dict As Object
    Dim val As Variant, strKey As String
    
    Set wb = ActiveWorkbook
    Set ws1 = wb.Sheets("1")
    Set ws2 = wb.Sheets("2")
    
    On Error Resume Next
    Set ws3 = wb.Sheets("3")
    On Error GoTo 0
    If ws3 Is Nothing Then
        Set ws3 = wb.Sheets.Add(After:=ws2)
        ws3.Name = "3"
    End If
    
    Set dict = CreateObject("Scripting.Dictionary")
    
    lastRow1 = ws1.Cells(ws1.Rows.Count, "C").End(xlUp).Row
    For i = 1 To lastRow1
        val = ws1.Cells(i, "C").Value
        If Not IsError(val) Then
            strKey = Left(Trim(CStr(val)), 10)
            if strKey <> "" Then
                dict(strKey) = True
            End If
        End If
    Next i
    
    nextRow3 = ws3.Cells(ws3.Rows.Count, "A").End(xlUp).Row
    If ws3.Range("A" & nextRow3).Value <> "" Then nextRow3 = nextRow3 + 1
    If nextRow3 = 1 Then nextRow3 = 2
    
    lastRow2 = ws2.Cells(ws2.Rows.Count, "D").End(xlUp).Row
    For i = lastRow2 To 1 Step -1
        val = ws2.Cells(i, "D").Value
        Dim shouldMove As Boolean
        shouldMove = False
        
        If IsError(val) Then
            shouldMove = True
        Else
            strKey = Left(Trim(CStr(val)), 10)
            If Not dict.Exists(strKey) Then
                shouldMove = True
            End If
        End If
        
        If shouldMove Then
            ws2.Rows(i).Copy Destination:=ws3.Rows(nextRow3)
            ws2.Rows(i).Delete
            nextRow3 = nextRow3 + 1
        End If
    Next i
End Sub
