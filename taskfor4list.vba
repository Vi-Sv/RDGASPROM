Sub FilterAndCollectMismatchesBothSheets()
    Dim wb As Workbook
    Dim ws1 As Worksheet, ws2 As Worksheet, ws3 As Worksheet, ws4 As Worksheet
    Dim lastRow1 As Long, lastRow2 As Long, i As Long, nextRow3 As Long, nextRow4 As Long
    Dim dict1 As Object, dict2 As Object
    Dim val As Variant, strKey As String
    
    Set wb = ActiveWorkbook
    Set ws1 = wb.Sheets("1")
    Set ws2 = wb.Sheets("2")
    
    On Error Resume Next
    Set ws3 = wb.Sheets("3")
    If ws3 Is Nothing Then
        Set ws3 = wb.Sheets.Add(After:=ws2)
        ws3.Name = "3"
    End If
    Set ws4 = wb.Sheets("4")
    If ws4 Is Nothing Then
        Set ws4 = wb.Sheets.Add(After:=ws3)
        ws4.Name = "4"
    End If
    On Error GoTo 0
    
    Set dict1 = CreateObject("Scripting.Dictionary")
    Set dict2 = CreateObject("Scripting.Dictionary")
    
    lastRow1 = ws1.Cells(ws1.Rows.Count, "C").End(xlUp).Row
    For i = 1 To lastRow1
        val = ws1.Cells(i, "C").Value
        If Not IsError(val) Then
            strKey = Left(Trim(CStr(val)), 10)
            If strKey <> "" Then
                dict1(strKey) = True
            End If
        End If
    Next i
    
    lastRow2 = ws2.Cells(ws2.Rows.Count, "D").End(xlUp).Row
    For i = 1 To lastRow2
        val = ws2.Cells(i, "D").Value
        If Not IsError(val) Then
            strKey = Left(Trim(CStr(val)), 10)
            If strKey <> "" Then
                dict2(strKey) = True
            End If
        End If
    Next i
    
    nextRow4 = ws4.Cells(ws4.Rows.Count, "A").End(xlUp).Row
    If ws4.Range("A" & nextRow4).Value <> "" Then nextRow4 = nextRow4 + 1
    If nextRow4 = 1 Then nextRow4 = 2
    
    For i = 1 To lastRow1
        val = ws1.Cells(i, "C").Value
        Dim missingInSheet2 As Boolean
        missingInSheet2 = False
        
        If IsError(val) Then
            missingInSheet2 = True
        Else
            strKey = Left(Trim(CStr(val)), 10)
            If Not dict2.Exists(strKey) Then missingInSheet2 = True
        End If
        
        If missingInSheet2 Then
            ws1.Rows(i).Copy Destination:=ws4.Rows(nextRow4)
            nextRow4 = nextRow4 + 1
        End If
    Next i
    
    nextRow3 = ws3.Cells(ws3.Rows.Count, "A").End(xlUp).Row
    If ws3.Range("A" & nextRow3).Value <> "" Then nextRow3 = nextRow3 + 1
    If nextRow3 = 1 Then nextRow3 = 2
    
    For i = lastRow2 To 1 Step -1
        val = ws2.Cells(i, "D").Value
        Dim missingInSheet1 As Boolean
        missingInSheet1 = False
        
        If IsError(val) Then
            missingInSheet1 = True
        Else
            strKey = Left(Trim(CStr(val)), 10)
            If Not dict1.Exists(strKey) Then missingInSheet1 = True
        End If
        
        If missingInSheet1 Then
            ws2.Rows(i).Copy Destination:=ws3.Rows(nextRow3)
            ws2.Rows(i).Delete
            nextRow3 = nextRow3 + 1
        End If
    Next i
End Sub
