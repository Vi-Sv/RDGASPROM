Sub ExtractDataFromVed()
    Dim wbSource As Workbook, wbNew As Workbook
    Dim wsSource As Worksheet, wsNew As Worksheet
    Dim lastRow As Long, i As Long, j As Long
    Dim nextNewRow As Long
    Dim currentNum As Double
    Dim mergedText As String
    Dim fVal As Variant, hVal As Variant
    
    Set wbSource = ActiveWorkbook
    On Error Resume Next
    Set wsSource = wbSource.Sheets("Ведомость")
    On Error GoTo 0
    If wsSource Is Nothing Then Exit Sub
    
    wsSource.Cells.UnMerge
    wsSource.Cells.Replace What:="~*", Replacement:="", LookAt:=xlPart
    
    lastRow = wsSource.Cells(wsSource.Rows.Count, "E").End(xlUp).Row
    With wsSource.Range("D1:D" & lastRow)
        .NumberFormat = "0"
        .Value = .Value
    End With
    
    Set wbNew = Workbooks.Add
    Set wsNew = wbNew.Sheets(1)
    nextNewRow = 2
    
    For i = 1 To lastRow
        If wsSource.Cells(i, "D").Value <> "" And IsNumeric(wsSource.Cells(i, "D").Value) Then
            currentNum = wsSource.Cells(i, "D").Value
            fVal = wsSource.Cells(i, "F").Value
            hVal = wsSource.Cells(i, "H").Value
            mergedText = Trim(CStr(wsSource.Cells(i, "E").Value))
            
            j = i + 1
            Do While j <= lastRow And wsSource.Cells(j, "D").Value = ""
                If Trim(CStr(wsSource.Cells(j, "E").Value)) <> "" Then
                    mergedText = mergedText & " " & Trim(CStr(wsSource.Cells(j, "E").Value))
                End If
                j = j + 1
            Loop
            
            wsNew.Cells(nextNewRow, "B").Value = currentNum
            wsNew.Cells(nextNewRow, "C").Value = Application.WorksheetFunction.Trim(mergedText)
            wsNew.Cells(nextNewRow, "D").Value = fVal
            wsNew.Cells(nextNewRow, "E").Value = hVal
            nextNewRow = nextNewRow + 1
            
            i = j - 1
        End If
    Next i
End Sub
