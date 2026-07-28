Sub ExtractDataFinal()
    Dim wbSource As Workbook, wbNew As Workbook
    Dim wsSource As Worksheet, wsNew As Worksheet
    Dim lastRow As Long, i As Long, j As Long, nextNewRow As Long
    Dim currentNum As Variant
    Dim mergedText As String, currentGroupHeader As String
    Dim fVal As Variant, hVal As Variant
    
    Set wbSource = ActiveWorkbook
    On Error Resume Next
    Set wsSource = wbSource.Sheets("Ведомость")
    On Error GoTo 0
    If wsSource Is Nothing Then Exit Sub
    
    wsSource.Cells.Replace What:="№строки", Replacement:="", LookAt:=xlPart, MatchCase:=False
    wsSource.Cells.UnMerge
    wsSource.Cells.Replace What:="~*", Replacement:="", LookAt:=xlPart
    
    lastRow = wsSource.Cells(wsSource.Rows.Count, "E").End(xlUp).Row
    
    Set wbNew = Workbooks.Add
    Set wsNew = wbNew.Sheets(1)
    nextNewRow = 2
    
    currentGroupHeader = ""
    
    For i = 1 To lastRow
        Dim cellD As Range
        Set cellD = wsSource.Cells(i, "D")
        
        If cellD.Value <> "" Then
            If Not IsNumeric(cellD.Value) Then
                currentGroupHeader = Trim(CStr(cellD.Value))
            Else
                currentNum = cellD.Value
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
                wsNew.Cells(nextNewRow, "C").Value = currentGroupHeader
                wsNew.Cells(nextNewRow, "D").Value = Application.WorksheetFunction.Trim(mergedText)
                wsNew.Cells(nextNewRow, "E").Value = fVal
                wsNew.Cells(nextNewRow, "F").Value = hVal
                nextNewRow = nextNewRow + 1
                
                i = j - 1
            End If
        End If
    Next i
End Sub
