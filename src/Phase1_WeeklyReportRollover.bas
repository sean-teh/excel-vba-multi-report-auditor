Sub Phase1_WeeklyReportRollover()

    Dim wb As Workbook
    Dim ws As Worksheet
    Dim oldStartDate As Date
    Dim newStartDate As Date
    Dim newEndDate As Date
    Dim newFileName As String
    Dim dateStr As String
    Dim curDate As Date
    Dim lastRow As Long
    Dim lastCol As Long
    Dim clearRng As Range
    Dim i As Integer
    
    Set wb = ActiveWorkbook
    
    ' --------------------------------------------------------
    ' 1 & 2. Calculate New Dates & Duplicate File via SaveAs
    ' --------------------------------------------------------
    ' Read the date from the first tab (e.g., "28 Mar 2026")
    On Error Resume Next
    oldStartDate = CDate(wb.Sheets(1).Name)
    On Error GoTo 0
    
    If oldStartDate = 0 Then
        MsgBox "Could not read the date from the first tab ('" & wb.Sheets(1).Name & "'). Please ensure it is a valid date.", vbCritical
        Exit Sub
    End If
    
    ' Calculate the next 7 days
    newStartDate = DateAdd("d", 7, oldStartDate)
    newEndDate = DateAdd("d", 6, newStartDate)
    
    ' Format the new file name (e.g., ...0404 - 1004)
    newFileName = wb.Path & "\Double Points Strategy Arrival List - " & _
                  Format(newStartDate, "ddmm") & " - " & Format(newEndDate, "ddmm")
    
    ' Save the file as a Macro-Enabled Workbook (.xlsm)
    ' This acts as the "duplicate", leaving your old file untouched
    Application.DisplayAlerts = False
    wb.SaveAs Filename:=newFileName & ".xlsm", FileFormat:=xlOpenXMLWorkbookMacroEnabled
    Application.DisplayAlerts = True

    ' --------------------------------------------------------
    ' 3 & 4. Loop through the 7 Daily Tabs to Rename & Clean
    ' --------------------------------------------------------
   For i = 1 To wb.Worksheets.Count
        Set ws = wb.Sheets(i)
        
        ' Calculate the specific date for this tab
        curDate = DateAdd("d", i - 1, newStartDate)
        dateStr = Format(curDate, "dd mmm yyyy") ' e.g., "04 Apr 2026"
        
        ' 3a. Rename Tab
        ws.Name = dateStr
        
        ' 3b. Update Cell B2
        ws.Range("B2").Value = "Arrival List - " & dateStr
        
        ' 4. Clean data and highlights from Row 5 downwards
        lastRow = ws.UsedRange.Rows.Count + ws.UsedRange.Row - 1
        lastCol = ws.UsedRange.Columns.Count + ws.UsedRange.Column - 1
        
        If lastRow >= 5 Then
            ' Define the range to clear (Row 5 down to the last used cell)
            Set clearRng = ws.Range(ws.Cells(5, 1), ws.Cells(lastRow, lastCol))
            
            ' Clear text and numbers
            clearRng.ClearContents
            
            ' Clear highlights (fill color) while leaving borders/fonts intact
            clearRng.Interior.ColorIndex = xlNone
        End If
    Next i
    
    ' Final save to lock in the clean up
    wb.Save
    
    MsgBox "Phase 1 Complete!" & vbCrLf & vbCrLf & _
           "File duplicated and saved as: " & wb.Name, vbInformation

End Sub
