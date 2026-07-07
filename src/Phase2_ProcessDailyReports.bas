Sub Phase2_ProcessDailyReports()

    Dim wbMaster As Workbook
    Dim wbDaily As Workbook
    Dim wsMaster As Worksheet
    Dim wsData As Worksheet
    Dim wsFiltered As Worksheet
    
    Dim i As Integer, j As Integer
    Dim targetPrefix As String, targetFallback As String
    Dim foundWorkbook As Boolean
    
    Dim lastRow As Long, filterLastRow As Long
    Dim keywords As Variant, k As Variant
    Dim cellValue As String
    Dim matchCol As Long
    Dim filterRange As Range
    Dim col As Variant
    Dim colsToDelete As Variant
    Dim c As Range
    
    ' Set the currently active workbook as the Master file
    Set wbMaster = ActiveWorkbook
    
    ' Portfolio Dummy Targets
    keywords = Array("WAYNE ENTERPRISES", _
                     "STARK INDUSTRIES", _
                     "ACME CORP", _
                     "MASSIVE DYNAMIC", _
                     "CYBERDYNE SYSTEMS", _
                     "UMBRELLA CORP", _
                     "INITECH")
    
    ' Optimize performance
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' Loop through the 7 daily tabs in the Master Workbook
    For i = 1 To wbMaster.Worksheets.Count
        Set wsMaster = wbMaster.Sheets(i)
        
        ' Convert tab name (e.g. "04 Apr 2026") into "mmdd" (e.g. "0404")
        On Error Resume Next
        Dim tabDate As Date
        tabDate = CDate(Trim(wsMaster.Name))
        On Error GoTo 0
        
        targetPrefix = "____" ' Default blank
        If tabDate <> 0 Then
            targetPrefix = Format(tabDate, "mmdd")
        End If
        
        ' Fallback sequence for sequential portfolio testing: maps sheet 1->0221, sheet 2->0222... etc.
        targetFallback = "022" & i
        
        ' 2. Search all OPEN workbooks for matching names or sequential mock names
        foundWorkbook = False
        For Each wbDaily In Application.Workbooks
            If wbDaily.Name <> wbMaster.Name Then
                ' Check if first 4 characters match standard mmdd OR match the 0221 portfolio naming format
                If Left(wbDaily.Name, 4) = targetPrefix Or Left(wbDaily.Name, 4) = targetFallback Then
                    foundWorkbook = True
                    Exit For
                End If
            End If
        Next wbDaily
        
        ' 3. Process the file if found
        If foundWorkbook Then
            Set wsData = Nothing
            
            ' Try fallback to first tab if the custom OPERA string doesn't exist
            On Error Resume Next
            Set wsData = wbDaily.Sheets("ARRIVALLLANDSCAPE_LETTER.RPT")
            If wsData Is Nothing Then Set wsData = wbDaily.Sheets(1)
            On Error GoTo 0
            
            If Not wsData Is Nothing Then
                
                ' --- APPLY FILTERING LOGIC ---
                ' Using Column S for company matching based on your updated file architecture
                lastRow = wsData.Cells(wsData.Rows.Count, "S").End(xlUp).Row
                
                ' Create MatchHelper column in Z
                wsData.Range("Z1").Value = "MatchHelper"
                For j = 2 To lastRow
                    If Not IsError(wsData.Cells(j, "S").Value) Then
                        cellValue = UCase(Trim(CStr(wsData.Cells(j, "S").Value)))
                        wsData.Cells(j, "Z").Value = "" ' Reset
                        
                        For Each k In keywords
                            If InStr(1, cellValue, k, vbTextCompare) > 0 Then
                                wsData.Cells(j, "Z").Value = "MATCH"
                                Exit For
                            End If
                        Next k
                    End If
                Next j
                
                ' Apply filter
                If wsData.AutoFilterMode Then wsData.AutoFilter.ShowAllData
                Set filterRange = wsData.Range("A1:Z" & lastRow)
                filterRange.AutoFilter Field:=26, Criteria1:="MATCH"
                
                ' Create/Replace Filtered_Results sheet
                Application.DisplayAlerts = False
                On Error Resume Next
                wbDaily.Sheets("Filtered_Results").Delete
                On Error GoTo 0
                Application.DisplayAlerts = True
                
                Set wsFiltered = wbDaily.Sheets.Add(After:=wsData)
                wsFiltered.Name = "Filtered_Results"
                
                ' Copy visible filtered rows
                filterRange.SpecialCells(xlCellTypeVisible).Copy Destination:=wsFiltered.Range("A1")
                wsData.AutoFilterMode = False ' Turn off autofilter on original
                
                ' --- FORMAT FILTERED RESULTS ---
                With wsFiltered
                    ' Delete specific columns
                    colsToDelete = Array("AB", "AA", "Y", "X", "W", "V", "Q", "O", "N", "M", "L", "J", "I", "G", "F", "E", "D")
                    For Each col In colsToDelete
                        On Error Resume Next
                        .Columns(col).Delete
                        On Error GoTo 0
                    Next col
                    
                    ' Delete MatchHelper column
                    matchCol = 0
                    For Each c In .Rows(1).Cells
                        If Not IsError(c.Value) And Trim(CStr(c.Value)) = "MatchHelper" Then
                            matchCol = c.Column
                            Exit For
                        End If
                    Next c
                    If matchCol > 0 Then .Columns(matchCol).Delete
                    
                    ' Final structural column adjustments
                    On Error Resume Next
                    .Columns("G").Cut
                    .Columns("K").Insert Shift:=xlToRight
                    Application.CutCopyMode = False
                    .Columns("A").Delete
                    On Error GoTo 0
                    
                    ' --- CHECK AND COPY DATA TO MASTER ---
                    filterLastRow = .Cells(.Rows.Count, "A").End(xlUp).Row
                    
                    If filterLastRow >= 2 Then
                        ' Copy records cleanly over
                        .Range("A2:I" & filterLastRow).Copy
                        wsMaster.Range("B5").PasteSpecial Paste:=xlPasteValues
                        Application.CutCopyMode = False
                    Else
                        ' Blank results logic
                        wsMaster.Range("B13").Value = "There is none for " & wsMaster.Name
                        wsMaster.Range("B13").Interior.ColorIndex = 36 ' Legacy light yellow
                    End If
                End With
            Else
                wsMaster.Range("B13").Value = "There is none for " & wsMaster.Name
                wsMaster.Range("B13").Interior.ColorIndex = 36
            End If
        Else
            wsMaster.Range("B13").Value = "There is none for " & wsMaster.Name & " (File missing)"
            wsMaster.Range("B13").Interior.ColorIndex = 36
        End If
    Next i

    ' Restore performance settings
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
    MsgBox "Phase 2 Complete! Portfolio test sweep finalized.", vbInformation

End Sub
