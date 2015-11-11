Attribute VB_Name = "combine"
' Global variables
Public Const THIS_FILE = "combine.xlsm"
Public Const RAW_FILE = "raw.xlsm"
Public COMBINE As Workbook

Dim ALL_CURRENT As Integer

Sub sooort(sheet As Worksheet, sort_index1 As Integer, sort_index2 As Integer, sort_index3 As Integer)
    sheet.UsedRange.sort key1:=sheet.Columns(sort_index1), Header:=xlYes, key2:=sheet.Columns(sort_index2), key3:=sheet.Columns(sort_index3)
End Sub

Sub soort(sheet As Worksheet, sort_index1 As Integer, sort_index2 As Integer)
    sheet.UsedRange.sort key1:=sheet.Columns(sort_index1), Header:=xlYes, key2:=sheet.Columns(sort_index2)
End Sub

Sub main()
    ALL_CURRENT = 2
    Set COMBINE = Workbooks(THIS_FILE)

    Dim RAW As Workbook
    Set RAW = Workbooks.Open(ActiveWorkbook.Path & Application.PathSeparator & RAW_FILE)

    Call clean

    Call phaseOne(RAW.Worksheets("SA"), RAW.Worksheets("Lead"), COMBINE.Worksheets("combine"))
    Call phaseTwo(RAW.Worksheets("Lead"), COMBINE.Worksheets("combine"))
    Call phaseThree(RAW.Worksheets("opp"), COMBINE.Worksheets("combine"))
    Call sooort(COMBINE.Worksheets("combine"), 1, 2, 4)
    Call combineA(COMBINE.Worksheets("combine"), COMBINE.Worksheets("Combined data-A-completed"), COMBINE.Worksheets("Combined data-A-not completed"))
    Call soort(COMBINE.Worksheets("Combined data-A-not completed"), 6, 14)
    Call combineB(COMBINE.Worksheets("Combined data-A-not completed"), COMBINE.Worksheets("Combined data-B"))
    Call sooort(COMBINE.Worksheets("Combined data-B"), 1, 2, 4)

    RAW.Close (False)
    Call refreshPivotTables
End Sub

Sub clean()
    Call cleanSheet(Workbooks(THIS_FILE).Worksheets("combine"))
    Call cleanSheet(Workbooks(THIS_FILE).Worksheets("misuse-sa"))
    Call cleanSheet(Workbooks(THIS_FILE).Worksheets("misuse-opp"))
    Call cleanSheet(Workbooks(THIS_FILE).Worksheets("Combined data-A-completed"))
    Call cleanSheet(Workbooks(THIS_FILE).Worksheets("Combined data-A-not completed"))
    Call cleanSheet(Workbooks(THIS_FILE).Worksheets("Combined data-B"))
    Call cleanPivotTable(Workbooks(THIS_FILE).Worksheets("Forecast").PivotTables("forecast"))
    Call cleanPivotTable(Workbooks(THIS_FILE).Worksheets("volume-1").PivotTables("volume1a"))
    'Call cleanPivotTable(Workbooks(THIS_FILE).Worksheets("volume-1").PivotTables("volume1b"))
    'Call cleanPivotTable(Workbooks(THIS_FILE).Worksheets("volume-2").PivotTables("volume2a"))
    Call cleanPivotTable(Workbooks(THIS_FILE).Worksheets("volume-2").PivotTables("volume2b"))
    Call cleanPivotTable(Workbooks(THIS_FILE).Worksheets("value").PivotTables("value"))
    Call cleanPivotTable(Workbooks(THIS_FILE).Worksheets("speed").PivotTables("speed"))
End Sub

Sub cleanSheet(sheet As Worksheet)
    For i = sheet.UsedRange.Rows.Count To 2 Step -1
        sheet.Rows(i).Clear
    Next
End Sub

Sub cleanPivotTable(table As PivotTable)
    table.ClearAllFilters
    table.PivotCache.Refresh
End Sub

Sub refreshPivotTables()
    Call refreshPivotTable(Workbooks(THIS_FILE).Worksheets("Forecast").PivotTables("forecast"))
    Call refreshPivotTable(Workbooks(THIS_FILE).Worksheets("volume-1").PivotTables("volume1a"))
    ' Call refreshPivotTable(Workbooks(THIS_FILE).Worksheets("volume-1").PivotTables("volume1b"))
    ' Call refreshPivotTable(Workbooks(THIS_FILE).Worksheets("volume-2").PivotTables("volume2a"))
    Call refreshPivotTable(Workbooks(THIS_FILE).Worksheets("volume-2").PivotTables("volume2b"))
    Call refreshPivotTable(Workbooks(THIS_FILE).Worksheets("value").PivotTables("value"))
    Call refreshPivotTable(Workbooks(THIS_FILE).Worksheets("speed").PivotTables("speed"))
End Sub

Sub refreshPivotTable(p As PivotTable)
    p.RefreshTable
End Sub

Sub combineB(a_not_completed As Worksheet, b As Worksheet)
    b_i = 2
    For a_not_completed_i = 2 To a_not_completed.UsedRange.Rows.Count
'        lead_id = a_not_completed.Cells(a_not_completed_i, 6).Value
'        activity_id = a_not_completed.Cells(a_not_completed_i, 14).Value
'        If lead_id = "" Or activity_id = "" Then
'            a_not_completed.Rows(a_not_completed_i).Copy Destination:=b.Rows(b_i)
'            b_i = b_i + 1
'        Else
'            If lead_id = a_not_completed.Cells(a_not_completed_i - 1, 6).Value Then
'                a_not_completed.Rows(a_not_completed_i).Copy Destination:=b.Rows(b_i - 1)
'            Else
'                a_not_completed.Rows(a_not_completed_i).Copy Destination:=b.Rows(b_i)
'                b_i = b_i + 1
'            End If
'        End If

        a_not_completed.Rows(a_not_completed_i).Copy Destination:=b.Rows(b_i)
        opportunity_id = b.Cells(b_i, 21).Value
        If opportunity_id <> "" And opportunity_id = b.Cells(b_i - 1, 21).Value Then
            b.Range(b.Cells(b_i - 1, 21), b.Cells(b_i - 1, b.UsedRange.Columns.Count)).Clear
        End If
        b_i = b_i + 1
    Next
End Sub

Sub combineA(a As Worksheet, a_completed As Worksheet, a_not_completed As Worksheet)
    a_completed_index = 2
    a_not_completed_index = 2
    
    Dim i As Integer
    i = 2 'start from row 2
    
    While a.Cells(i, "B").Value <> ""
        f = "A" & i & ":" & "AJ" & i
        If a.Cells(i, "V") = "Won" Or a.Cells(i, "V") = "Lost" Or a.Cells(i, "H") = "Lost" Then
            t = "A" & CStr(a_completed_index) & ":" & "AJ" & CStr(a_completed_index)
            a.Range(f).Copy _
                Destination:=a_completed.Range(t)
            a_completed_index = a_completed_index + 1
        ElseIf a.Cells(i, "O") = "Completed" And a.Cells(i, "H") = "" And a.Cells(i, "V") = "" Then
            t = "A" & CStr(a_completed_index) & ":" & "AJ" & CStr(a_completed_index)
            a.Range(f).Copy _
                Destination:=a_completed.Range(t)
            a_completed_index = a_completed_index + 1
        Else
            t = "A" & CStr(a_not_completed_index) & ":" & "AJ" & CStr(a_not_completed_index)
            a.Range(f).Copy _
                Destination:=a_not_completed.Range(t)
            a_not_completed_index = a_not_completed_index + 1
        End If
        i = i + 1
    Wend
End Sub


Sub phaseThree(opp As Worksheet, a As Worksheet)
    Dim misuse As Worksheet
    Set misuse = COMBINE.Worksheets("misuse-opp")
    misuse_index = 2
    
    Dim i As Integer
    i = 2 'start from row 2
    While opp.Cells(i, "J").Value <> ""
        lead_id = opp.Cells(i, "A").Value
        If lead_id = "" Then
            opp.Range("E" & i & ":" & "I" & i).Copy _
                Destination:=a.Range("A" & ALL_CURRENT & ":" & "E" & ALL_CURRENT)
        
            opp.Range("J" & i & ":" & "X" & i).Copy _
                Destination:=a.Range("U" & ALL_CURRENT & ":" & "AI" & ALL_CURRENT)
        
            ALL_CURRENT = ALL_CURRENT + 1
        Else
            With a.Range("F:F")
                Set c = .Find(What:=lead_id, LookAt:=xlWhole)
                If Not c Is Nothing Then 'copy opp to all
                    firstAddress = c.Address
                    Do
                        If a.Cells(c.Row, "U").Value = "" Or opp.Cells(i, "J").Value > a.Cells(c.Row, "U").Value Then
                            opp.Range("J" & i & ":" & "X" & i).Copy _
                                Destination:=a.Range("U" & c.Row & ":" & "AI" & c.Row)
                        End If
                        Set c = .FindNext(c)
                    Loop While Not c Is Nothing And c.Address <> firstAddress
                Else 'append to combine & copy to misuse-opp
                    opp.Range("E" & i & ":" & "I" & i).Copy _
                        Destination:=a.Range("A" & ALL_CURRENT & ":" & "E" & ALL_CURRENT)
        
                    a.Cells(ALL_CURRENT, "F") = opp.Cells(i, "A") 'Lead ID
                    
                    opp.Range("J" & i & ":" & "X" & i).Copy _
                        Destination:=a.Range("U" & ALL_CURRENT & ":" & "AI" & ALL_CURRENT)
        
                    ALL_CURRENT = ALL_CURRENT + 1
        
                    opp.Range("E" & i & ":" & "I" & i).Copy _
                        Destination:=misuse.Range("A" & misuse_index & ":" & "E" & misuse_index)
        
                    misuse.Cells(misuse_index, "F") = opp.Cells(i, "A") 'Lead ID
                    opp.Range("J" & i & ":" & "X" & i).Copy _
                        Destination:=misuse.Range("U" & misuse_index & ":" & "AI" & misuse_index)
                    misuse_index = misuse_index + 1
                End If
            End With
        End If
        i = i + 1
    Wend
End Sub


Sub phaseTwo(lead As Worksheet, a As Worksheet)
    Dim i As Integer
    i = 2
    While lead.Cells(i, "F").Value <> ""
        in_combined = lead.Cells(i, "N")
        If in_combined = "N" Then 'copy to all
            lead.Range("A" & i & ":" & "M" & i).Copy _
                Destination:=a.Range("A" & ALL_CURRENT & ":" & "M" & ALL_CURRENT)
            ALL_CURRENT = ALL_CURRENT + 1
            lead.Cells(i, "N") = "Y"
        End If
        i = i + 1
    Wend
End Sub

Sub phaseOne(sa As Worksheet, lead As Worksheet, a As Worksheet)
    Dim misuse As Worksheet
    Set misuse = COMBINE.Worksheets("misuse-sa")
    
    Dim misuse_index As Integer
    misuse_index = 2
    
    Dim i As Integer
    i = 2 'start from row 2
    While sa.Cells(i, "G").Value <> ""
        lead_id = sa.Cells(i, "F").Value
        If lead_id = "" Then
            'copy to all
            sa.Range("A" & i & ":" & "E" & i).Copy _
                Destination:=a.Range("A" & ALL_CURRENT & ":" & "E" & ALL_CURRENT)
            a.Cells(ALL_CURRENT, "F") = "" 'lead ID
            sa.Range("G" & i & ":" & "M" & i).Copy _
                Destination:=a.Range("N" & ALL_CURRENT & ":" & "T" & ALL_CURRENT)
            a.Cells(ALL_CURRENT, "AJ") = sa.Cells(i, "N")
            ALL_CURRENT = ALL_CURRENT + 1
        Else
            With lead.Range("F:F")
                Set c = .Find(What:=lead_id, LookAt:=xlWhole)
                If Not c Is Nothing Then 'lead record found
                        If lead.Cells(c.Row, "E").Value = sa.Cells(i, "E").Value Then
                            sa.Range("A" & i & ":" & "E" & i).Copy _
                            Destination:=a.Range("A" & ALL_CURRENT & ":" & "E" & ALL_CURRENT)
                            lead.Range("F" & c.Row & ":" & "M" & c.Row).Copy _
                                Destination:=a.Range("F" & ALL_CURRENT & ":" & "M" & ALL_CURRENT)
                            lead.Cells(c.Row, "N").Value = "Y"
                            sa.Range("G" & i & ":" & "M" & i).Copy _
                                Destination:=a.Range("N" & ALL_CURRENT & ":" & "T" & ALL_CURRENT)
                            a.Cells(ALL_CURRENT, "AJ") = sa.Cells(i, "N")
                            ALL_CURRENT = ALL_CURRENT + 1
                        Else 'account name not matched, copy to misuse-sa
                            sa.Range("A" & i & ":" & "F" & i).Copy _
                                Destination:=misuse.Range("A" & misuse_index & ":" & "F" & misuse_index)
                            sa.Range("G" & i & ":" & "M" & i).Copy _
                                Destination:=misuse.Range("N" & misuse_index & ":" & "T" & misuse_index)
                            misuse.Cells(misuse_index, "AJ") = sa.Cells(i, "N")
                            misuse_index = misuse_index + 1
                        End If
                Else 'lead record not found, copy to misuse-sa
                    'copy to misuse too
                    sa.Range("A" & i & ":" & "F" & i).Copy _
                        Destination:=misuse.Range("A" & misuse_index & ":" & "F" & misuse_index)
                    sa.Range("G" & i & ":" & "M" & i).Copy _
                        Destination:=misuse.Range("N" & misuse_index & ":" & "T" & misuse_index)
                    misuse.Cells(misuse_index, "AJ") = sa.Cells(i, "N")
                    misuse_index = misuse_index + 1
                End If
            End With
        End If
        i = i + 1
    Wend
End Sub

