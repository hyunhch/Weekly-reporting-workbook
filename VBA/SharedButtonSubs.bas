Attribute VB_Name = "SharedButtonSubs"
Option Explicit

Sub RemoveSelectedButton()

    Dim DelSheet As Worksheet
    Dim SearchRange As Range
    Dim CheckRange As Range
    Dim SortRange As Range
    Dim DelRange As Range
    Dim c As Range
    Dim d As Range
    Dim DelTable As ListObject
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    
    Set DelSheet = ActiveSheet
    
    'Make sure there's a table with at least one student checked
    If CheckTable(DelSheet) <> 1 Then
        GoTo Footer
    End If
    
    Set DelTable = DelSheet.ListObjects(1)
    Set SearchRange = DelTable.DataBodyRange
    Set SortRange = DelTable.ListColumns("Select").DataBodyRange
    
    'On the Report Page, there should always be a Totals row under the headers, but that row can't be checked
    'We don't want to sort the Totals row
    If DelSheet.Name = "Report Page" Then
        Set c = SearchRange.Resize(SearchRange.Rows.Count - 1, SearchRange.Columns.Count)
        Set SearchRange = c.Offset(1, 0)

        Set c = SortRange.Resize(SortRange.Rows.Count - 1, 1)
        Set SortRange = c.Offset(1, 0)
    End If
    
    'Find the checked rows
    Set DelRange = FindChecks(SortRange)
    
    'Different procedures for different sheets
    If DelSheet.Name = "Roster Page" Then
        Call RemoveFromRoster(DelRange.Offset(0, 1))
    ElseIf DelSheet.Range("A1").Value = "Practice" Then
        Call RemoveFromActivity(DelSheet, DelRange.Offset(0, 1))
    Else
        Call RemoveRows(DelSheet, SearchRange, SortRange, DelRange)
    End If
    
Footer:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True

End Sub

Sub SelectAllButton()
'Assigns value in Select column to "a" or ""
'The report sheet doesn't have a table, consider changing that in the future

    Dim CheckSheet As Worksheet
    Dim CheckRange As Range
    Dim CheckRows As Long
    Dim CheckTable As ListObject
    
    Set CheckSheet = ActiveSheet
    
    Call UnprotectSheet(CheckSheet)
    
    'If there isn't a table
    If CheckSheet.ListObjects.Count = 0 Then
        GoTo Footer
    End If
    
    Set CheckTable = CheckSheet.ListObjects(1)
    Set CheckRange = CheckTable.ListColumns("Select").DataBodyRange
    
    'If there are no rows in the data body range
    If CheckRange Is Nothing Then
        GoTo Footer
    End If
    
    CheckRows = CheckRange.Rows.Count
    
    'Exclude first row on Report sheet
    If CheckSheet.Name = "Report Page" Then
        If CheckRange.Rows.Count < 2 Then
            GoTo Footer
        Else
            Set CheckRange = CheckRange.Offset(1, 0).Resize(CheckRange.Rows.Count - 1, 1)
            CheckRows = CheckRows - 1
        End If
    End If
    
    'Check all if any are blank, uncheck all if none are blank
    'Only apply to visible cells
    CheckRange.Font.Name = "Marlett"
    
    If Application.CountIf(CheckRange, "a") = CheckRows Then
        CheckRange.SpecialCells(xlCellTypeVisible).Value = ""
    Else
        CheckRange.SpecialCells(xlCellTypeVisible).Value = "a"
    End If

Footer:
    Call ResetProtection

End Sub


