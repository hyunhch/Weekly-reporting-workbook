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
    Dim i As Long
    Dim DelTable As ListObject
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    
    Set DelSheet = ActiveSheet
    
    'Make sure there's a table with at least one row checked
    If DelSheet.Name = "Report Page" Then
        i = CheckReport(DelSheet)
    Else
        i = CheckTable(DelSheet)
    End If
    
    If i <> 1 Then
        GoTo Footer
    End If
    
    Set DelTable = DelSheet.ListObjects(1)
    Set SearchRange = DelTable.DataBodyRange
    Set SortRange = DelTable.ListColumns("Select").DataBodyRange
    
    'Find the checked rows
    Set DelRange = FindChecks(SortRange)

    'Different procedures on the Roster, Activity, and Report pages
    Select Case DelSheet.Name
        Case "Roster Page"
            Call RemoveFromRoster(DelSheet, DelRange, "Prompt")
        
        Case "Report Page"
            Call RemoveFromReport(DelRange)
            
        Case Else
            Call RemoveFromActivity(DelSheet, DelRange)

    End Select
    
Footer:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True

End Sub

Sub SelectAllButton()
'Assigns value in Select column to "a" or ""
'The report sheet doesn't have a table, consider changing that in the future

    Dim SelectSheet As Worksheet
    Dim SelectRange As Range
    Dim c As Range
    Dim i As Long
    Dim CheckRows As Long
    Dim SelectTable As ListObject
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    
    Set SelectSheet = ActiveSheet
    
    'If Check for a table with rows
    If SelectSheet.Name = "Report Page" Then
        i = CheckReport(SelectSheet)
    Else
        i = CheckTable(SelectSheet)
    End If
    
    If i > 2 Then
        GoTo Footer
    End If
        
    'Define range to search
    Set SelectTable = SelectSheet.ListObjects(1)
    Set SelectRange = SelectTable.ListColumns("Select").DataBodyRange

    'Exclude first row on Report sheet
    If SelectSheet.Name = "Report Page" Then
        If Not SelectRange.Rows.Count > 1 Then
            GoTo Footer
        End If
        
        Set c = SelectRange
        Set SelectRange = c.Offset(1, 0).Resize(c.Rows.Count - 1, 1)
    End If
    
    'Check all if any are blank, uncheck all if none are blank
    Call UnprotectSheet(SelectSheet)
    
    With SelectRange
        .Font.Name = "Marlett"
        i = .Cells.Count
        
        'Only apply to visible cells
        If Application.CountIf(SelectRange, "a") = i Then
            .SpecialCells(xlCellTypeVisible).ClearContents
        Else
            .SpecialCells(xlCellTypeVisible).Value = "a"
        End If
    End With

Footer:
    Call ResetProtection

    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

End Sub

