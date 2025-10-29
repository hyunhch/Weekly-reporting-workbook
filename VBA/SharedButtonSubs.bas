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

    Dim CheckSheet As Worksheet
    Dim CheckRange As Range
    Dim c As Range
    Dim i As Long
    Dim CheckRows As Long
    Dim TargetTable As ListObject
    
    Set CheckSheet = ActiveSheet
    
    Call UnprotectSheet(CheckSheet)
    
    'If Check for a table with rows
    If CheckSheet.Name <> "Report Page" Then
        i = CheckReport(CheckSheet)
    Else
        i = CheckTable(CheckSheet)
    End If
    
    If i > 2 Then
        GoTo Footer
    End If
        
    'Define range to search
    Set TargetTable = CheckSheet.ListObjects(1)
    Set c = TargetTable.ListColumns("Select").DataBodyRange
    Set CheckRange = c.SpecialCells(xlCellTypeVisible)
        If CheckRange Is Nothing Then 'No visible cells
            GoTo Footer
        End If
    
    CheckRows = CheckRange.Rows.Count
    c.Font.Name = "Marlett"
    
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
    For Each c In CheckRange
        If c.Value <> "a" Then
            CheckRange.Value = "a"
            
            GoTo Footer
        End If
    Next c
    
    CheckRange.Value = ""

Footer:
    Call ResetProtection

    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

End Sub


