Attribute VB_Name = "RosterButtonSubs"
Option Explicit

Sub ReadRosterButton()
'Read in the roster, make a sortable/filterable table, add Marlett boxes, conditional formatting

    Dim RosterSheet As Worksheet
    Dim RefSheet As Worksheet
    Dim RosterTableStart As Range
    Dim ColNames() As Variant
    Dim NumStudents As Long
    Dim i As Long
    
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    Set RefSheet = Worksheets("Ref Tables")
    Set RosterSheet = Worksheets("Roster Page")
    Set RosterTableStart = RosterSheet.Range("A6")
    
    'Unprotect
    Call UnprotectCheck(RosterSheet)
    
    'The column headers need to remain unprotected to allow sorting
    'However, the column names and order need to remain constant
    ColNames = Application.Transpose(ActiveWorkbook.Names("ColumnNamesList").RefersToRange.Value)
    Call ResetColumns(RosterSheet, RosterTableStart, ColNames)
    
    'Remove any formatting in the header
    If RosterSheet.AutoFilterMode = True Then
        RosterSheet.AutoFilterMode = False
    End If
    
    'Delete any table objects and remove formatting
    Dim OldTable As ListObject
    
    For Each OldTable In RosterSheet.ListObjects
        OldTable.Unlist
    Next OldTable
    
    RosterSheet.Cells.FormatConditions.Delete
    RosterSheet.Cells.Borders.LineStyle = Excel.XlLineStyle.xlLineStyleNone
    
    'Make sure there are some students added
    NumStudents = CheckTableLength(RosterSheet, RosterTableStart.Offset(0, 1))
    If Not NumStudents > 0 Then
        MsgBox ("Please add at least one student.")
        GoTo Footer
    End If
    
    'Remove any empty rows
    For i = RosterTableStart.Row + NumStudents To RosterTableStart.Row + 1 Step -1
        If Not Len(RosterSheet.Cells(i, 2).Value) > 0 Then
            RosterSheet.Cells(i, 2).EntireRow.Delete
        End If
    Next i
    
    'Make a table called "RosterTable" and format
    Call TableCreate(RosterSheet, RosterTableStart, "RosterTable")
    
    'If a column header was deleted, it's renamed "column" and a number.Delete all columns that contain "column"
    Dim c As Range
    
    For i = RosterSheet.ListObjects(1).ListColumns.Count To 1 Step -1
        Set c = RosterSheet.Cells(RosterTableStart.Row, i)
        If c.Value Like "Column*" Then
            c.EntireColumn.Delete
        End If
    Next i
    
    'Add Marlett Boxes
    Dim BoxRange As Range
    
    Set BoxRange = RosterSheet.ListObjects("RosterTable").ListColumns("Select").DataBodyRange
    Call AddMarlettBox(BoxRange, RosterSheet)
    
    'Push names to the Records sheet
    Call PushRosterNames

    'Push tabulated totals to the Report sheet
    Call PullReportTotals

Footer:
    'Reprotect
    Call ResetProtection
    
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

End Sub

Sub ClearRosterButton()
'Delete everything and reset to default columns

    Dim RosterSheet As Worksheet
    Dim RosterTableStart As Range
    Dim ColNames() As Variant

    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    Set RosterSheet = Worksheets("Roster Page")
    Set RosterTableStart = RosterSheet.Range("A6")
    
    'Unprotect
    Call UnprotectCheck(RosterSheet)
    
    'Everything starting at RosterTableStart will be deleted
    Call ClearSheet(RosterTableStart, 0, RosterSheet)
    
    'Put the default columns back in
    ColNames = Application.Transpose(ActiveWorkbook.Names("ColumnNamesList").RefersToRange.Value)
    Call ResetColumns(RosterSheet, RosterTableStart, ColNames)
    
    'Ask if the records sheet should be cleared as well
    Dim DelConfirm As Long
    
    DelConfirm = MsgBox("Would you like to delete all recorded activities and attendance?" & vbCr & "This cannot be undone.", vbQuestion + vbYesNo + vbDefaultButton2)
    If DelConfirm = vbYes Then
        Call ClearRecords
    End If
    
    'Clear the Report sheet, depending on response
    If DelConfirm = vbYes Then
        Call ClearReportButton(1)
    Else
        Call ClearReportTotals
    End If
    
Footer:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
    'Reprotect
    Call ResetProtection

End Sub

Sub OpenNewActivityButton()

    Dim RosterSheet As Worksheet
    Dim CoverSheet As Worksheet
    Dim RosterTableStart As Range
    Dim CoverInfoRange As Range
    Dim SearchRange As Range
    Dim c As Range

    Set RosterSheet = Worksheets("Roster Page")
    Set RosterTableStart = RosterSheet.Range("A:A").Find("Select", , xlValues, xlWhole)
    
    'Make sure the roster has been parsed
    If RosterSheet.ListObjects.Count < 1 Then
        MsgBox ("Please parse the roster first")
        GoTo Footer
    End If

    'Make sure there are students in the roster
    If Not CheckTableLength(RosterSheet, RosterTableStart) > 0 Then
        MsgBox ("You don't have any students on this page.")
        GoTo Footer
    End If

    'Make sure at least one student is selected
    Set SearchRange = RosterSheet.ListObjects(1).ListColumns("Select").DataBodyRange
    
    If FindChecks(SearchRange) Is Nothing Then
        MsgBox ("Please select at least one student")
        GoTo Footer
    End If

    'Make sure the information on the Cover sheet is added
    Set CoverSheet = Worksheets("Cover Page")
    Set CoverInfoRange = CoverSheet.Range("B3:B5")
    
    For Each c In CoverInfoRange
        If Not Len(Trim(c.Value)) > 0 Then
            MsgBox ("Please fill out name, date, and center on the Cover Page.")
            GoTo Footer
        End If
    Next c

    NewActivityForm.Show
    
Footer:

End Sub

Sub OpenLoadActivityButton()

    Dim RecordsSheet As Worksheet
    Dim FCol As Long
    Dim LCol As Long
    
    Set RecordsSheet = Worksheets("Records Page")

    'Make sure there's something to load
    FCol = RecordsSheet.Range("1:1").Find("V BREAK", , xlValues, xlWhole).Column
    LCol = RecordsSheet.Range("1:1").Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
       
    If FCol = LCol Then
        MsgBox ("You don't have any saved activities")
        GoTo Footer
    End If

    LoadActivityForm.Show

Footer:

End Sub

Sub AddSelectedStudentsButton()
'Adds checked students on the roster sheet to an activity sheet
'Skip ones already present

    Dim RosterSheet As Worksheet
    Dim CheckRange As Range
    Dim AddNames As Range
    Dim RosterTableStart As Range
    
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    Set RosterSheet = Worksheets("Roster Page")
    Set RosterTableStart = RosterSheet.Range("A:A").Find("Select", , xlValues, xlWhole)
    
    'This will only happen if the columns have been changed
    If RosterTableStart Is Nothing Then
        MsgBox ("Something has gone wrong. Please parse the roster and try again.")
        GoTo Footer
    End If
    
    'Make sure the Roster has students and is parsed
    If Not CheckTableLength(RosterSheet, RosterTableStart) > 0 Then
        MsgBox ("You have no students in your roster.")
        GoTo Footer
    End If
    
    If RosterSheet.ListObjects.Count < 1 Then
        Call ReadRosterButton
    End If
    
    'Define the range of selected students
    Set CheckRange = RosterSheet.ListObjects("RosterTable").ListColumns("Select").DataBodyRange
    
    'Make sure at least one student is checked
    If FindChecks(CheckRange) Is Nothing Then
        MsgBox ("Please select at least one student.")
        GoTo Footer
    Else
        Set AddNames = FindChecks(CheckRange).Offset(0, 1)
    End If
    
    'Open the userform
    AddStudentsForm.Show
    
Footer:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
End Sub
