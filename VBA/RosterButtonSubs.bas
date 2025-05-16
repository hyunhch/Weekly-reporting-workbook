Attribute VB_Name = "RosterButtonSubs"
Option Explicit

Sub OpenAddStudentsButton()
'Adds selected students to a saved activity

    Dim RosterSheet As Worksheet
    Dim RecordsSheet As Worksheet
    Dim RosterNameRange As Range
    Dim RosterCheckRane As Range
    Dim AddCheck As Long
    Dim RosterTable As ListObject
    
    Set RosterSheet = Worksheets("Roster Page")
    Set RecordsSheet = Worksheets("Records Page")
    
    'Make sure there are any saved activities
    AddCheck = CheckRecords(RecordsSheet)
    If AddCheck > 2 Then
        MsgBox ("You have no saved activities.")
        GoTo Footer
    End If

    'Make sure there is a roster table, that there's at least one student, and that there's at least one checked
    If RosterSheet.ListObjects.Count = 0 Then
        GoTo Footer
    End If
    
    Set RosterTable = RosterSheet.ListObjects(1)
    If Not RosterTable.ListRows.Count > 0 Then
        GoTo Footer
    End If

    Set RosterNameRange = RosterTable.ListColumns("First").DataBodyRange
    If FindChecks(RosterNameRange.Offset(0, -1)) Is Nothing Then
        GoTo Footer
    End If

    'Show form
    AddStudentsForm.Show

Footer:

End Sub

Sub OpenLoadActivityButton()
'Checks to see if there are any saved activities and opens the load activity form

    Dim RecordsSheet As Worksheet
    Dim LoadCheck As Long
    
    Set RecordsSheet = Worksheets("Records Page")
    
    'Check to make sure there are any saved activities. Activities with no students are fine.
    LoadCheck = CheckRecords(RecordsSheet)
    If LoadCheck > 2 Then
        MsgBox ("You have no saved activities")
        GoTo Footer
    End If
    
    'Show form
    LoadActivityForm.Show
    
Footer:

End Sub

Sub OpenNewActivityButton()
'Opens form to create a new activity. Does not require any selected students

    Dim RosterSheet As Worksheet
    Dim RosterNameRange As Range
    Dim RosterCheckRange As Range
    Dim RosterTable As ListObject
    
    Set RosterSheet = Worksheets("Roster Page")
    
    'Make sure there's a parsed table with at least one student
    If CheckTable(RosterSheet) > 1 Then
        GoTo Footer
    End If
    
    'Make sure the CoverSheet is filled out
    If CheckCover <> 1 Then
        MsgBox ("Please fill out your name, the date, and your center on the Cover Page")
        
        GoTo Footer
    End If
    
    'Show form
    NewActivityForm.Show
    
Footer:

End Sub

Sub RosterClearButton()
'Delete everything, reset columns, clear records
    
    Dim OldBook As Workbook
    Dim RosterSheet As Worksheet
    Dim RecordsSheet As Worksheet
    Dim ActivitySheet As Worksheet
    Dim RecordsLabelRange As Range
    Dim NameRange As Range
    Dim DelCell As Range
    Dim ColNames() As Variant
    Dim RosterTable As ListObject
    
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    Set OldBook = ThisWorkbook
    Set RosterSheet = Worksheets("Roster Page")
    Set RecordsSheet = Worksheets("Records Page")
    
    'Skip if there's no table
    If RosterSheet.ListObjects.Count = 0 Then
        GoTo Footer
    End If
    
    'Skip if it's an empty table
    Set RosterTable = RosterSheet.ListObjects(1)
    Set NameRange = RosterTable.ListColumns("First").DataBodyRange
    
    If NameRange Is Nothing Then
        GoTo Footer
    End If

    'Pass to the RemoveSelected sub, which already handles takings things off the Records sheet and exporting
    Call UnprotectSheet(RosterSheet)
    NameRange.Offset(0, -1).Value = "a"
    If RemoveFromRoster(NameRange) = 0 Then
        GoTo Footer
    End If
    
    'Clear everything on the Records and Report sheets
    Call ClearRecords
    Call ClearReportButton
    
    'Turn everything back off
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    'Delete the roster table and clear formatting
    Set DelCell = RosterSheet.Range("A6") 'Harded coded in cases where the table has been unlisted, the column names changed, etc.
    Call ClearSheet(RosterSheet, 1, DelCell)
    
    'Reset columns
    ColNames = Application.Transpose(ActiveWorkbook.Names("ColumnNamesList").RefersToRange.Value)
    Call ResetTableHeaders(RosterSheet, DelCell, ColNames)
    
    'Remove green around Parse button
    RosterSheet.Range("A1:C3").Interior.Pattern = xlNone
    
    'Delete any activity sheets
    For Each ActivitySheet In ActiveWorkbook.Sheets
        If ActivitySheet.Range("A1") = "Practice" Then
            ActivitySheet.Delete
        End If
    Next ActivitySheet
    
Footer:
    Call ResetProtection

    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

End Sub

Sub RosterParseButton()
'Read in the roster, table with conditional formatting, Marlett boxes, push to the ReportSheet

    Dim RosterSheet As Worksheet
    Dim RecordsSheet As Worksheet
    Dim RosterNameRange As Range
    Dim RosterTableRange As Range
    Dim RecordsNameRange As Range
    Dim DelRange As Range
    Dim BoxRange As Range
    Dim c As Range
    Dim d As Range
    Dim NumDuplicate As Long
    Dim ColNames() As Variant
    Dim RosterTable As ListObject
    
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    Set RosterSheet = Worksheets("Roster Page")
    Set RecordsSheet = Worksheets("Records Page")
    Set c = RosterSheet.Range("A6")
    
    Call UnprotectSheet(RosterSheet)
    
    'Find the range for the new table, break if there is nothing but the header
    Set RosterTableRange = FindTableRange(RosterSheet)
    
    If Not RosterTableRange.Rows.Count > 1 Then
        'Reset the headers
        ColNames = Application.Transpose(ActiveWorkbook.Names("ColumnNamesList").RefersToRange.Value)
        Call ResetTableHeaders(RosterSheet, c, ColNames)
        GoTo Footer
    End If

    'Make sure there are no filters on the table
    If RosterSheet.AutoFilterMode = True Then
        RosterSheet.AutoFilterMode = False
    End If
    
    'Remove the existing table and formatting
    Call RemoveTable(RosterSheet)
    
    'Headers remain unlocked for sorting and filtering. Headers recorded in a list
    ColNames = Application.Transpose(ActiveWorkbook.Names("ColumnNamesList").RefersToRange.Value)
    Call ResetTableHeaders(RosterSheet, c, ColNames) 'This preserves any extra columns added
    
    'Make a table called "RosterTable" so that any added students are included and any blank rows are removed
    Set RosterTable = CreateTable(RosterSheet, "RosterTable", RosterTableRange)
    
    'Identify and remove duplicate students and blank rows
    Set RosterNameRange = RosterTable.ListColumns("First").DataBodyRange
    Set c = FindDuplicate(RosterNameRange)
    Set d = FindBlanks(RosterSheet, RosterNameRange)

    If Not c Is Nothing Then
        NumDuplicate = c.Cells.Count
        Set DelRange = c
    End If
    
    If Not d Is Nothing Then
        If DelRange Is Nothing Then
            Set DelRange = d
        Else
            Set DelRange = Union(DelRange, d)
        End If
    End If

    If Not DelRange Is Nothing Then
        Call RemoveRows(RosterSheet, RosterTable.DataBodyRange, RosterNameRange, DelRange)
        Set RosterTable = RosterSheet.ListObjects(1)
    End If
    
    'Add Marlett boxes
    Set BoxRange = RosterTable.ListColumns("Select").DataBodyRange
    Call AddMarlettBox(BoxRange)
    
    'Format if it hasn't been done already
    Set c = FindTableHeader(RosterSheet, "Gender").Offset(1, 0) 'Could be ethnicity, grade, major, etc. instead
    
    If c.FormatConditions.Count < 2 Then 'Flag when blank, flag if there's a bad value
        Call FormatTable(RosterSheet, RosterTable) 'This is one of the bottlenecks for speed in my code, so only run it if we need to
    End If
    
    'Push names to records and report sheets
    Call CopyToRecords(RecordsSheet, RosterSheet, RosterNameRange)
    Call TabulateReportTotals
    
    'See if the RosterSheet and ReportSheet have the same number of students. This shouldn't happen
    Set RosterNameRange = RosterTable.ListColumns("First").DataBodyRange
    Set RecordsNameRange = FindRecordsName(RecordsSheet)
        
    If RosterNameRange.Rows.Count = RecordsNameRange.Rows.Count Then
        RosterSheet.Range("A1:C3").Interior.Pattern = xlNone
    Else
        'Set c = FindUnique(RecordsNameRange, RosterNameRange) 'Need to account for if the Roster list is short
        RosterSheet.Range("A1:C3").Interior.ColorIndex = 43 'May want to change this to checkling the two list of names again
    End If
    
Footer:
    'Reprotect
    Call ResetProtection
    
    If NumDuplicate > 0 Then
        MsgBox (NumDuplicate & " duplicates removed.")
    End If
    
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
End Sub





