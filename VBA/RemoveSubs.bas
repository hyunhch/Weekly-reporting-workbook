Attribute VB_Name = "RemoveSubs"
Option Explicit

Sub RemoveFromActivity(ActivitySheet As Worksheet, DelRange As Range)
'Called whenever students need to be removed from an activity sheet
'Remove students, save activity, retabulate
'DelRange can be from RosterSheet, the RecordsSheet, or the same ActivitySheet

    Dim RecordsSheet As Worksheet
    Dim ActivityNameRange As Range
    Dim ActivityDelRange As Range
    Dim LabelCell As Range
    Dim ActivityTable As ListObject
    
    'Make sure there's a table with students
    If CheckTable(ActivitySheet) > 2 Then
        GoTo Footer
    End If
    
    Set ActivityTable = ActivitySheet.ListObjects(1)
    Set ActivityNameRange = ActivityTable.ListColumns("First").DataBodyRange
    
    'If the range of names is already on the sheet, skip name matching
    If DelRange.Worksheet.Name = ActivitySheet.Name Then
        Set ActivityDelRange = DelRange
    Else
        Set ActivityDelRange = FindName(DelRange, ActivityNameRange)
    End If
    
    Call UnprotectSheet(ActivitySheet)
    
    'No matches fouund
    If ActivityDelRange Is Nothing Then
        GoTo Footer
    End If
    
    'Remove students and save
    Set RecordsSheet = Worksheets("Records Page")
    Set LabelCell = FindActivityLabel(ActivitySheet)
    
    If LabelCell Is Nothing Then 'This shouldn't happen
        MsgBox ("Something has gone wrong. Please delete this sheet and remake it.")
        GoTo Footer
    End If
    
    Call RemoveRows(ActivitySheet, ActivityTable.DataBodyRange, ActivityNameRange, ActivityDelRange)
    Call RecordsPullAttendance(ActivitySheet, ActivityNameRange, LabelCell)
    Call SaveActivity(ActivitySheet, RecordsSheet, LabelCell)
    
Footer:

End Sub

Sub RemoveFromRecords(RecordsSheet As Worksheet, RosterNameRange As Range, RosterDelRange As Range)
'Called when students are removed from the Roster Page
'Removes students from Records Sheet and any open activity sheets
'Retabulates all activities on the Report Page

    Dim ReportSheet As Worksheet
    Dim RecordsNameRange As Range
    Dim RecordsDelRange As Range
    Dim RecordsLabelRange As Range
    Dim RecordsFullNameRange As Range
    Dim ActivityNameRange As Range
    Dim ActivityDelRange As Range
    Dim i As Long
    Dim ReportTable As ListObject
    
    Set ReportSheet = Worksheets("Report Page")
    
    'Check if there are students or activities on the Records Sheet
    i = CheckRecords(RecordsSheet)
    
    If i = 2 Or i = 4 Then 'No students, nothing to delete
        GoTo Footer
    End If
    
    Call UnprotectSheet(RecordsSheet)
    
    'Define ranges to search and delete on the RecordsSheet
    Set RecordsNameRange = FindRecordsName(RecordsSheet)
    Set RecordsFullNameRange = RecordsNameRange.Resize(RecordsNameRange.Rows.Count, 2)
    Set RecordsDelRange = FindName(RosterDelRange, RecordsNameRange)
    
    If RecordsDelRange Is Nothing Then 'This shouldn't be able to happen
        GoTo Footer
    End If

    Call RemoveRows(RecordsSheet, RecordsFullNameRange.EntireRow, RecordsNameRange, RecordsDelRange)

Footer:

End Sub

Sub RemoveFromReport(DelRange As Range)
'Removes one or more activities from the ReportSheet
'DelRange should be cells containing labels, so this is different from the button to remove rows that looks for checks

    Dim ReportSheet As Worksheet
    Dim ReportDelRange As Range
    Dim ReportLabelRange As Range
    Dim c As Range
    Dim d As Range
    Dim ReportTable As ListObject
    
    Set ReportSheet = Worksheets("Report Page")
    Set ReportTable = ReportSheet.ListObjects(1)
    
    'Check that there's a table with more than 2 rows
    If CheckTable(ReportSheet) > 2 Then
        GoTo Footer
    ElseIf ReportTable.ListRows < 2 Then
        GoTo Footer
    End If
    
    Call UnprotectSheet(ReportSheet)
    
    'Make a range to remove
    For Each c In DelRange
        Set d = FindReportLabel(ReportSheet, c)
        
        If Not d Is Nothing Then
            Set ReportDelRange = BuildRange(c, ReportDelRange)
        End If
    Next c
    
    If ReportDelRange Is Nothing Then
        GoTo Footer
    End If
    
    'Pass to remove
    Set ReportLabelRange = FindReportLabel(ReportSheet)
    
    Call RemoveRows(ReportSheet, ReportTable.DataBodyRange, ReportLabelRange, ReportDelRange)

Footer:

End Sub

Function RemoveFromRoster(DelRange As Range) As Long
'Prompts confirm deletion and export Attendance, removes from records and open activity sheets, retabulates
'Exports based on passed long
'Passing "Clear" Will wipe everything instead of deleting some rows
'Returns 1 for removing
'Returns 2 for removing and exporting
'Returns 0 for aborting

    Dim OldBook As Workbook
    Dim NewBook As Workbook
    Dim RosterSheet As Worksheet
    Dim RecordsSheet As Worksheet
    Dim ReportSheet As Worksheet
    Dim ActivitySheet As Worksheet
    Dim RosterNameRange As Range
    Dim RecordsNameRange As Range
    Dim RecordsDelRange As Range
    Dim i As Long
    Dim DelConfirm As Long
    Dim ExportConfirm As Long
    Dim RosterTable As ListObject
    Dim DeleteAll As Boolean
    
    RemoveFromRoster = 0
    
    Set RosterSheet = Worksheets("Roster Page")
    Set RecordsSheet = Worksheets("Records Page")
    Set ReportSheet = Worksheets("Report Page")
    
    'Make sure there's a table, that there's at least one student, and at least one checked student
    If CheckTable(RosterSheet) > 1 Then
        GoTo Footer
    End If
    
    'Confirm Deletion
    DelConfirm = MsgBox("This will remove the selected students from any recorded activities. Do you wish to continue?", vbQuestion + vbYesNo + vbDefaultButton2)
    If DelConfirm <> vbYes Then
        GoTo Footer
    End If
    
    Set RosterTable = RosterSheet.ListObjects(1)
    Set RosterNameRange = RosterTable.ListColumns("First").DataBodyRange 'One column to the right of the checks

    'If the entire column is being deleted
    If DelRange.Cells.Count = RosterNameRange.Cells.Count Then
        DeleteAll = True
    Else
        DeleteAll = False
    End If
        
    'Define the range of students to remove. Skip forward if there are no matches
    Set RecordsNameRange = FindRecordsName(RecordsSheet)
    Set RecordsDelRange = FindName(RecordsNameRange, DelRange)
    
    If RecordsDelRange Is Nothing Then
        GoTo SkipRecords
    End If
        
    'Confirm exporting student Attendance. Only display if there are both students and activities on the RecordsSheet
    i = CheckRecords(RecordsSheet)
    
    If i = 2 Or i = 4 Then 'No students on RecordsSheet
        GoTo SkipRecords
    ElseIf i = 1 Then 'Students and activities on RecordsSheet
        ExportConfirm = MsgBox("Do you wish to save a copy of these students' attendance before removing them?", vbQuestion + vbYesNo + vbDefaultButton2)
    End If

    'Export identified students
    If ExportConfirm <> vbYes Then
        GoTo SkipExport
    End If

    Set OldBook = ActiveWorkbook
    Set NewBook = MakeNewBook(RecordsSheet, ReportSheet, RosterSheet, DelRange)

    Call SaveNewBook(OldBook, NewBook)
    OldBook.Activate
    
    RemoveFromRoster = RemoveFromRoster + 1
    
SkipExport:
    'Delete students from records and retabulate
    Call RemoveFromRecords(RecordsSheet, RecordsNameRange, RecordsDelRange)
    
SkipRecords:
    'If we are clearing the roster entirely
    If DeleteAll = True Then
        'Clear report and records sheets
        Call ClearRecords 'This can be run without any students
        Call ClearReport
        
        'Delete any open activity sheet
        For Each ActivitySheet In ActiveWorkbook.Sheets
            If ActivitySheet.Range("A1").Value = "Practice" Then
                ActivitySheet.Delete
            End If
        Next ActivitySheet
        
        'Delete content and formats starting below the header
        Call ClearSheet(RosterSheet, "No", RosterTable.HeaderRowRange(1, 1).Offset(1, 0))
    Else
        'Otherwise, look through each sheet and remove matching students
        For Each ActivitySheet In ThisWorkbook.Sheets
            If ActivitySheet.Range("A1").Value = "Practice" Then
                Call RemoveFromActivity(ActivitySheet, DelRange.Offset) 'Includes saving the activity
            End If
        Next ActivitySheet
        
        'Delete students from the RosterSheet and parse
        Call RemoveRows(RosterSheet, RosterTable.DataBodyRange, RosterNameRange, DelRange)
        Call RosterParseButton
        Call RetabulateReport
        
    End If
        
    RemoveFromRoster = RemoveFromRoster + 1
        
Footer:

End Function

Sub RemoveRows(TargetSheet As Worksheet, SearchRange As Range, SortRange As Range, DelRange As Range)
'Sorts SearchRange and deletes everything in DelRange
'Needs to be passed the full range to sort, i.e. a table DataBodyRange

    'Dim SortRange As Range
    Dim SortDelRange As Range
    Dim c As Range
    Dim d As Range
    Dim i As Long
    Dim TargetTable As ListObject
    Dim HasTable As Boolean
    
    Call UnprotectSheet(TargetSheet)

    'I don't think this is needed since I'm defining a number of cells to be deleted rather than the entire row. Need to test
    'Remove any table and formatting
    If TargetSheet.ListObjects.Count > 0 Then
        HasTable = True
        Call RemoveTable(TargetSheet)
    End If
    
    SearchRange.FormatConditions.Delete
    
    'Flag each row to be deleted
    DelRange.Interior.Color = vbRed
    
    'Sort by color
    With TargetSheet.Sort
        .SortFields.Clear
        .SortFields.Add2(SortRange, xlSortOnCellColor, xlAscending, , xlSortNormal).SortOnValue.Color = RGB(255, 0, 0)
        .SetRange SearchRange
        .Header = xlGuess
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    'Find the bounds of the red cells
    'Not looking at contents because the sub can be called to delete any row
    Set c = SearchRange.Rows(1)
    
    'For Each d In SortRange.Cells 'This is giving me the wrong row. I'm not sure why
        'If d.Interior.Color <> vbRed Then
            'Set d = SearchRange.Rows(d.Row - 1)
            'Exit For
        'End If
    'Next d
    
    For i = c.Row To SearchRange.Rows(SearchRange.Rows.Count + 1).Row 'In case every row is checked
        Set d = TargetSheet.Cells(i, SortRange.Column)
        If d.Interior.Color <> vbRed Then
            Set d = d.Offset(-1, 0)
            Exit For
        End If
    Next i
    
    'Make a range and delete
    Set SortDelRange = TargetSheet.Range(c, d)
    SortDelRange.Delete Shift:=xlUp
    
    'Put the table back in, if applicable
    If HasTable = False Then
        GoTo Footer
    End If
    
    If TargetSheet.Name = "Report Page" Then
        Set TargetTable = CreateReportTable
        Call FormatReportTable(TargetSheet, TargetTable)
    Else
        Set TargetTable = CreateTable(TargetSheet)
        Call FormatTable(TargetSheet, TargetTable)
    End If
    
Footer:

End Sub
