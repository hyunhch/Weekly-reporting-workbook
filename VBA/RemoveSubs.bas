Attribute VB_Name = "RemoveSubs"
Option Explicit

Function RemoveBadRows(TargetSheet As Worksheet, TableBodyRange As Range, SearchRange As Range, Optional SearchType As String) As Long
'Removes duplicate students and blank rows in a given table data body range
'Passing "Duplicate" or "Blank" will restrict deletions to those instances
'Returns the number of duplicates removed
'Returns nothing on error

    Dim c As Range
    Dim d As Range
    Dim DelRange As Range
    Dim i As Long
    
    i = 0
    
    If Not SearchType = "Duplicate" Then
       Set c = FindBlanks(SearchRange)
       
        If Not c Is Nothing Then
            Set DelRange = c
            
        End If
    End If
    
    If Not SearchType = "Blank" Then
        Set d = FindDuplicate(SearchRange)
        
        If Not d Is Nothing Then
            Set DelRange = BuildRange(d, DelRange)
            
            i = i + d.Cells.Count
        End If
    End If
    
    'Nothing to remove
    If DelRange Is Nothing Then
        RemoveBadRows = 0
        
        GoTo Footer
    End If

    'Delete
    'i = DelRange.Cells.Count
    Call RemoveRows(TargetSheet, TableBodyRange, SearchRange, DelRange)
    
    RemoveBadRows = i

Footer:

End Function

Sub RemoveFromActivity(ActivitySheet As Worksheet, DelRange As Range)
'Called whenever students need to be removed from an activity sheet
'Remove students and repulls the attendance
'DelRange can be from RosterSheet, the RecordsSheet, or the same ActivitySheet

    Dim RecordsSheet As Worksheet
    Dim ActivityNameRange As Range
    Dim ActivityDelRange As Range
    Dim LabelCell As Range
    Dim c As Range
    Dim DelConfirm As Long
    Dim ActivityTable As ListObject
    
    'Make sure there's a table with students
    If CheckTable(ActivitySheet) > 2 Then
        GoTo Footer
    End If
    
    Set RecordsSheet = Worksheets("Records Page")
    Set ActivityTable = ActivitySheet.ListObjects(1)
    Set ActivityNameRange = ActivityTable.ListColumns("First").DataBodyRange
    
    'Prompt if everyone is being deleted
    If DelRange.Cells.Count = ActivityNameRange.Cells.Count Then
        DelConfirm = MsgBox("This activity will be permanently deleted. Do you wish to continue?", vbQuestion + vbYesNo + vbDefaultButton2)
        
        If DelConfirm <> vbYes Then
            GoTo Footer
        End If
    End If
    
    'Nudge, if needed
    Set DelRange = NudgeToHeader(DelRange.Worksheet, DelRange, "First")
    
    'If the range of names is already on the sheet, skip name matching
    If DelRange.Worksheet.Name = ActivitySheet.Name Then
        Set ActivityDelRange = DelRange
    Else
        Set c = FindName(DelRange, ActivityNameRange)
        
        If Not c Is Nothing Then
            Set ActivityDelRange = c
        End If
    End If
    
    'No matches fouund
    If ActivityDelRange Is Nothing Then
        GoTo Footer
    End If
    
    'Remove students and save
    Call UnprotectSheet(ActivitySheet)
    
    Set LabelCell = FindActivityLabel(ActivitySheet)
        If LabelCell Is Nothing Then 'This shouldn't happen
            MsgBox ("Something has gone wrong. Please delete this sheet and remake it.")
            GoTo Footer
        End If
    
    Call RemoveRows(ActivitySheet, ActivityTable.DataBodyRange, ActivityNameRange, ActivityDelRange)
    
    'Repull attendance if there are students left
    Set ActivityTable = ActivitySheet.ListObjects(1)
    Set ActivityNameRange = ActivityTable.ListColumns("First").DataBodyRange
        If ActivityNameRange Is Nothing Then
            'Clear from the Records and Report if no students remain
            Call RemoveRecordsActivity(RecordsSheet, LabelCell) 'Removes from both
            
            GoTo Footer
        End If

    Call ActivityPullAttendence(ActivitySheet, ActivityNameRange, LabelCell)
    
    'Retabulate
    Call ActivitySave(ActivitySheet, RecordsSheet, LabelCell)
    
Footer:

End Sub

Sub RemoveFromOpenActivity(DelRange As Range)
'Small helper function that loops through all sheets in the workbook and passes to RemoveFromActivity

    Dim ActivitySheet As Worksheet
    Dim ActivityTable As ListObject
    
    For Each ActivitySheet In ThisWorkbook.Sheets
        If Not ActivitySheet.Range("A1").Value = "Practice" Then
            GoTo NextSheet
        End If
        
        Call RemoveFromActivity(ActivitySheet, DelRange)
NextSheet:
    Next ActivitySheet

Footer:

End Sub

Sub RemoveFromRecords(RosterSheet As Worksheet, RosterDelRange As Range)
'Deletes passed students on the Records page
'Removes student from the roster and any open activity sheet
'Does NOT export, that should happen in a parent sub
'Retabulates activities
    
    Dim RecordsSheet As Worksheet
    Dim ReportSheet As Worksheet
    Dim RecordsNameRange As Range
    Dim RecordsDelRange As Range
    Dim RecordsRange As Range
    Dim ReportLabelRange As Range
    Dim c As Range
    Dim i As Long
    Dim LabelString As String
    Dim RosterTable As ListObject
    Dim LabelArray As Variant
    
    Set ReportSheet = Worksheets("Report Page")
    Set RecordsSheet = Worksheets("Records Page")
    Set RecordsRange = FindRecordsRange(RecordsSheet)
    Set RosterTable = RosterSheet.ListObjects(1)
    
    'Ensure that there are students to remove
    i = CheckRecords(RecordsSheet)
        If i = 2 Or i = 4 Then 'No students
            GoTo Footer
        End If
    
    'Define the range of names and match
    Set RecordsNameRange = FindRecordsName(RecordsSheet)
    Set RecordsDelRange = FindName(RosterDelRange, RecordsNameRange)
        If RecordsDelRange Is Nothing Then
            GoTo Footer
        End If
    
    'Loop through any open Activity Sheet and remove students
    'Call RemoveFromOpenActivity(RecordsDelRange) 'This retabulates the activities
    
    'Remove from the RosterSheet
    'Call RemoveRows(RosterSheet, RosterTable.DataBodyRange, RosterTable.ListColumns("First").DataBodyRange, RosterDelRange)
    
    'Remove from the RecordsSheet
    Call RemoveRows(RecordsSheet, RecordsRange, RecordsNameRange, RecordsDelRange)
    
    'Check what's on the report
    i = CheckReport(ReportSheet)
        If Not i < 3 Then 'No activities tabulated
            GoTo Footer
        End If
    
    'Grab activities currently on the report
    Set ReportLabelRange = FindReportLabel(ReportSheet)
    
    ReDim LabelArray(1 To ReportLabelRange.Cells.Count)
    i = 1
    
    For Each c In ReportLabelRange
        LabelArray(i) = c.Value
    
        i = i + 1
    Next c
    
    'Clear the report and retabulate activities
    Call ReportClear
    Call TabulateReportTotals
    
    For i = 1 To UBound(LabelArray)
        LabelString = LabelArray(i)
        
        Set c = FindRecordsLabel(RecordsSheet, , LabelString)
            If c Is Nothing Then
                GoTo NextLabel
            End If
        
        Call TabulateActivity(c)
NextLabel:
    Next i
    
Footer:

End Sub

Sub RemoveFromReport(DelRange As Range)
'Removes one or more activities from the ReportSheet
'DelRange should be cells containing labels, so this is different from the button to remove rows that looks for checks

    Dim ReportSheet As Worksheet
    Dim NudgeDelRange As Range
    Dim ReportDelRange As Range
    Dim ReportLabelRange As Range
    Dim c As Range
    Dim d As Range
    Dim ReportTable As ListObject
    
    Set ReportSheet = Worksheets("Report Page")
    Set ReportTable = ReportSheet.ListObjects(1)
    
    'Check that there's a table with more than 2 rows
    If CheckReport(ReportSheet) > 2 Then
        GoTo Footer
    End If
    
    Call UnprotectSheet(ReportSheet)
    
    'If we're already on the Report, we don't need to search
    If DelRange.Worksheet.Name = "Report Page" Then
        Set c = DelRange
        Set ReportDelRange = NudgeToHeader(ReportSheet, c, "Label")
        
        GoTo RemoveActivities
    End If
    
    'If called from a different sheet, make a range to remove
    For Each c In DelRange
        Set d = FindReportLabel(ReportSheet, c.Value)
        
        If Not d Is Nothing Then
            Set ReportDelRange = BuildRange(c, ReportDelRange)
        End If
    Next c
    
RemoveActivities:
    If ReportDelRange Is Nothing Then
        GoTo Footer
    End If

    'Pass to remove. Don't pass the totals row so it isn't sorted
    Set c = FindReportLabel(ReportSheet)
        If c Is Nothing Then
            GoTo Footer
        End If
    
    If ReportTable.Range.Rows.Count > 3 Then
        Set ReportLabelRange = c.Resize(c.Rows.Count - 1, c.Columns.Count).Offset(1, 0) 'Need at least two rows
        
        If ReportLabelRange Is Nothing Then
            GoTo Footer
        End If
    Else
        Set ReportLabelRange = c
    End If
    
    Set c = ReportTable.DataBodyRange
    Set d = c.Resize(c.Rows.Count - 1, c.Columns.Count).Offset(1, 0)
        If d Is Nothing Then
            GoTo Footer
        End If
        
    Call RemoveRows(ReportSheet, d, ReportLabelRange, ReportDelRange)

Footer:

End Sub

Function RemoveFromRoster(RosterSheet As Worksheet, DelRange As Range, Optional PromptString As String) As Long
'Called when a student is removed from the RosterSheet
'Removes the student from any open Activity sheets, the Records sheet, the Roster sheet, then retabulates
'Checking for a table with rows should be done in a previous sub
'Passing "Prompt" shows the delete and export prompts
'Returns 1 if successful
        
    Dim OldBook As Workbook
    Dim NewBook As Workbook
    Dim RosterNameRange As Range
    Dim i As Long
    Dim RosterTable As ListObject
    Dim DeleteAll As Boolean
    
    RemoveFromRoster = 0
    DeleteAll = False
    
    Set OldBook = ThisWorkbook
    Set RosterTable = RosterSheet.ListObjects(1)
    Set RosterNameRange = RosterTable.ListColumns("First").DataBodyRange
    
    'If this gets passed after the range has been deleted
    If DelRange Is Nothing Then
        GoTo Footer
    'If all students are being removed, we can skip some steps
    ElseIf DelRange.Offset(0, 1).Address = RosterNameRange.Address Then
        DeleteAll = True
    End If
    
    If PromptString <> "Prompt" Then
        GoTo ActivityRemove
    End If
    
    'Prompt to confirm deletion
    i = PromptRemoveRoster(DelRange)
        If i <> 1 Then
            GoTo Footer
        End If
    
    'Prompt to confirm exporting
    i = PromptExport(DelRange)
        If i <> 1 Then
            GoTo ActivityRemove
        End If
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    
    'Generate a new book
    Set NewBook = ExportFromRoster(DelRange)
        If NewBook Is Nothing Then
            GoTo ActivityRemove
        End If
        
    'Save and close the new book
    If ExportLocalSave(OldBook, NewBook) > 0 Then
        On Error Resume Next
        NewBook.Close savechanges:=False
        On Error GoTo 0
    End If
    
ActivityRemove:
    'Loop through and remove from any open activity sheets. Empty sheets are deleted and the activity deleted from the Records and Report
    Call RemoveFromOpenActivity(DelRange)
      
    'Wipe the Records and Report if everyone is deleted
    If DeleteAll = True Then
        Call RecordsClear
        Call ReportClear
        Call RosterClearButton
        
        GoTo Footer
    End If
      
    'Remove from the Records sheet. This retabulates activities as well
    Call RemoveFromRecords(RosterSheet, DelRange)

    'Finally, delete the students from the Roster and retabulate totals
    Call RemoveRows(RosterSheet, RosterTable.DataBodyRange, RosterNameRange, DelRange)
    Call TabulateReportTotals
    
Footer:

End Function

Function RemoveNonNumeric(FullString As String) As String 'or as long, if you prefer
    Dim re As Object
    
    Set re = CreateObject("VBScript.Regexp")
    With re
        .Pattern = "[^.0-9]"
        .Global = True
        RemoveNonNumeric = .Replace(FullString, "")
    End With
End Function

Sub RemoveRecordsActivity(RecordsSheet As Worksheet, LabelCell As Range)
'Deletes an activity from the Records sheet and Report sheet, if present
'Deletes any open activity sheets of the activity being removed
'Loops through passed range and deletes the column

    Dim ReportSheet As Worksheet
    Dim ActivitySheet As Worksheet
    Dim RecordsDelRange As Range
    Dim ReportDelRange As Range
    Dim LabelString As String
    
    Set ReportSheet = Worksheets("Report Page")
    
    'Make sure there's something to delete
    If LabelCell Is Nothing Then
        GoTo Footer
    ElseIf CheckRecords(RecordsSheet) > 2 Then
        GoTo Footer
    End If

    'Find the column we want on the RecordsSheet and delete
    LabelString = LabelCell.Value
        If Not Len(LabelString) > 0 Then
            GoTo Footer
        End If
    
    Set RecordsDelRange = FindRecordsLabel(RecordsSheet, LabelCell)
        If RecordsDelRange Is Nothing Then
            GoTo RemoveReportActivity
        End If
        
    RecordsDelRange.EntireColumn.Delete
    
RemoveReportActivity:
    'Pass to remove from the Report Sheet
    Set ReportDelRange = FindReportLabel(ReportSheet, LabelString)
        If ReportDelRange Is Nothing Then
            GoTo RemoveOpenSheet
        End If

    Call RemoveFromReport(ReportDelRange)

RemoveOpenSheet:
    'Loop through ActivitySheets and delete if there's a match
    Set ActivitySheet = FindSheet(LabelString)
        If ActivitySheet Is Nothing Then
            GoTo Footer
        End If
    
    ActivitySheet.Delete

Footer:

End Sub

Sub RemoveRows(TargetSheet As Worksheet, SearchRange As Range, SortRange As Range, DelRange As Range)
'SearchRange is the bound of what to delete, done to avoid some errors with tables
'SortRange is the column being sorted, usually "Select" or "First"
'Del range are the cells in SortRange to delete. The row is removed
'Needs to be passed the SearchRange to sort, i.e. a table DataBodyRange

    Dim SortDelRange As Range
    Dim c As Range
    Dim d As Range
    Dim i As Long
    Dim TargetTable As ListObject
    Dim HasTable As Boolean
    
    Call UnprotectSheet(TargetSheet)

    'I don't think removing that table is needed since I'm defining a number of cells to be deleted rather than the entire row. Need to test
    'Remove any table and formatting
    If TargetSheet.ListObjects.Count > 0 Then
        HasTable = True
        
        'Nudge to the select column
        Set SortRange = NudgeToHeader(TargetSheet, SortRange, "Select")
        Set DelRange = NudgeToHeader(TargetSheet, DelRange, "Select")
        
        Call RemoveTable(TargetSheet)
    End If
    
    SearchRange.FormatConditions.Delete
    
    'Flag each row to be deleted
    DelRange.Interior.Color = vbRed
    
    'Sort by color
    With TargetSheet.Sort
        .SortFields.Clear
        .SortFields.Add2(SortRange.Offset, xlSortOnCellColor, xlAscending, , xlSortNormal).SortOnValue.Color = RGB(255, 0, 0)
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
        Set TargetTable = MakeReportTable
        Call TableFormatReport(TargetSheet, TargetTable)
    Else
        Set TargetTable = MakeTable(TargetSheet)
        Call TableFormat(TargetSheet, TargetTable)
    End If
    
Footer:

End Sub

Sub RemoveTable(TargetSheet As Worksheet)
'Unlists all table objects and removes formatting

    Dim OldTableRange As Range
    Dim OldTable As ListObject
    
    Call UnprotectSheet(TargetSheet)
    
    For Each OldTable In TargetSheet.ListObjects
        Set OldTableRange = OldTable.Range
        
        OldTable.Unlist
        OldTableRange.FormatConditions.Delete
        OldTableRange.Borders.LineStyle = Excel.XlLineStyle.xlLineStyleNone
    Next OldTable

End Sub

