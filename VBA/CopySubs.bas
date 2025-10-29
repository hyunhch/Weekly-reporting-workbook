Attribute VB_Name = "CopySubs"
Option Explicit

Function CopyRow(SourceSheet As Worksheet, SourceRange As Range, TargetSheet As Worksheet, TargetRange As Range) As Range
'Copies over rows directly from a passed cell in the "Select" column
'Destination does not need to have a table
'Returns range of all added rows

    Dim CopyRange As Range
    Dim PasteRange As Range
    Dim RowRange As Range
    Dim ReturnRange As Range
    Dim c As Range
    Dim d As Range
    Dim i As Long
    Dim j As Long
    
    'Find the width of the row we are copying
    Set c = SourceSheet.Cells(SourceRange.Row, 1)
    Set d = SourceSheet.Cells.Find("*", SearchOrder:=xlByColumns, SearchDirection:=xlPrevious)
        If d.Column = SourceRange.Column Then
            GoTo Footer
        End If
    
    Set RowRange = SourceSheet.Range(c, Cells(c.Row, d.Column).Address)
    
    'Loop through and copy over
    i = 1
    For Each c In SourceRange
        j = c.Row - RowRange.Row
        
        Set CopyRange = RowRange.Offset(j, 0)
        Set d = TargetSheet.Cells(TargetRange.Row, 1) 'Not sure if this helps
        Set PasteRange = d.Resize(1, CopyRange.Columns.Count).Offset(i - 1, 0)
        Set ReturnRange = BuildRange(PasteRange, ReturnRange)
        
        PasteRange.Value = CopyRange.Value
        i = i + 1
    Next c
    
    'Return
    Set CopyRow = ReturnRange
    
Footer:
    
End Function

Function CopyTableRow(CopySheet As Worksheet, PasteSheet As Worksheet, CopyCell As Range, Optional PasteCell As Range) As Range
'General function that tries to match headers between two tables and copy over
'Headers that are not found are ignored
'By default, copies to the bottom of the table
'If a PasteCell is passed, pastes in that row
'Returns the pasted row if successful, nothing on error

    Dim CopyRange As Range
    Dim PasteRange As Range
    Dim CopyHeaderRange As Range
    Dim PasteHeaderRange As Range
    Dim c As Range
    Dim d As Range
        
    Call UnprotectSheet(CopySheet)
    Call UnprotectSheet(PasteSheet)

    'Grab source and destination headers. Verifying there is a should be done in a parent sub
    Set CopyHeaderRange = CopySheet.ListObjects(1).HeaderRowRange
    Set PasteHeaderRange = PasteSheet.ListObjects(1).HeaderRowRange

    'Paste under the last row unless a PateCell was passed
    If PasteCell Is Nothing Then
        Set c = FindLastRow(PasteSheet)
            If c Is Nothing Then
                GoTo Footer
            End If
        
        Set PasteCell = c.Offset(1, 0)
    End If
    
    'Match headers and move over the information. Programatic for changes in order and when some aren't found
    For Each c In CopyHeaderRange
        Set d = FindTableHeader(PasteSheet, c.Value)
        
        If Not d Is Nothing Then
            Set CopyRange = CopySheet.Cells(CopyCell.Row, c.Column)
            Set PasteRange = PasteSheet.Cells(PasteCell.Row, d.Column)
                PasteRange.Value = CopyRange.Value
        End If
    Next c

    'Return row
    Set c = FindTableRow(PasteSheet, PasteCell)
        If c Is Nothing Then
            GoTo Footer
        End If
        
    Set CopyTableRow = c

Footer:

End Function

Function CopyNames(CheckRange As Range) As Range
'Copies only the checked first and last names
'For adding students to the Records Page, filtering should be done beforehand

    Dim RosterSheet As Worksheet
    Dim RecordsSheet As Worksheet
    Dim CopyRange As Range
    Dim PasteRange As Range
    Dim NameRange As Range
    Dim c As Range
    Dim ReturnRange As Range
    Dim i As Long
    Dim RosterTable As ListObject
    
    Set RosterSheet = Worksheets("Roster Page")
    Set RecordsSheet = Worksheets("Records Page")
    Set RosterTable = RosterSheet.ListObjects(1)
    
    'Find the last row
    Set c = RecordsSheet.Range("A:A").Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
        If c Is Nothing Then
            GoTo Footer
        End If
    
    Set PasteRange = c.Resize(1, 2) 'Not the row under
    
    'Nudge over from the select column to the First column
    Set c = RosterTable.ListColumns("First").DataBodyRange
        If c Is Nothing Then
            GoTo Footer
        End If
        
    Set NameRange = Intersect(CheckRange.EntireRow, c)
    
    'Loop through
    i = 1
    For Each c In NameRange
        Set CopyRange = c.Resize(1, 2)
            PasteRange.Offset(i, 0).Value = CopyRange.Value
            
        Set ReturnRange = BuildRange(PasteRange.Offset(i, 0), ReturnRange)
        
        i = i + 1
    Next c
    
    If ReturnRange Is Nothing Then
        GoTo Footer
    End If

    Set ReturnRange = ReturnRange.Resize(i - 1, 1)
    Set CopyNames = ReturnRange

Footer:

End Function

Sub CopyAllStudents(RosterSheet As Worksheet, TargetSheet As Worksheet, TargetCell As Range)
'Copies all students to the passed cell

    Dim RosterNameRange As Range
    Dim PasteRange As Range
    Dim RosterNameArray As Variant
    Dim RosterTable As ListObject

    Set RosterTable = RosterSheet.ListObjects(1)
    Set RosterNameRange = RosterTable.ListColumns("First").DataBodyRange
    
    'Resize to get last names and read into an array
    Set RosterNameRange = RosterNameRange.Resize(RosterNameRange.Rows.Count, 2)
    RosterNameArray = RosterNameRange.Value
    
    'Resize the TargetCell and assign values
    Set PasteRange = TargetCell.Resize(UBound(RosterNameArray), 2)
    PasteRange.Value = RosterNameArray

End Sub

Function CopyFromRecords(ActivitySheet As Worksheet, LabelCell As Range, Optional OperationString As String) As Range
'Grabs all students marked present or absent on the Records page for an activity and pastes them into an an Activity sheet
'Prompts for exporting and deleting if some students aren't found
'Passing "Present" grabs only present students
'Passing "Absent" grabs only absent students
    
    Dim RecordsSheet As Worksheet
    Dim RosterSheet As Worksheet
    Dim RosterNameRange As Range
    Dim RecordsAttendanceRange As Range
    Dim ActivityNameRange As Range
    Dim PresentRange As Range
    Dim AbsentRange As Range
    Dim CopyRange As Range
    Dim PasteRange As Range
    Dim c As Range
    Dim d As Range
    Dim ActivityTable As ListObject
    Dim RosterTable As ListObject
    
    Set RecordsSheet = Worksheets("Records Page")
    Set RosterSheet = Worksheets("Roster Page")
    
    'Grab attendance
    Set RecordsAttendanceRange = FindRecordsAttendance(RecordsSheet, , LabelCell)
        If RecordsAttendanceRange Is Nothing Then
            GoTo Footer
        'Make sure there are any students marked present or absent
        ElseIf FindStudentAttendance(RecordsSheet, RecordsAttendanceRange, "Both") Is Nothing Then
            GoTo Footer
        End If
        
    'Find all present and absent students
    Set c = FindStudentAttendance(RecordsSheet, RecordsAttendanceRange)
    Set d = FindStudentAttendance(RecordsSheet, RecordsAttendanceRange, "Absent")
        If c Is Nothing And d Is Nothing Then
            GoTo Footer
        End If
        
    'Nudge to the names
    If Not c Is Nothing Then
        Set PresentRange = c.Offset(0, -c.Column + 1)
    End If
    
    If Not d Is Nothing Then
        Set AbsentRange = d.Offset(0, -d.Column + 1)
    End If
    
    'Make sure we have a table on the ActivitySheet
    If CheckTable(ActivitySheet) > 3 Then
        Set ActivityTable = MakeActivityTable(ActivitySheet)
    End If
    
    Set ActivityTable = ActivitySheet.ListObjects(1)
    
    'Find where to start pasting
    Set PasteRange = FindLastRow(ActivitySheet).Offset(1, 0)
        If PasteRange Is Nothing Then 'This shouldn't be able to happen
            MsgBox ("Something has gone wrong. Please delete this sheet and create it again.")
            GoTo Footer
        End If
    
    'Match students and copy
    Set RosterTable = RosterSheet.ListObjects(1)
    Set RosterNameRange = RosterTable.ListColumns("First").DataBodyRange
    
    'Match present and absent students
    If Not PresentRange Is Nothing Then
        Set c = FindName(PresentRange, RosterNameRange)
    End If
    
    If Not AbsentRange Is Nothing Then
        Set d = FindName(AbsentRange, RosterNameRange)
    End If
    
    Select Case OperationString
    
        Case "Absent"
            Set CopyRange = d
        
        Case "Present"
            Set CopyRange = c
        
        Case Else
            If Not c Is Nothing Then
                Set CopyRange = c
            End If
            
            If Not d Is Nothing Then
                Set CopyRange = BuildRange(d, CopyRange)
            End If
            
    End Select
    
    If CopyRange Is Nothing Then
        GoTo Footer
    End If
    
    'Copy over and pull attendance
    Call CopyRow(RosterSheet, CopyRange.Offset(0, -1), ActivitySheet, PasteRange)
    
    ActivitySheet.Activate
    Call ActivityPullAttendenceButton
    
RemakeTable:
    Set ActivityTable = MakeActivityTable(ActivitySheet)
    
    Call TableFormat(ActivitySheet, ActivityTable)
    
    'Pull in attendance
    Set ActivityNameRange = ActivityTable.ListColumns("First").DataBodyRange
    Call ActivityPullAttendence(ActivitySheet, ActivityNameRange, LabelCell)
    
Footer:

End Function

Function CopyToActivity(RosterSheet As Worksheet, ActivitySheet As Worksheet, RosterAddRange As Range) As Range
'Copies unique students in the passed range
'Passes the "First" column

    Dim ActivityNameRange As Range
    Dim CopyRange As Range
    Dim PasteRange As Range
    Dim c As Range
    Dim i As Long
    Dim ActivityTable As ListObject
    
    If RosterAddRange Is Nothing Then
        GoTo Footer
    End If
    
    'Determine where to start pasting
    i = CheckTable(ActivitySheet)
    
    If i = 4 Then 'No table, this shouldn't happen
        Set ActivityTable = MakeActivityTable(ActivitySheet)
        Call TableFormat(ActivitySheet, ActivityTable)
    End If
    
    Set ActivityTable = ActivitySheet.ListObjects(1)
    Set PasteRange = FindLastRow(ActivitySheet).Offset(1, 0)
    
    If i = 3 Then 'No students, copy over everyone
        Set CopyRange = RosterAddRange
    Else
        'Find all the students checked on the Roster that are not on the Activity
        Set ActivityNameRange = ActivityTable.ListColumns("First").DataBodyRange
        Set CopyRange = FindUnique(RosterAddRange, ActivityNameRange)
        
        If CopyRange Is Nothing Then
            GoTo CleanUp
        End If
    End If

CopyStudents:
    'Copy over
    Call CopyRow(RosterSheet, CopyRange, ActivitySheet, PasteRange)
    
    Set ActivityTable = MakeActivityTable(ActivitySheet)
    Set ActivityNameRange = ActivityTable.ListColumns("First").DataBodyRange
    
CleanUp:
    'Remove any duplicates or blanks
    Call RemoveBadRows(ActivitySheet, ActivityTable.DataBodyRange, ActivityNameRange)
    
    'Remove any students on the ActivitySheet but not the RosterSheet. This shouldn't happen
    Set c = FindUnique(ActivityNameRange, RosterSheet.ListObjects(1).ListColumns("First").DataBodyRange)
        If Not c Is Nothing Then
            Call RemoveRows(ActivitySheet, ActivityTable.DataBodyRange, ActivityNameRange, c)
        End If
    
    'Return
    Set c = PasteStartRange.EntireColumn.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    Set CopyToActivity = ActivitySheet.Range(PasteStartRange, c)

Footer:

End Function

Function CopySelected(CopySheet As Worksheet, PasteSheet As Worksheet, Optional OperationString As String, Optional RawRange As Range) As Range
'Copies over non-duplicate students who are checked and returns the range of those added
'Returns nothing if all students were duplicates
'Pasting to the RecordsSheet is difference since there is no table
'Passing "All" ignores checks and copies all students
'Passing a range skips looking for checks

    Dim CopyNameRange As Range
    Dim CopyCheckRange As Range
    Dim PasteNameRange As Range
    Dim FilteredNameRange As Range
    Dim PasteStartRange As Range
    Dim PasteRange As Range
    Dim CopyRange As Range
    Dim c As Range
    Dim d As Range
    Dim i As Long
    Dim j As Long
    Dim CopyTable As ListObject
    Dim PasteTable As ListObject
    
    'Make sure there's a table with checked students
    If CheckTable(CopySheet) <> 1 Then
        GoTo Footer
    End If

    Call UnprotectSheet(CopySheet)
    Call UnprotectSheet(PasteSheet)
    
    'Define copy range
    Set CopyTable = CopySheet.ListObjects(1)
    
    If RawRange Is Nothing Then
        Set CopyNameRange = CopyTable.ListColumns("First").DataBodyRange
        Set CopyCheckRange = FindChecks(CopyTable.ListColumns("Select").DataBodyRange.SpecialCells(xlCellTypeVisible))
    Else
        Set CopyCheckRange = RawRange
    End If
    
    
    'Define paste ranges
    'Records Page
    If PasteSheet.Name = "Records Page" Then
        GoTo PasteToRecords
    End If

PasteToActivity:
    'See if there is anything in the table on the PasteSheet and define ranges
    i = CheckTable(PasteSheet)
    
    If i = 4 Then
        'No Table
        GoTo Footer
    Else
        Set PasteTable = PasteSheet.ListObjects(1)
    End If
    
    If i = 3 Then
        'No rows, paste under header
        Set c = PasteTable.HeaderRowRange.Find("First", , xlValues, xlWhole)
        Set PasteStartRange = c.Offset(1, 0)
    Else
        'At least one row, past at bottom of range
        Set PasteNameRange = PasteTable.ListColumns("First").DataBodyRange
        Set c = PasteNameRange.Resize(1, 1)
        Set PasteStartRange = c.Offset(PasteNameRange.Rows.Count)
    End If
    
    GoTo PasteStudents
    
PasteToRecords:
    'See if there are already students and define where to begin pasting
    If CheckRecords(PasteSheet) = 1 Or CheckRecords(PasteSheet) = 3 Then
        Set PasteNameRange = FindRecordsName(PasteSheet)
        Set c = PasteNameRange.Resize(1, 1)
        Set PasteStartRange = c.Offset(PasteNameRange.Rows.Count)
    Else
        Set c = FindRecordsName(PasteSheet) 'The "H BREAK" padding cell
        Set PasteStartRange = c.Offset(1, 0)
    End If

PasteStudents:
    'Need to nudge to the left if a range of names is passed instead of checks
    Set c = CopyCheckRange(1, 1)
    Set d = c.Offset(-c.Row + 1, 0)
    If Range(c, d).Find("First", , xlValues, xlWhole) Is Nothing Then
        Set CopyCheckRange = CopyCheckRange.Offset(0, 1)
    End If

    'If there are no students on the PasteSheet, copy all checked students
    'If "All" is passed, copy all students
    If PasteNameRange Is Nothing Then
        If OperationString = "All" Then
            Set FilteredNameRange = CopyNameRange
        Else
            Set FilteredNameRange = CopyCheckRange
        End If
    Else
        'Create a range of unique students
        Set FilteredNameRange = FindName(PasteSheet, PasteNameRange, CopyCheckRange, "Unique")
    End If
    
    If FilteredNameRange Is Nothing Then
        'No unique students
        GoTo Footer
    End If
    
    'Add students. How many columns are added depends on the sheet being pasted to
    i = 0
    For Each c In FilteredNameRange
        If PasteSheet.Name = "Records Page" Then
            Set CopyRange = c.Resize(1, 2)
            Set d = PasteStartRange.Resize(1, 2)
        Else
            Set CopyRange = c.Resize(1, CopyTable.ListColumns.Count - 1)
            Set d = PasteStartRange.Resize(1, CopyRange.Columns.Count - 1)
        End If
        
        Set PasteRange = d.Offset(i, 0)
        PasteRange.Value = CopyRange.Value
        
        i = i + 1
    Next c

    'Remake the table if not on RecordsSheet
    If PasteSheet.Name <> "Records Page" Then
        Set PasteTable = CreateTable(PasteSheet)
        Call FormatTable(PasteSheet, PasteTable)
    End If

    'ReturnRange
    Set CopySelected = PasteSheet.Range(PasteStartRange, PasteStartRange.Offset(i, 0))
        
Footer:

End Function

Function CopyToRecords(RecordsSheet As Worksheet, RosterSheet As Worksheet, RosterNameRange As Range) As Range
'Copies all students not already on the Roster to the Records page, returns a range of the students copied
'Returns nothing if no students are added
'Checks don't matter, called when the roster is parsed

    Dim ActivitySheet As Worksheet
    Dim RecordsNameRange As Range
    Dim RecordsAttendanceRange As Range
    Dim CopyRange As Range
    Dim PasteRange As Range
    Dim c As Range
    Dim i As Long

    'See if there are existing students on the RecordsSheet
    i = CheckRecords(RecordsSheet)
    
    If i = 2 Or i = 4 Then 'No students on the sheet
        Set CopyRange = RosterNameRange
        
        GoTo CopyStudents
    End If

    'Grab students on the RecordsSheet and identify any not the RosterSheet anymore
    Set RecordsNameRange = FindRecordsName(RecordsSheet)
    Set CopyRange = FindUnique(RosterNameRange, RecordsNameRange)
    Set DelRange = FindUnique(RecordsNameRange, RosterNameRange)
        If DelRange Is Nothing Then
            GoTo CopyStudents
        End If
     
    'Remove from any open activity sheet
    For Each ActivitySheet In ThisWorkbook.Sheets
        If ActivitySheet.Range("A1").Value = "Practice" Then
            'This checks if there's a table with students, matches names, removed, and saves the activity again
            Call RemoveFromActivity(ActivitySheet, DelRange)
        End If
    Next ActivitySheet
     
    'Delete students on the RecordsSheet but not the RosterSheet
    Set RecordsAttendanceRange = FindRecordsAttendance(RecordsSheet)
    Set c = RecordsSheet.Range(RecordsNameRange, RecordsAttendanceRange) 'Needs to include the first two columns

    Call RemoveRows(RecordsSheet, c, RecordsNameRange, DelRange)

CopyStudents:
    'No new students to add
    If CopyRange Is Nothing Then
        GoTo CleanUp
    End If
    
    'Copy over
    Set PasteRange = CopyNames(CopyRange)
    
CleanUp:
    Set RecordsNameRange = FindRecordsName(RecordsSheet) 'To account for added/removed students
    Set RecordsAttendanceRange = FindRecordsAttendance(RecordsSheet)
    Set c = RecordsSheet.Range(RecordsNameRange, RecordsAttendanceRange) 'Needs to include all columns for sorting
    
    Call RemoveBadRows(RecordsSheet, c, RecordsNameRange)

    'Return
    If CopyRange Is Nothing Then 'If no students were added
        GoTo Footer
    End If

    Set CopyToRecords = PasteRange

Footer:

End Function

Function CopyToReport(ReportSheet As Worksheet, LabelString As String, PasteArray As Variant) As Range
'SourceSheet can be the roster or an activity
'Can take the "Total" label to calculate

    Dim ReportHeaderRange As Range
    Dim ReportLabelRange As Range
    Dim TabulateRange As Range
    Dim PasteCell As Range
    Dim ReturnRange As Range
    Dim OtherCell As Range
    Dim c As Range
    Dim d As Range
    Dim i As Long
    Dim OtherValue As Long
    Dim TempHeader As String
    Dim ReportTable As ListObject
    Dim TempValue As Variant 'Can be strings
    
    Set ReportTable = ReportSheet.ListObjects(1)
    Set ReportHeaderRange = ReportTable.HeaderRowRange
    Set ReportLabelRange = ReportTable.ListColumns("Label").DataBodyRange

    Call UnprotectSheet(ReportSheet)
    
    'Check if the label already exists on the table. If not, add a new row
    Set PasteCell = FindReportLabel(ReportSheet, LabelString)
    
    'If not found, paste at the bottom of the table
    If PasteCell Is Nothing Then
        Set PasteCell = FindLastRow(ReportSheet, "Label").Offset(1, 0)
    End If
    
    'This will return the 2nd row if there are no activities
    If PasteCell.Value = "Total" And Not LabelString = "Total" Then
        Set PasteCell = PasteCell.Offset(1, 0)
    End If
    
    'Find each header in the passed array
    For i = 1 To UBound(PasteArray, 2)
        TempHeader = PasteArray(1, i)
        TempValue = PasteArray(2, i)
        
        Set c = ReportHeaderRange.Find(TempHeader, , xlValues, xlWhole)
        
        If c Is Nothing Then 'All bad entires go into the Other category
            If IsNumeric(TempValue) Then
                OtherValue = OtherValue + TempValue
            End If
            
            GoTo NextHeader
        ElseIf InStr(1, TempHeader, "Other") > 0 Then
            Set OtherCell = c.Offset(1, 0)
        End If
        
        Set d = ReportSheet.Cells(PasteCell.Row, c.Column)
        
        d.ClearContents
        d.Value = TempValue
        
        'No zeroes
        If d.Value = 0 Then
            d.ClearContents
        End If
        
        Set ReturnRange = BuildRange(d, ReturnRange)
NextHeader:
    Next i
    
    'Add any bad values to the 'Other' category, if it exists
    If OtherValue > 0 Then
        If Not OtherCell Is Nothing Then
            OtherCell.Value = OtherCell.Value + OtherValue
        End If
    End If
    
ReturnRange:
    If Not ReturnRange Is Nothing Then
        Set CopyToReport = ReturnRange
    End If

Footer:

End Function
