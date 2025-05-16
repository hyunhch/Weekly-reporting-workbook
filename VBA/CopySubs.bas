Attribute VB_Name = "CopySubs"
Option Explicit

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

Sub CopyFromRecordsTest()

    Dim RecordsSheet As Worksheet
    Dim ActivitySheet As Worksheet
    Dim LabelCell As Range
    Dim OperationString As String
    
    Set RecordsSheet = Worksheets("Records Page")
    Set ActivitySheet = ActiveSheet
    Set LabelCell = RecordsSheet.Range("G1")
    
    'Call CreateActivityTable(ActivitySheet)
    
    Call CopyFromRecords(RecordsSheet, ActivitySheet, LabelCell)

End Sub

Function CopyFromRecords(RecordsSheet As Worksheet, ActivitySheet As Worksheet, LabelCell As Range, Optional OperationString As String) As Range
'Grabs all students marked present or absent on the Records page for an activity and pastes them into an an Activity sheet
'Prompts for exporting and deleting if some students aren't found
'Passing "Present" grabs only present students
'Passing "Absent" grabs only absent students

    Dim RosterSheet As Worksheet
    Dim RosterNameRange As Range
    Dim RosterCopyRange As Range
    Dim RecordsAttendanceRange As Range
    Dim RecordsCopyRange As Range
    Dim ActivityNameHeader As Range
    Dim ActivityNameRange As Range
    Dim PresentRange As Range
    Dim AbsentRange As Range
    Dim TempRange As Range
    Dim CopyRange As Range
    Dim PasteRange As Range
    Dim c As Range
    Dim d As Range
    Dim ActivityTable As ListObject
    Dim RosterTable As ListObject
    
    Set RosterSheet = Worksheets("Roster Page")
    
    'Make sure there are students and activities
    If CheckRecords(RecordsSheet) > 1 Then
        GoTo Footer
    End If
    
    'Grab attendance
    Set RecordsAttendanceRange = FindRecordsAttendance(RecordsSheet, , LabelCell)
    
    If RecordsAttendanceRange Is Nothing Then
        GoTo Footer
    End If
    
    'Find all present and absent students
    For Each c In RecordsAttendanceRange
        If c.Value = "1" Then
            Set PresentRange = BuildRange(c.Offset(0, -c.Column + 1), PresentRange) 'Move to where the names are
        ElseIf c.Value = "0" Then
            Set AbsentRange = BuildRange(c.Offset(0, -c.Column + 1), AbsentRange)
        End If
    Next c
    
    'Make sure we have a table on the ActivitySheet
    If CheckTable(ActivitySheet) > 3 Then
        Set ActivityTable = CreateActivityTable(ActivitySheet)
    Else
        Set ActivityTable = ActivitySheet.ListObjects(1)
        Set ActivityNameRange = ActivityTable.ListColumns("First").DataBodyRange
    End If
    
    'Find where to start copying
    Set ActivityNameHeader = FindTableHeader(ActivitySheet, "First")
    
    If ActivityNameHeader Is Nothing Then 'This shouldn't be able to happen
        MsgBox ("Something has gone wrong. Please delete this sheet and create it again.")
        GoTo Footer
    End If
    
    'Match students and copy
    Set RosterTable = RosterSheet.ListObjects(1)
    Set RosterNameRange = RosterTable.ListColumns("First").DataBodyRange
    
    'All present students
    If OperationString <> "Absent" Then
        Set PasteRange = ActivityNameHeader.EntireColumn.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Offset(1, 0)
        Set TempRange = FindName(PresentRange, RosterNameRange)
        
        If ActivityNameRange Is Nothing Then
            'Copy all present students
            Set RosterCopyRange = TempRange
        Else
            'Only copy new present students
            Set RosterCopyRange = FindUnique(TempRange, ActivityNameRange)
        End If
        
        If RosterCopyRange Is Nothing Then
            GoTo CopyAbsent
        End If
            
        For Each c In RosterCopyRange
            Set d = c.Resize(1, RosterTable.ListColumns.Count - 1)
            Set CopyRange = BuildRange(d, CopyRange)
        Next c
        
        'Copy and mark checked
        Set c = CopyRows(RosterSheet, CopyRange, ActivitySheet, PasteRange)
        
        c.Offset(0, -1).Value = "a"
        
        'Return range part 1
        Set CopyFromRecords = c
    End If

CopyAbsent:
    'All absent students
    If OperationString <> "Present" Then
        Set PasteRange = ActivityNameHeader.EntireColumn.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Offset(1, 0)
        Set TempRange = FindName(AbsentRange, RosterNameRange)
        
        If ActivityNameRange Is Nothing Then
            'Copy all absent students
            Set RosterCopyRange = TempRange
        Else
            'Only copy new absent students
            Set RosterCopyRange = FindUnique(TempRange, ActivityNameRange)
        End If
        
        If RosterCopyRange Is Nothing Then
            GoTo RemakeTable
        End If
        
        Set CopyRange = Nothing 'Otherwise the range for present students is still included
        For Each c In RosterCopyRange
            Set d = c.Resize(1, RosterTable.ListColumns.Count - 1)
            Set CopyRange = BuildRange(d, CopyRange)
        Next c
    
        'Copy and mark unchecked
        Set c = CopyRows(RosterSheet, CopyRange, ActivitySheet, PasteRange)
        
        c.Offset(0, -1).Value = ""
        
        'ReturnRange part 2
        Set CopyFromRecords = BuildRange(c, CopyFromRecords)
    End If
    
RemakeTable:
    Set ActivityTable = CreateActivityTable(ActivitySheet)
    
    Call FormatTable(ActivitySheet, ActivityTable)
    
    'Pull in attendance
    Set ActivityNameRange = ActivityTable.ListColumns("First").DataBodyRange
    Call RecordsPullAttendance(ActivitySheet, ActivityNameRange, LabelCell)
    
Footer:

End Function

Function CopyRows(SourceSheet As Worksheet, SourceRange As Range, TargetSheet As Worksheet, TargetRange As Range) As Range
'Pastes starting at the PasteStart
'Checking that there are no duplicates, etc. should all be done in parent function
'The entire portion of the row to be moved over are passed

    Dim CopyRange As Range
    Dim PasteRange As Range
    Dim ReturnRange As Range
    Dim c As Range
    Dim d As Range
    Dim i As Long
    
    i = 0
    For Each c In SourceRange.Rows
        Set CopyRange = c
        Set d = TargetRange.Resize(1, c.Columns.Count)
        Set PasteRange = d.Offset(i, 0)
        Set ReturnRange = BuildRange(TargetRange.Offset(i, 0), ReturnRange)
        
        PasteRange.Value = CopyRange.Value
        i = i + 1
    Next c
    
    Set CopyRows = ReturnRange
    
End Function

Function CopyToActivity(RosterSheet As Worksheet, ActivitySheet As Worksheet, RosterAddRange As Range) As Range
'Copies unique students in the passed range

    Dim PasteStartRange As Range
    Dim ActivityNameHeader As Range
    Dim ActivityNameRange As Range
    Dim FilteredAddRange As Range
    Dim CopyRange As Range
    Dim ActivityDelRange As Range
    Dim c As Range
    Dim d As Range
    Dim e As Range
    Dim i As Long
    Dim j As Long
    Dim HeaderArray() As Variant
    Dim ActivityTable As ListObject
    
    'Start with the full list of students
    Set FilteredAddRange = RosterAddRange
    
    'Determine where to start pasting
    i = CheckTable(ActivitySheet)
    
    If i = 4 Then 'No students, this shouldn't happen
        Set c = ActivitySheet.Range("A6")
        
        j = 1
        ReDim HeaderArray(1 To RosterSheet.ListObjects(1).ListColumns.Count)
        For Each d In RosterSheet.ListObjects(1).HeaderRowRange
            HeaderArray(j) = d.Value
            
            j = j + 1
        Next d
        
        Call ResetTableHeaders(ActivitySheet, c, HeaderArray)
        Call CreateTable(ActivitySheet)
    End If
    
    Set ActivityTable = ActivitySheet.ListObjects(1)
    Set ActivityNameHeader = FindTableHeader(ActivitySheet, "First")
        
    If i = 3 Then 'No students
        Set PasteStartRange = ActivityNameHeader.Offset(1, 0)
    Else
        Set ActivityNameRange = ActivityTable.ListColumns("First").DataBodyRange
        Set PasteStartRange = ActivityNameHeader.Offset(ActivityTable.Range.Rows.Count, 0)
        
        'Find all the students checked on the Roster that are not on the Activity
        Set FilteredAddRange = FindUnique(RosterAddRange, ActivityNameRange)

        If FilteredAddRange Is Nothing Then
            GoTo CleanUp:
        End If
    End If
    
    'Resize to the entire row, minus the first column
    For Each c In FilteredAddRange
        Set d = c.Resize(1, ActivityTable.ListColumns.Count - 1)
        Set e = RosterSheet.Range(c, d)
        Set CopyRange = BuildRange(e, CopyRange)
    Next c
    
    'Set c = FilteredAddRange
    'Set d = FilteredAddRange.Offset(0, ActivityTable.ListColumns.Count - ActivityNameHeader.Column)
    'Set FilteredAddRange = RosterSheet.Range(c, d)


CopyStudents:
    'Copy over
    Call CopyRows(RosterSheet, CopyRange, ActivitySheet, PasteStartRange)
    
    Set ActivityTable = CreateTable(ActivitySheet)
    Set ActivityNameRange = ActivityTable.ListColumns("First").DataBodyRange
    
CleanUp:
    'If there are students on the ActivitySheet but not the Roster, duplicates, or blank rows
    Set c = FindUnique(ActivityNameRange, RosterSheet.ListObjects(1).ListColumns("First").DataBodyRange)
    Set d = FindDuplicate(ActivityNameRange)
    Set e = FindBlanks(ActivitySheet, ActivityNameRange)

    'Build into a single range
    If Not c Is Nothing Then
        Set ActivityDelRange = c
    End If
    
    If Not d Is Nothing Then
        If ActivityDelRange Is Nothing Then
            Set ActivityDelRange = d
        Else
            Set ActivityDelRange = Union(ActivityDelRange, d)
        End If
    End If
    
    If Not e Is Nothing Then
        If ActivityDelRange Is Nothing Then
            Set ActivityDelRange = e
        Else
            Set ActivityDelRange = Union(ActivityDelRange, e)
        End If
    End If
    
    'Delete the combined range. The table is remade in the child sub
    If Not ActivityDelRange Is Nothing Then
        Call RemoveRows(ActivitySheet, ActivityTable.DataBodyRange, ActivityNameRange, ActivityDelRange)
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

    Dim OldBook As Workbook
    Dim NewBook As Workbook
    Dim ActivitySheet As Worksheet
    Dim ActivityDelRange As Range
    Dim RecordsNameRange As Range
    Dim RecordsFullNameRange As Range
    Dim RecordsLabelRange As Range
    Dim AddStudentRange As Range
    Dim RemoveStudentRange As Range
    Dim PasteStartRange As Range
    Dim c As Range
    Dim d As Range
    Dim ExportConfirm As Long
    Dim i As Long
    Dim j As Long

    'Find names on RecordsSheet. It will be the "H BREAK" padding cell if there are none
    Set RecordsNameRange = FindRecordsName(RecordsSheet)

    'See if there are existing students
    i = CheckRecords(RecordsSheet)
    
    If i = 2 Or i = 4 Then 'No students on the sheet
        Set AddStudentRange = RosterNameRange
        GoTo CopyStudents
    End If

    'See if there are students on the Records sheet but not on the Roster sheet anymore
    Set RecordsFullNameRange = Union(RecordsNameRange, RecordsNameRange.Offset(0, 1))
    Set AddStudentRange = FindUnique(RosterNameRange, RecordsNameRange)
    Set RemoveStudentRange = FindUnique(RecordsNameRange, RosterNameRange)

    If RemoveStudentRange Is Nothing Then
        GoTo CopyStudents
    End If
    
    'Prompt to export
    j = RemoveStudentRange.Cells.Count
    ExportConfirm = MsgBox(j & " students are no longer on your roster. " & _
        "Do you want to save a copy of these students' attendance before removing them?", vbQuestion + vbYesNo + vbDefaultButton2)
        
    If ExportConfirm = vbYes Then
        Set OldBook = ThisWorkbook
        Set NewBook = MakeNewBook(RecordsSheet, RosterSheet, RemoveStudentRange)
        
        Call SaveNewBook(OldBook, NewBook)
        OldBook.Activate
    End If
    
    'Remove from any open activity sheet
    For Each ActivitySheet In ThisWorkbook.Sheets
        If ActivitySheet.Range("A1").Value = "Practice" Then
            'This checks if there's a table with students, matches names, removed, and saved the activity again
            Call RemoveFromActivity(ActivitySheet, RemoveStudentRange)
        End If
    Next ActivitySheet

    'Remove students from RecordsSheet
    Call RemoveRows(RecordsSheet, RecordsFullNameRange.EntireRow, RecordsNameRange, RemoveStudentRange)

CopyStudents:
    'No new students to add
    If AddStudentRange Is Nothing Then
        GoTo CleanUp
    End If

    'Copy over unique students
    Set c = RecordsNameRange.Resize(1, 1)
    Set PasteStartRange = c.Offset(RecordsNameRange.Rows.Count, 0)
    Set d = Union(AddStudentRange, AddStudentRange.Offset(0, 1))
    
    Call CopyRows(RosterSheet, d, RecordsSheet, PasteStartRange)
    
CleanUp:
    Set RecordsNameRange = FindRecordsName(RecordsSheet) 'To account for added/removed students
    Set c = FindDuplicate(RecordsNameRange)
    Set d = FindBlanks(RecordsSheet, RecordsNameRange)
    
    'Duplicates
    If Not c Is Nothing Then
        Set RemoveStudentRange = c
    End If
    
    'Blanks
    If Not d Is Nothing Then
        If RemoveStudentRange Is Nothing Then
            Set RemoveStudentRange = d
        Else
            Set RemoveStudentRange = Union(RemoveStudentRange, d)
        End If
    End If
    
    'Remove duplicates and blank rows
    If Not c Is Nothing Or Not d Is Nothing Then
        Call RemoveRows(RecordsSheet, RecordsFullNameRange.EntireRow, RecordsNameRange, RemoveStudentRange)
    End If
    
    'Return
    If AddStudentRange Is Nothing Then 'If no students were added
        GoTo Footer
    End If
    
    Set c = RecordsSheet.Range("A:A").Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    Set CopyToRecords = RecordsSheet.Range(PasteStartRange, c)

Footer:

End Function

Function CopyToReport(ReportSheet As Worksheet, PasteCell As Range, PasteArray As Variant) As Range
'Copies values passed in the array to the row of the PasteCell
'Can pass anything in the report header
    
    Dim ReportHeaderRange As Range
    Dim ReturnRange As Range
    Dim c As Range
    Dim d As Range
    Dim PasteOffset As Long
    Dim i As Long
    Dim j As Long
    Dim OtherIndex As Long
    Dim OtherString As String
    Dim HeaderString As String
    Dim ReportTable As ListObject

    Set ReportTable = ReportSheet.ListObjects(1)
    Set ReportHeaderRange = ReportTable.HeaderRowRange

    Call UnprotectSheet(ReportSheet)
    
    'Push to Report Page
    j = 0
    For i = 1 To UBound(PasteArray)
        HeaderString = PasteArray(i, 1)
        
        'Grab the index of the "Other" category
        If InStr(1, HeaderString, "Other") > 0 Then
            OtherIndex = i
            OtherString = PasteArray(OtherIndex, 1)
        End If
        
        Set c = ReportHeaderRange.Find(HeaderString, , xlValues, xlWhole)
        
        If Not c Is Nothing Then
            'Paste under the matching header
            Set d = ReportSheet.Cells(PasteCell.Row, c.Column)
            d.Value = PasteArray(i, 2)
            Set ReturnRange = BuildRange(d, ReturnRange)
            
            'Get rid of zeroes
            If d.Value = 0 Then
                d.ClearContents
            End If
        Else
            'Sum up elements that aren't found in the header
            If IsNumeric(PasteArray(i, 2)) = True Then
                j = j + PasteArray(i, 2) 'Using this so we don't run into problems with strings
            End If
        End If
    Next i

    'Skip adding up "others" if no other category was found, such as with Low Income
    If OtherIndex < 1 Then
        GoTo ReturnRange
    End If

    'Push all leftover elements into the "Other" category
    'This will allow the list of categories to change in the future
    If j > 0 Then
        PasteArray(OtherIndex, 2) = PasteArray(OtherIndex, 2) + j
        Set c = ReportHeaderRange.Find(OtherString, , xlValues, xlWhole)
        Set d = ReportSheet.Cells(PasteCell.Row, c.Column)
        
        d.Value = PasteArray(OtherIndex, 2)
        Set ReturnRange = BuildRange(d, ReturnRange)
    End If

ReturnRange:
    If Not ReturnRange Is Nothing Then
        Set CopyToReport = ReturnRange
    End If

Footer:

End Function
