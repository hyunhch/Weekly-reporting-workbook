Attribute VB_Name = "CheckSubs"
Option Explicit

Function CheckCover() As Long
'Returns 1 if all of the information is filled out on the CoverSheet

    Dim CoverSheet As Worksheet
    Dim c As Range
    Dim i As Long
    Dim SearchString As String
    Dim SearchArray() As Variant
    
    Set CoverSheet = Worksheets("Cover Page")
    
    ReDim SearchArray(1 To 3)
    SearchArray(1) = "Name"
    SearchArray(2) = "Date"
    SearchArray(3) = "Center"

    CheckCover = 0

    For i = 1 To UBound(SearchArray)
        SearchString = SearchArray(i)
        Set c = CoverSheet.Range("A:A").Find(SearchString, , xlValues, xlWhole).Offset(0, 1)
        
        If Len(c.Value) < 1 Then
            GoTo Footer
        End If
    Next i

    'If nothing failed
    CheckCover = 1

Footer:

End Function

Function CheckAttendance(RecordsSheet As Worksheet, NameCell As Range, Optional CountAbsent As String) As Long
'Returns 1 if the passed student was present for anything, 0 otherwise
'Passing "Absent" will consider both present (1) and absent (0) as attending

    Dim AttendanceRange As Range
    Dim i As Long
    
    Set AttendanceRange = FindRecordsAttendance(RecordsSheet, NameCell)
    
    'Either sum or count cells with a number in them
    If Not AttendanceRange Is Nothing Then
        i = WorksheetFunction.Sum(AttendanceRange)
    ElseIf Not AttendanceRange Is Nothing And CountAbsent = "Absent" Then
        i = WorksheetFunction.CountA(AttendanceRange)
    Else
        GoTo Footer
    End If

    'Return a binary answer
    If i > 0 Then
        i = 1
    End If
    
    CheckAttendance = i

Footer:
    
End Function

Function CheckRecords(RecordsSheet As Worksheet) As Long
'Four possibilities:
'4 - No students or activities
'3 - Students but no activities
'2 - Activities but no students
'1 - Both students and activities

    Dim LRowCell As Range
    Dim LColCell As Range

    Set LRowCell = RecordsSheet.Range("A:A").Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    Set LColCell = RecordsSheet.Range("1:1").Find("*", SearchOrder:=xlByColumns, SearchDirection:=xlPrevious)
    
    CheckRecords = 1
    If LRowCell = "H BREAK" Then
        CheckRecords = CheckRecords + 1
    End If
    
    If LColCell = "V BREAK" Then
        CheckRecords = CheckRecords + 2
    End If

End Function

Function CheckReport(ReportSheet As Worksheet) As Long
'Checks that there are student totals and activities
    '1 - Totals, activities, and check
    '2 - Totals and at least one activities
    '3 - Totals but no activities
    '4 - No table
    
    Dim c As Range
    Dim i As Long
    Dim ReportTable As ListObject
    
    'Is there a table. There always should be
    If ReportSheet.ListObjects.Count < 1 Then
        i = 4
        GoTo Footer
    End If
    
    'If there are any totals
    Set ReportTable = ReportSheet.ListObjects(1) 'There will always be at least two rows
    
    If ReportTable.Range.Rows.Count = 2 Then
        i = 3
        GoTo Footer
    End If
    
    'If there are any activities
    If ReportTable.Range.Rows.Count > 2 Then
        i = 2
    End If
    
    'If there are any checked rows
    If IsChecked(ReportTable.ListColumns("Select").DataBodyRange) Then
        i = 1
    End If

Footer:
    CheckReport = i

End Function

Function CheckTable(TargetSheet As Worksheet) As Long
'Checks that there is a table, that there is at least one list row, and that there is at least one row checked
'Report sheet will need an additional check since there are two rows at the top
'1 -> Table, rows, checks
'2 -> Table, rows
'3 -> Table
'4 -> None
'Return an null value if there's an error

    Dim TargetCheckRange As Range
    Dim i As Long
    Dim j As Long
    Dim TargetTable As ListObject
    Dim TableHasCheck As Boolean
    
    'Is there a table
    If TargetSheet.ListObjects.Count < 1 Then
        i = 4
        GoTo Footer
    End If
    
    If TargetSheet.Name = "Report Page" Then 'Two rows at the top for the ReportSheet
        Err.Raise vbObjectError + 513, , "Wrong function"
        j = 2
    Else
        j = 1
    End If
    
    'Are there rows
    Set TargetTable = TargetSheet.ListObjects(1)
    If TargetTable.ListRows.Count < j Then
        i = 3
        GoTo Footer
    End If
    
    'Are there checks
    TableHasCheck = IsChecked(TargetTable.ListColumns("Select").DataBodyRange)
    If TableHasCheck = False Then
        i = 2
        GoTo Footer
    End If
    
    i = 1
    
Footer:
    CheckTable = i

End Function
