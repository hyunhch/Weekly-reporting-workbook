Attribute VB_Name = "FindSubs"
Option Explicit

Function FindActivityLabel(ActivitySheet As Worksheet) As Range
'Returns the cell containing the label on an activity
'Returns nothing if it's not found. That shouldn't happen

    Dim LabelCell As Range
    Dim SearchRange As Range
    Dim FCell As Range
    Dim LCell As Range
    Dim LRow As Long
    Dim LCol As Long
    
    On Error GoTo Footer
    
    'Define the search range. The label should always be in the 1st row, but doing this programmatically in case that changes
    Set FCell = ActivitySheet.Range("A1")
    LRow = FCell.EntireColumn.Find("Select", , xlValues, xlWhole).Offset(-1, 0).Row 'One row above the table
    LCol = FCell.EntireRow.Find("*", SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
    Set LCell = ActivitySheet.Cells(LRow, LCol)
    Set SearchRange = ActivitySheet.Range(FCell, LCell)
    
    'Search and return
    Set LabelCell = SearchRange.Find("Label", , xlValues, xlWhole).Offset(0, 1)
    
    If Not LabelCell Is Nothing Then
        Set FindActivityLabel = LabelCell
    End If
    
Footer:

End Function

Function FindBlanks(TargetSheet As Worksheet, SearchRange As Range) As Range
'Finds empty cells in a column of a larger range, returns as a range
'Returns nothing if there are no blanks

    Dim DelRange As Range
    Dim c As Range
    
    'Build range of blanks
    For Each c In SearchRange.Cells
        If c.Value = "" Then
            Set DelRange = BuildRange(c, DelRange)
        End If
    Next c
    
    'If there are no blanks
    If DelRange Is Nothing Then
        GoTo Footer
    End If

    'REturn
    Set FindBlanks = DelRange
    
Footer:

End Function

Function FindChecks(TargetRange As Range, Optional SearchType As String) As Range
'Returns a range that contains all cells that are not empty or 0
'Passing "Absent" returns absent students

    Dim CheckedRange As Range
    Dim c As Range
    
    If SearchType <> "Absent" Then
        GoTo OnlyPresent
    End If
    
    For Each c In TargetRange
        If c.Value = "0" Then  'Only get absenses
            Set CheckedRange = BuildRange(c, CheckedRange)
        End If
    Next c
    
    GoTo SetRange
    
OnlyPresent:
    For Each c In TargetRange
        If c.Value <> "" And c.Value <> "0" Then 'Ignore empty spaces and absenses on the Records sheet
            Set CheckedRange = BuildRange(c, CheckedRange)
        End If
    Next c
    
SetRange:
    Set FindChecks = CheckedRange
    
Footer:
    
End Function

Function FindDuplicate(SourceRange As Range) As Range
'Returns the range of all duplicates in the range
'Returns nothing if no duplicates are found

    Dim DuplicateRange As Range
    Dim c As Range
    Dim NameString As String
    Dim NameDict As Object
    
    Set NameDict = CreateObject("Scripting.Dictionary")

    'Loop through passed range, read into dictionary
    For Each c In SourceRange
        If Len(c.Value) < 1 Then
            GoTo NextName
        End If
    
        NameString = c.Value & " " & c.Offset(0, 1).Value
        If Not NameDict.Exists(NameString) Then
            NameDict.Add NameString, c
        Else
            Set DuplicateRange = BuildRange(c, DuplicateRange)
        End If
NextName:
    Next c

    'Return
    If Not DuplicateRange Is Nothing Then
        Set FindDuplicate = DuplicateRange
    End If

Footer:

End Function

Function FindName(SourceRange As Range, TargetRange As Range) As Range
'Returns a range of all matching names in the TargetRange
'Returns nothing if no matches found

    Dim MatchRange As Range
    Dim MatchCell As Range
    Dim c As Range
    Dim d As Range
    Dim NameString As String
    Dim NameDict As Object
    
    Set NameDict = CreateObject("Scripting.Dictionary")

    'Loop through source range, read all unique names into dictionary
    For Each c In SourceRange
        If Len(c.Value) < 1 Then
            GoTo NextName1
        End If
        
        NameString = c.Value & " " & c.Offset(0, 1).Value
        If Not NameDict.Exists(NameString) Then
            NameDict.Add NameString, c
        End If
NextName1:
    Next c

    'Loop through target range, find any matches
    For Each d In TargetRange
        If Len(d.Value) < 1 Then
            GoTo NextName2
        End If
    
        NameString = d.Value & " " & d.Offset(0, 1).Value
        If NameDict.Exists(NameString) Then
            Set MatchCell = d
            Set MatchRange = BuildRange(MatchCell, MatchRange)
            
            If SourceRange.Cells.Count = 1 Then 'So we don't loop the entire list of names if we are only looking for one
                GoTo ReturnRange
            End If
        End If
NextName2:
    Next d
    
ReturnRange:
    If Not MatchRange Is Nothing Then
        Set FindName = MatchRange
    End If
Footer:

End Function

Function FindPresent(RecordsSheet As Worksheet, LabelCell As Range, Optional OperationString As String) As Range
'Returns the range of all present students given the passed cell
'Returns nothing if there are no students recorded as present, or if the activity isn't found
'Returns absent students if "Absent" is passed
'Returns both absent and present if "All" is passed

    Dim RecordsNameRange As Range
    Dim RecordsAttendanceRange As Range
    Dim c As Range
    Dim d As Range
    Dim e As Range
    Dim IsPresent As Boolean
    Dim IsAbsent As Boolean
    
    'Make sure there are both students and activities
    If CheckRecords(RecordsSheet) <> 1 Then
        GoTo Footer
    End If
    
    'Find the vertical range containing attendance information
    Set RecordsAttendanceRange = FindRecordsAttendance(RecordsSheet, , LabelCell)
    
    If RecordsAttendanceRange Is Nothing Then
        GoTo Footer
    End If

    'Check that there are students to return
    IsPresent = IsChecked(RecordsAttendanceRange)
    IsAbsent = IsChecked(RecordsAttendanceRange, "Absent")
    
    'No student attendance
    If IsPresent = False And IsAbsent = False Then 'This checks the contents of the range, not if the range exists
        GoTo Footer
    'No absent students
    ElseIf OperationString = "Absent" And IsAbsent = False Then
        GoTo Footer
    'No present students
    ElseIf Len(OperationString) < 1 And IsPresent = False Then
        GoTo Footer
    End If
    
    'Define the range of names and grab all that were present/absent
    Set RecordsNameRange = FindRecordsName(RecordsSheet) 'Should always be in the A column, but making it programmatic
    Set c = FindChecks(RecordsAttendanceRange)
    Set d = FindChecks(RecordsAttendanceRange, "Absent")
    
    'Return
    If Len(OperationString) < 1 Then
        Set FindPresent = c.Offset(0, -RecordsAttendanceRange.Column + RecordsNameRange.Column)
    ElseIf OperationString = "Absent" Then
        Set FindPresent = d.Offset(0, -RecordsAttendanceRange.Column + RecordsNameRange.Column)
    ElseIf OperationString = "All" Then
        Set e = Union(c, d)
        Set FindPresent = e.Offset(0, -RecordsAttendanceRange.Column + RecordsNameRange.Column)
    End If
    
Footer:

End Function

Function FindRecordsActivityHeaders(RecordsSheet As Worksheet, Optional LabelCell As Range) As Range
'Finds the vertical headers for activities on the Records Page
'If LabelCell is passed, returns the headers for that activity

    Dim VCell As Range
    Dim HCell As Range
    Dim c As Range
    Dim d As Range
    
    'Find the padding cells
    Set VCell = RecordsSheet.Range("1:1").Find("V BREAK", , xlValues, xlWhole)
    Set HCell = RecordsSheet.Range("A:A").Find("H BREAK", , xlValues, xlWhole)
    
    'Headers are one cell left of the VCell, two cells up from the HCell
    Set c = VCell.Offset(0, -1).Resize(HCell.Row - 2, 1)
    
    'If a label was passed
    If Not LabelCell Is Nothing Then
        Set d = FindRecordsLabel(RecordsSheet, LabelCell)
        If d Is Nothing Then
            GoTo Footer
        End If
        
        Set FindRecordsActivityHeaders = d.Resize(c.Rows.Count, 1)
    'If nothing was passed
    Else
        Set FindRecordsActivityHeaders = c
    End If
    
Footer:

End Function

Function FindRecordsAttendance(RecordsSheet As Worksheet, Optional NameCell As Range, Optional LabelCell As Range) As Range
'Returns the intersection of all rows containing students and all columns containing activities
'Passing a cell with a name will return the attendance for just that student
'Passing a cell with a label will return the Attendance for that activity
'Returns nothing if there are either no students or no activities

    Dim RecordsNameRange As Range
    Dim RecordsLabelRange As Range
    Dim IntersectRange As Range
    Dim MatchCell As Range
    Dim c As Range
    
    'Make sure there are students and activites
    If CheckRecords(RecordsSheet) <> 1 Then
        GoTo Footer
    End If
    
    'Define ranges to search
    Set RecordsNameRange = FindRecordsName(RecordsSheet)
    Set RecordsLabelRange = FindRecordsLabel(RecordsSheet, LabelCell)
    
    'If a name was passed
    If NameCell Is Nothing Then
        GoTo LabelCheck
    End If
    
    Set MatchCell = FindRecordsName(RecordsSheet, NameCell)
    Set FindRecordsAttendance = RecordsLabelRange.Offset(MatchCell.Row - 1, 0)
    GoTo Footer

LabelCheck:
    'If a label is passed
    If LabelCell Is Nothing Then
        GoTo AllCheck
    End If
    
    Set MatchCell = FindRecordsLabel(RecordsSheet, LabelCell)
    
    If MatchCell Is Nothing Then
        GoTo Footer
    End If
    
    Set FindRecordsAttendance = RecordsNameRange.Offset(0, MatchCell.Column - 1)
    GoTo Footer
    
AllCheck:
    'Return the entire range of Attendance
    Set IntersectRange = Intersect(RecordsNameRange.EntireRow, RecordsLabelRange.EntireColumn)
    
    If Not IntersectRange Is Nothing Then
        Set FindRecordsAttendance = IntersectRange
    End If
    
Footer:

End Function

Function FindRecordsLabel(RecordsSheet As Worksheet, Optional LabelCell As Range) As Range
'Returns the range of all activity labels
'If there are no activities, returns the "V BREAK" padding cell
'Returns the cell containing the label if LabelCell is passed
'Returns nothing if LabelCell is passed and a match not found

    Dim FCell As Range
    Dim LCell As Range
    Dim LabelRange As Range
    
    'Define the range of labels
    Set FCell = RecordsSheet.Range("1:1").Find("V BREAK", , xlValues, xlWhole)
    Set LCell = RecordsSheet.Range("1:1").Find("*", SearchOrder:=xlByColumns, SearchDirection:=xlPrevious)
    
    'If no activities
    If LCell.Value = "V BREAK" Then
        Set FindRecordsLabel = FCell
        GoTo Footer
    End If
    
    Set LabelRange = RecordsSheet.Range(FCell.Offset(0, 1), LCell)
    
    'If a name is passed
    If Not LabelCell Is Nothing Then
        Set FCell = LabelRange.Find(LabelCell.Value, , xlValues, xlWhole)
        If Not FCell Is Nothing Then
            Set FindRecordsLabel = FCell
            GoTo Footer
        'If a match isn't found
        Else
            GoTo Footer
        End If
    End If
    
    'Entire range
    Set FindRecordsLabel = LabelRange

Footer:

End Function

Function FindRecordsName(RecordsSheet As Worksheet, Optional NameCell As Range)
'Returns the entire range of names if nothing passed
'Returns the "H BREAK" padding cell if there are no names
'Returns cell with the student's first name if NameCell is passed
'Returns nothing if a range is passed and a match not found

    Dim FCell As Range
    Dim LCell As Range
    Dim MatchCell As Range
    Dim NameRange As Range
    
    'Define the range of names
    Set FCell = RecordsSheet.Range("A:A").Find("H BREAK", , xlValues, xlWhole)
    Set LCell = RecordsSheet.Range("A:A").Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious)

    'If there are no names
    If LCell.Value = "H BREAK" Then
        Set FindRecordsName = LCell
        GoTo Footer
    End If
    
    Set NameRange = RecordsSheet.Range(FCell.Offset(1, 0), LCell)
    
    'If a name is passed
    If Not NameCell Is Nothing Then
        Set MatchCell = FindName(NameCell, NameRange)
        If Not MatchCell Is Nothing Then
            Set FindRecordsName = MatchCell
            GoTo Footer
        'If a match isn't found
        Else
            GoTo Footer
        End If
    End If
    
    'Entire range
    Set FindRecordsName = NameRange
        
Footer:

End Function

Function FindReportLabel(ReportSheet As Worksheet, Optional LabelCell As Range) As Range
'Returns the cell containing the passed label
'Returns the "Total" row if there are no activities
'Returns the range of all labels if a string isn't passed
'Returns nothing if LabelCell is passed and a match not found

    Dim LabelColumn As Range
    Dim MatchCell As Range
    Dim ReportTable As ListObject

    'Make sure there is a table on the page
    If Not ReportSheet.ListObjects.Count > 0 Then
        Call CreateReportTable
    End If
    
    'If there are no activities, there will only be the Totals row
    Set ReportTable = ReportSheet.ListObjects(1)
    
    If ReportTable.ListRows.Count = 1 Then
        Set FindReportLabel = ReportTable.ListColumns("Label").DataBodyRange 'In case the placement of the Total row changes in the future
        GoTo Footer
    End If

    'If no string is passed
    If LabelCell Is Nothing Then
        Set MatchCell = ReportTable.ListColumns("Label").DataBodyRange 'This will always be at least 1
        Set FindReportLabel = MatchCell.Offset(1, 0).Resize(MatchCell.Rows.Count - 1, 1)
        GoTo Footer
    End If
    
    'If a string is passed
    Set LabelColumn = ReportTable.ListColumns("Label").DataBodyRange
    Set MatchCell = LabelColumn.Find(LabelCell.Value, , xlValues, xlWhole)

    'If it's not found, return nothing
    If MatchCell Is Nothing Then
        GoTo Footer
    End If
    
    'If found
    Set FindReportLabel = MatchCell

Footer:

End Function

Function FindSheet(SearchString As String, Optional SearchBook As Workbook) As Worksheet
'Returns the activity sheet with the passed label
'Will search in the passed workbook

    Dim TargetSheet As Worksheet
    
ActivitySheet:
    For Each TargetSheet In ThisWorkbook.Sheets
        If TargetSheet.Range("A1").Value = "Practice" And _
        Not TargetSheet.Range("1:1").Find(SearchString, , xlValues, xlWhole) Is Nothing Then
            Set FindSheet = TargetSheet
            GoTo Footer
        End If
    Next TargetSheet

Footer:

End Function

Function FindTableHeader(TargetSheet As Worksheet, StartString As String, Optional EndString As String) As Range
'Returns the cell containing the passed string in a table's header
'Returns all header cells between two strings if a second one is passed
'Returns nothing if the header isn't found

    Dim StartCell As Range
    Dim EndCell As Range
    Dim TargetTable As ListObject
    
    'Make sure there's a table
    If TargetSheet.ListObjects.Count < 1 Then
        GoTo Footer
    End If

    Set TargetTable = TargetSheet.ListObjects(1)
    Set StartCell = TargetTable.HeaderRowRange.Find(StartString, , xlValues, xlWhole)

    If StartCell Is Nothing Then
        GoTo Footer
    End If

    If Not Len(EndString) > 0 Then
        'Return one cell
        Set FindTableHeader = StartCell
    Else
        'Return a range
        Set EndCell = TargetTable.HeaderRowRange.Find(EndString, , xlValues, xlWhole)
        Set FindTableHeader = TargetSheet.Range(StartCell, EndCell)
    End If
    
Footer:

End Function

Function FindTableRange(TargetSheet As Worksheet) As Range
'Returns the range that will be used to create a table
'Returns empty if there's an error

    Dim FCell As Range
    Dim LCell As Range
    Dim LRow As Long
    Dim LCol As Long
    
    On Error GoTo Footer
    
    'All tables used will have a cell with "Select" in the 1st column
    Set FCell = TargetSheet.Range("A:A").Find("Select", , xlValues, xlWhole)
    LCol = FCell.EntireRow.Find("*", SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
    
    'Which column to search can change so search all cells
    LRow = TargetSheet.Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    Set LCell = TargetSheet.Cells(LRow, LCol)
    Set FindTableRange = TargetSheet.Range(FCell, LCell)

Footer:

End Function

Function FindTabulateRange(RosterSheet As Worksheet, RecordsSheet As Worksheet, LabelCell As Range) As Range
'Finds all students marked present on the Records Sheet
'Returns where the same students are on the Roster sheet for tabulation

    Dim RosterNameRange As Range
    Dim AttendanceRange As Range
    Dim PresentRange As Range
    Dim TabulateRange As Range
    Dim c As Range
     
    'Find the activity and identify students who were present
    Set PresentRange = FindPresent(RecordsSheet, LabelCell)
    
    'Handled by a function now
    'Set AttendanceRange = FindRecordsAttendance(RecordsSheet, , LabelCell)
    
    'For Each c In AttendanceRange
        'If c.Value = "1" Then
            'Set PresentRange = BuildRange(c.Offset(0, -c.Column + 1), PresentRange)
            'If Not PresentRange Is Nothing Then
                'Set PresentRange = Union(PresentRange, RecordsSheet.Cells(c.Row, 1))
            'Else
                'Set PresentRange = RecordsSheet.Cells(c.Row, 1)
            'End If
        'End If
    'Next c
    
    'If nothing was found on the Records sheet
    If PresentRange Is Nothing Then
        GoTo Footer
    End If
    
    'Match the students to the Roster sheet
    Set RosterNameRange = RosterSheet.ListObjects(1).ListColumns("First").DataBodyRange
    Set TabulateRange = FindName(PresentRange, RosterNameRange)
    
    If TabulateRange Is Nothing Then 'This shouldn't happen
        GoTo Footer
    End If
    
    Set FindTabulateRange = TabulateRange
    
Footer:
    
End Function

Function FindUnique(SourceRange As Range, TargetRange As Range) As Range
'Returns a range of all non-matching names. Names in the source range but not the target range
'Returns nothing if no matches found

    Dim NoMatchRange As Range
    Dim c As Range
    Dim d As Range
    Dim NameString As String
    Dim NameDict As Object
    
    Set NameDict = CreateObject("Scripting.Dictionary")

    'Loop through source range, read all unique names into dictionary
    For Each c In TargetRange
        If Len(c.Value) < 1 Then
            GoTo NextTargetName
        End If
        
        NameString = c.Value & " " & c.Offset(0, 1).Value
        If Not NameDict.Exists(NameString) Then
            NameDict.Add NameString, c
        End If
NextTargetName:
    Next c

    'Loop through target range, find those that don't match
    For Each c In SourceRange
        If Len(c.Value) < 1 Then
            GoTo NextSourceName
        End If
    
        NameString = c.Value & " " & c.Offset(0, 1).Value
        If Not NameDict.Exists(NameString) Then
            Set NoMatchRange = BuildRange(c, NoMatchRange)
        End If
NextSourceName:
    Next c
    
    'Return
    If Not NoMatchRange Is Nothing Then
        Set FindUnique = NoMatchRange
    End If
Footer:

End Function
