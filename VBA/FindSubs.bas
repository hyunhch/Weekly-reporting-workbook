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
    
    If LabelCell Is Nothing Or Not Len(LabelCell.Value) > 0 Then
        GoTo Footer
    End If
    
    Set FindActivityLabel = LabelCell
    
Footer:

End Function

Function FindBlanks(SearchRange As Range) As Range
'Finds empty cells in a column of a larger range, returns as a range
'Returns nothing if there are no blanks

    Dim DelRange As Range
    Dim c As Range
    
    'Build range of blanks
    For Each c In SearchRange.Cells
        If Not Len(c.Value) > 0 Then
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

Function FindChecks(SearchRange As Range, Optional SearchType As String) As Range
'Returns a range that contains an "a", not used on RecordsSheet
'Passing "Absent" returns the range of all blank boxes
'Passing "First" returns only the first found "a"
'Returns nothing on error

    Dim SearchSheet As Worksheet
    Dim NudgedRange As Range
    Dim CheckedRange As Range
    Dim c As Range
    
    If SearchRange Is Nothing Then
        GoTo Footer
    End If
    
    Set SearchSheet = Worksheets(SearchRange.Worksheet.Name)
    Set NudgedRange = NudgeToHeader(SearchSheet, SearchRange, "Select")
    
    For Each c In NudgedRange
        Select Case SearchType
        
            Case "First"
                If c.Value = "a" Then
                    Set CheckedRange = BuildRange(c, CheckedRange)
                    
                    GoTo ReturnRange
                End If
                
            Case "Absent"
                If c.Value <> "a" Then
                    Set CheckedRange = BuildRange(c, CheckedRange)
                End If
                
            Case Else
                If c.Value = "a" Then
                    Set CheckedRange = BuildRange(c, CheckedRange)
                End If
        End Select
    Next c
    
    If CheckedRange Is Nothing Then
        GoTo Footer
    End If

ReturnRange:
    Set FindChecks = CheckedRange
    
Footer:
    
End Function

Function FindDuplicate(SourceRange As Range) As Range
'Intermediate function that determines OS because MacOS doesn't support dictionaries
'Returns the range of all duplicates in the range
'Returns nothing if no duplicates are found

    Dim SourceSheet As Worksheet
    Dim SearchRange As Range
    
    'Nudge range to names
    Set SourceSheet = Worksheets(SourceRange.Worksheet.Name)
    
    If Not SourceSheet.Name = "Records Page" Then 'Not on the RecordsSheet, there's no table
        Set SearchRange = NudgeToHeader(SourceSheet, SourceRange, "First")
    Else
        Set SearchRange = SourceRange
    End If

    'Detect OS
    If Application.OperatingSystem Like "*Mac*" Then
        Set FindDuplicate = FindDuplicateMac(SearchRange)
    Else
        Set FindDuplicate = FindDuplicateWin(SearchRange)
    End If

Footer:

End Function

Function FindDuplicateMac(SourceRange As Range) As Range
'Returns the range of all duplicates in the range
'Returns nothing if no duplicates are found

    Dim DuplicateRange As Range
    Dim CompareRange As Range
    Dim LCell As Range
    Dim c As Range
    Dim d As Range
    Dim i As Long
    Dim j As Long
    Dim SourceString As String
    Dim CompareString As String

    'Return nothing if it's a single cell
    If Not SourceRange.Cells.Count > 1 Then
        GoTo Footer
    End If

    Set LCell = SourceRange.Rows(SourceRange.Rows.Count)
    
    'Loop through and build range of duplicates
    i = 1
    For Each c In SourceRange.Resize(SourceRange.Cells.Count - 1, 1).Cells 'Stop at the 2nd to last cell
        Set CompareRange = Range(c.Offset(1, 0), LCell)
        SourceString = c.Value & " " & c.Offset(0, 1).Value
        
        For Each d In CompareRange
            CompareString = d.Value & " " & d.Offset(0, 1).Value
            
            If SourceString = CompareString Then
                Set DuplicateRange = BuildRange(d, DuplicateRange)
            End If
        Next d
    Next c

    'Return
    If Not DuplicateRange Is Nothing Then
        Set FindDuplicateMac = DuplicateRange
    End If

Footer:

End Function

Function FindDuplicateWin(SourceRange As Range) As Range
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
        Set FindDuplicateWin = DuplicateRange
    End If

Footer:

End Function

Function FindLastRow(TargetSheet As Worksheet, Optional TargetHeader As String) As Range
'Returns the a cell in the last used row
'Returns the "Select" column by default, the specified column if a string is passed
'Returns nothing on error

    Dim c As Range
    Dim d As Range
    Dim HeaderString As String
    Dim TargetTable As ListObject

    Set TargetTable = TargetSheet.ListObjects(1)
    
    'If a header was passed
    If Len(TargetHeader) > 0 Then
        HeaderString = TargetHeader
    Else
        HeaderString = "Select"
    End If
    
    Set c = TargetTable.ListColumns(HeaderString).Range
    Set d = TargetSheet.Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    
    If c Is Nothing Then
        GoTo Footer
    ElseIf d Is Nothing Then
        GoTo Footer
    End If
    
    Set FindLastRow = TargetSheet.Cells(d.Row, c.Column)

Footer:
    
End Function

Function FindName(SourceRange As Range, TargetRange As Range) As Range
'Intermediate function that determines OS because MacOS doesn't support dictionaries
'Returns a range of all matching names in the TargetRange
'Returns nothing if no matches found

    Dim SourceSheet As Worksheet
    Dim SearchRange As Range
    
    'Nudge range to names
    Set SourceSheet = Worksheets(SourceRange.Worksheet.Name)
    
    If Not SourceSheet.Name = "Records Page" Then 'Not on the RecordsSheet, there's no table
        Set SearchRange = NudgeToHeader(SourceSheet, SourceRange, "First")
    Else
        Set SearchRange = SourceRange
    End If

    'Detect OS
    If Application.OperatingSystem Like "*Mac*" Then
        Set FindName = FindNameMac(SearchRange, TargetRange)
    Else
        Set FindName = FindNameWin(SearchRange, TargetRange)
    End If

Footer:

End Function

Function FindNameMac(SourceRange As Range, TargetRange As Range) As Range
'Returns a range of all matching names in the TargetRange
'Returns nothing if no matches found
'MacOS doesn't support dictionaries

    Dim MatchRange As Range
    Dim c As Range
    Dim d As Range
    Dim SourceName As String
    Dim TargetName As String
    
    'Loop through the SourceRange, only looking for a single match
    For Each c In SourceRange
        SourceName = c.Value & " " & c.Offset(0, 1).Value
        
        For Each d In TargetRange
            TargetName = d.Value & " " & d.Offset(0, 1).Value
        
            If SourceName = TargetName Then
                Set MatchRange = BuildRange(d, MatchRange)
                
                GoTo NextName
            End If
        Next d
NextName:
    Next c

    'Return
    Set FindNameMac = MatchRange

Footer:

End Function

Function FindNameWin(SourceRange As Range, TargetRange As Range) As Range
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
        Set FindNameWin = MatchRange
    End If
Footer:

End Function

Function FindRecordsActivityHeaders(RecordsSheet As Worksheet, Optional LabelCell As Range, Optional OperationString As String) As Range
'Finds the vertical headers for activities on the Records Page
'If LabelCell is passed, returns the headers for that activity
'If "All" is passed, return the headers for ALL activities

    Dim VCell As Range
    Dim HeaderRange As Range
    Dim LabelRange As Range
    Dim c As Range
    Dim d As Range
    Dim i As Long
    
    i = CheckRecords(RecordsSheet)
    
    'Find the padding cell
    Set VCell = RecordsSheet.Range("1:1").Find("V BREAK", , xlValues, xlWhole)
        If VCell Is Nothing Then
            GoTo Footer
        End If
    
    'Column of headers is one column to the left of VCELL
    Set c = VCell.Offset(0, -1)
    Set d = c.EntireColumn.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
        If d Is Nothing Then
            GoTo Footer
        End If

    Set HeaderRange = RecordsSheet.Range(c, d)
    
    'Either grab the entire range of headers or one specific one
    Set LabelRange = FindRecordsLabel(RecordsSheet, LabelCell)
        If LabelRange Is Nothing Then
            GoTo Footer
        End If
    
    'Return
    If OperationString = "All" Then
        Set FindRecordsActivityHeaders = LabelRange.Resize(HeaderRange.Rows.Count, LabelRange.Columns.Count)
    ElseIf LabelCell Is Nothing Then
        Set FindRecordsActivityHeaders = HeaderRange
    Else
        Set FindRecordsActivityHeaders = LabelRange.Resize(HeaderRange.Rows.Count, 1)
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

Function FindRecordsLabel(RecordsSheet As Worksheet, Optional LabelCell As Range, Optional LabelString As String) As Range
'Returns the range of all activity labels
'If there are no activities, returns the "V BREAK" padding cell
'Returns the cell containing the label if LabelCell or LabelString is passed
'Returns nothing if LabelCell is passed and a match not found

    Dim FCell As Range
    Dim LCell As Range
    Dim LabelRange As Range
    Dim SearchString As String
    
    'Define the range of labels
    Set FCell = RecordsSheet.Range("1:1").Find("V BREAK", , xlValues, xlWhole)
    Set LCell = RecordsSheet.Range("1:1").Find("*", SearchOrder:=xlByColumns, SearchDirection:=xlPrevious)
    
    'If no activities
    If LCell.Value = "V BREAK" Then
        Set FindRecordsLabel = FCell
        GoTo Footer
    End If
    
    Set LabelRange = RecordsSheet.Range(FCell.Offset(0, 1), LCell)

    'If a label is passed
    If Not LabelCell Is Nothing Then
        SearchString = LabelCell.Value
    ElseIf Len(LabelString) > 0 Then
        SearchString = LabelString
    End If
    
    If Len(SearchString) > 0 Then
        Set FCell = LabelRange.Find(SearchString, , xlValues, xlWhole)
        
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

Function FindRecordsRange(RecordsSheet As Worksheet, Optional OperationString As String) As Range
'Returns the used range on the Records sheet, i.e. "A1" to the bottom most row and right most column
'Passing "Names" returns the range of names out to the right most column
'Passing "Labels" returns the range of labels to the bottom most row
'Only returns columns for first and last names if no activities
'Only returns activity headers if no students
'Returns nothing on error or missing both activities and students

    Dim FCell As Range
    Dim LCell As Range
    Dim c As Range
    Dim LRow As Long
    Dim LCol As Long
    Dim i As Long
    Dim ReturnRange As Range
    
    Dim RecordsNameRange As Range
    Dim RecordsLabelRange As Range
    
    'Make sure there are students and activities
    i = CheckRecords(RecordsSheet)
        If i = 4 Then
            GoTo Footer
        End If
        
    'Define some ranges that we may or may not use
    Set RecordsNameRange = FindRecordsName(RecordsSheet)
    Set RecordsLabelRange = FindRecordsActivityHeaders(RecordsSheet, , "All")
   
    'Where to stop
    LRow = RecordsSheet.Range("A:A").Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    LCol = RecordsSheet.Range("1:1").Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Column
    
    'Where to start
    Select Case i
        
        Case 1
            If OperationString = "Names" Then
                Set FCell = RecordsNameRange
                Set LCell = RecordsNameRange.Offset(0, LCol - 1)
                
            ElseIf OperationString = "Labels" Then
                Set FCell = RecordsLabelRange
                Set c = FCell.Resize(1, FCell.Columns.Count) 'Reduce to one row
                Set LCell = c.Offset(LRow - 1, 0)
                
            ElseIf Not Len(OperationString) > 0 Then
                Set FCell = RecordsSheet.Range("A1")
                Set LCell = RecordsSheet.Cells(LRow, LCol)
                
            Else
                Debug.Print "Invalid OperationString"
                GoTo Footer
                
            End If
            
        Case 2 'No students
            If OperationString = "Names" Then
                GoTo Footer
                                
            ElseIf OperationString = "Labels" Then
                Set FCell = RecordsLabelRange
                
            ElseIf Not Len(OperationString) > 0 Then
                Set FCell = RecordsLabelRange
                
            Else
                Debug.Print "Invalid OperationString"
                GoTo Footer
                
            End If
            
        Case 3 'No activities
            If OperationString = "Names" Then
                Set FCell = RecordsNameRange
                Set LCell = RecordsNameRange.Offset(0, 1)
                                
            ElseIf OperationString = "Labels" Then
                GoTo Footer
                
            ElseIf Not Len(OperationString) > 0 Then
                Set FCell = RecordsNameRange
                Set LCell = RecordsNameRange.Offset(0, 1)
                
            Else
                Debug.Print "Invalid OperationString"
                GoTo Footer
                
            End If
    
    End Select
                
    'Return
    If LCell Is Nothing Then
        Set LCell = FCell
    End If
    
    Set ReturnRange = RecordsSheet.Range(FCell, LCell)
    
    If Not ReturnRange Is Nothing Then
        Set FindRecordsRange = ReturnRange
    End If
                
Footer:

End Function

Sub rangetest()

    Dim RecordsSheet As Worksheet
    Dim rng As Range
    Dim str As String
    
    Set RecordsSheet = Worksheets("Records Page")
    'str = "Names"
    'str = "Labels"
    
    Set rng = FindRecordsRange(RecordsSheet, str)
    
    If Not rng Is Nothing Then
        Debug.Print rng.Address
    Else
        Debug.Print "Fail"
    End If

End Sub

Function FindReportLabel(ReportSheet As Worksheet, Optional LabelString As String) As Range
'Returns the cell containing the passed label
'Returns the "Total" row if there are no activities
'Returns the range of all labels if a string isn't passed
'Returns nothing if LabelCell is passed and a match not found

    Dim LabelRange As Range
    Dim MatchCell As Range
    Dim i As Long
    Dim ReportTable As ListObject

    'Make sure there is a table on the page
    i = CheckReport(ReportSheet)
    
    If i > 3 Then
        Call MakeReportTable
    End If
    
    Set ReportTable = ReportSheet.ListObjects(1)
    
    'If there is only the Totals row, return that
    If i > 2 Then
        Set MatchCell = ReportTable.ListColumns("Label").DataBodyRange(1, 1) 'Should only be one cell
        
        If MatchCell.Cells.Count > 1 Then
            GoTo Footer
        End If
        
        Set FindReportLabel = MatchCell
        GoTo Footer
    End If

    'If no string is passed, return the entire DataBodyRange, except the first row
    If Not Len(LabelString) > 0 Then
        Set LabelRange = ReportTable.ListColumns("Label").DataBodyRange
        Set FindReportLabel = LabelRange.Offset(1, 0).Resize(LabelRange.Rows.Count - 1, 1)
        
        GoTo Footer
    End If

    'If a string is passed, find return. Return nothing if not found
    Set LabelRange = ReportTable.ListColumns("Label").DataBodyRange.Find(LabelString, , xlValues, xlWhole)
        If LabelRange Is Nothing Then
            GoTo Footer
        End If
    
    'If found
    Set FindReportLabel = LabelRange

Footer:

End Function

Function FindSheet(SearchString As String, Optional SearchBook As Workbook) As Worksheet
'Returns the activity sheet with the passed label
'Will search in the passed workbook

    Dim TargetBook As Workbook
    Dim TargetSheet As Worksheet
    
    'If a book isn't passed, use this one
    If SearchBook Is Nothing Then
        Set TargetBook = ThisWorkbook
    Else
        Set TargetBook = SearchBook
    End If
    
ActivitySheet:
    For Each TargetSheet In TargetBook.Sheets
        If TargetSheet.Range("A1").Value = "Practice" And _
        Not TargetSheet.Range("1:1").Find(SearchString, , xlValues, xlWhole) Is Nothing Then
            Set FindSheet = TargetSheet
            
            GoTo Footer
        End If
    Next TargetSheet

Footer:

End Function

Function FindStudentAttendance(RecordsSheet As Worksheet, AttendanceRange As Range, Optional SearchType As String) As Range
'Searches horizontally for the attendance of a student or students, or vertically for an activity or activities
'Returns all the cells on the records sheet within the passed range that contain a "1"
'Returns cells that contain "0" if "Absent" is passed
'Returns cell that contain "1" or "0" if "Both" is passed
'Returns cell that are blank if "Blank" is passed
'Returns the first "1" if "First" is passed
'Returns nothing on error

    Dim CheckedRange As Range
    Dim c As Range
    
    If AttendanceRange Is Nothing Then
        GoTo Footer
    End If
    
    
    For Each c In AttendanceRange
        Select Case SearchType
        
            Case "First"
                If c.Value = "1" Then
                    Set CheckedRange = BuildRange(c, CheckedRange)
                    
                    GoTo ReturnRange
                End If
        
            Case "Blank"
                If Not Len(c.Value) > 0 Then
                    Set CheckedRange = BuildRange(c, CheckedRange)
                End If
                
            Case "Absent"
                If c.Value = "0" Then
                    Set CheckedRange = BuildRange(c, CheckedRange)
                End If
                
            Case "Both"
                If c.Value = "0" Or c.Value = "1" Then
                    Set CheckedRange = BuildRange(c, CheckedRange)
                End If
                
            Case Else
                If c.Value = "1" Then
                    Set CheckedRange = BuildRange(c, CheckedRange)
                End If
            
        End Select
    Next c
    
    If CheckedRange Is Nothing Then
        GoTo Footer
    End If
    
ReturnRange:
    Set FindStudentAttendance = CheckedRange

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
        If FCell Is Nothing Then
            GoTo Footer
        End If
    
    LCol = FCell.EntireRow.Find("*", SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
    'Which column to search can change so search all cells
    LRow = TargetSheet.Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    
    Set LCell = TargetSheet.Cells(LRow, LCol)
    Set FindTableRange = TargetSheet.Range(FCell, LCell)

Footer:

End Function

Function FindTableRow(TargetSheet As Worksheet, TargetCell As Range) As Range
'Takes one or more cells in a row and returns the entire table row
'Returns nothing on error
'Can only take a single row in passed range

    Dim TargetTableRange As Range
    Dim c As Range
    Dim TargetTable As ListObject
    
    If Not TargetSheet.ListObjects.Count > 0 Then
        GoTo Footer
    End If
    
    Set TargetTable = TargetSheet.ListObjects(1)
    Set TargetTableRange = TargetTable.DataBodyRange
    Set c = Intersect(TargetCell, TargetTableRange)

    'Check that the passed range only has one row, and that it's contained within the table
    If TargetCell.Rows.Count > 1 Then
        GoTo Footer
    ElseIf c Is Nothing Then
        GoTo Footer
    ElseIf Not c.Address = TargetCell.Address Then
        GoTo Footer
    End If
    
    'Find the intersect
    Set c = Intersect(TargetCell.EntireRow, TargetTableRange)
        
    Set FindTableRow = c
    
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
    Set AttendanceRange = FindRecordsAttendance(RecordsSheet, , LabelCell)
        If AttendanceRange Is Nothing Then
            GoTo Footer
        End If
    
    Set c = FindStudentAttendance(RecordsSheet, AttendanceRange)
        If c Is Nothing Then
            GoTo Footer
        End If
        
    'Scoot the range over to grab the names
    Set PresentRange = c.Offset(0, -c.Column + 1)
        
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
'Intermediate function that determines OS because MacOS doesn't support dictionaries
'Returns a range of all names in the source range but not the target range
'Returns nothing if no matches found
    
    Dim SourceSheet As Worksheet
    Dim SearchRange As Range
    
    'Nudge range to names
    Set SourceSheet = Worksheets(SourceRange.Worksheet.Name)
    
    If Not SourceSheet.Name = "Records Page" Then 'Not on the RecordsSheet, there's no table
        Set SearchRange = NudgeToHeader(SourceSheet, SourceRange, "First")
    Else
        Set SearchRange = SourceRange
    End If
    
    If Application.OperatingSystem Like "*Mac*" Then
        Set FindUnique = FindUniqueMac(SearchRange, TargetRange)
    Else
        Set FindUnique = FindUniqueWin(SearchRange, TargetRange)
    End If

Footer:

End Function

Function FindUniqueMac(SourceRange As Range, TargetRange As Range) As Range
'Returns a range of all non-matching names. Names in the source range but not the target range
'MacOS doesn't support dictionaries
'Returns nothing if no matches found

    Dim UniqueRange As Range
    Dim c As Range
    Dim d As Range
    Dim SourceName As String
    Dim TargetName As String
    
    'Loop through the SourceRange, only looking for a single match
    For Each c In SourceRange
        SourceName = c.Value & " " & c.Offset(0, 1).Value
        
        For Each d In TargetRange
            TargetName = d.Value & " " & d.Offset(0, 1).Value
        
            If SourceName = TargetName Then
                GoTo NextName
            End If
        Next d
        
        'Unique name
        Set UniqueRange = BuildRange(c, UniqueRange)
NextName:
    Next c

    'Return
    Set FindUniqueMac = UniqueRange
                
Footer:

End Function

Function FindUniqueWin(SourceRange As Range, TargetRange As Range) As Range
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
        Set FindUniqueWin = NoMatchRange
    End If
Footer:

End Function
