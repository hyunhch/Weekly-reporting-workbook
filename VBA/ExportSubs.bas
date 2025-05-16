Attribute VB_Name = "ExportSubs"
Option Explicit

Sub SaveNewBook(OldBook As Workbook, NewBook As Workbook)
'Takes the book passed to it, renames it, and prompts a save

    Dim CoverSheet As Worksheet
    Dim CenterRange As Range
    Dim DateRange As Range
    Dim DateString As String
    Dim LocalPath As String
    Dim FileString As String
    Dim SaveString As String
    
    'Grab the information we need for the file name
    With NewBook
        Set CoverSheet = Worksheets("Cover Page")
        Set CenterRange = CoverSheet.Range("A:A").Find("Center", , xlValues, xlWhole).Offset(0, 1)
        Set DateRange = CoverSheet.Range("A:A").Find("Date", , xlValues, xlWhole).Offset(0, 1)
        DateString = Replace(DateRange.Value, "/", "-")
        
        LocalPath = GetLocalPath(ThisWorkbook.path)
        FileString = CenterRange.Value & " " & DateString & ".xlsm"
                
        'Save dialog for Win and Mac
        If Application.OperatingSystem Like "*Mac*" Then
            SaveString = Application.GetSaveAsFilename(LocalPath & "\" & FileString, "Excel Files (*.xlsm), *.xlsm")
            If SaveString = "False" Then
                .Close SaveChanges:=False
                GoTo Footer
            End If
            .SaveAs FileName:=LocalPath & "/" & FileString, FileFormat:=xlOpenXMLWorkbookMacroEnabled
            .Close SaveChanges:=False
        Else
            SaveString = Application.GetSaveAsFilename(LocalPath & "\" & FileString, "Excel Files (*.xlsm), *.xlsm")
            If SaveString = "False" Then
                .Close SaveChanges:=False
                GoTo Footer
            End If
            .SaveAs FileName:=SaveString, FileFormat:=xlOpenXMLWorkbookMacroEnabled
            .Close SaveChanges:=False
        End If
    End With

Footer:

End Sub

Function MakeNewBook(RecordsSheet As Worksheet, ReportSheet As Worksheet, Optional RosterSheet As Worksheet, Optional NameRange As Range, Optional BookType As String) As Workbook
'Creates a new workbook, creates summarized and detailed attendance reports
'ExportRange can be from the RosterSheet or the RecordsSheet
'Does everything except save the workbook
'Passing a range limits exporting to that range and skips the report
'Passing "SharePoint" only generates the cover and report

    Dim OldBook As Workbook
    Dim NewBook As Workbook
    Dim OldCoverSheet As Worksheet
    Dim NewCoverSheet As Worksheet
    Dim RecordsNameRange As Range
    Dim RosterNameRange As Range
    Dim DetailedExportRange As Range
    Dim SimpleExportRange As Range
    Dim ExportNameRange As Range
    Dim CopyRange As Range
    Dim PasteRange As Range
    Dim c As Range
    Dim i As Long

    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    'Make sure there's something to export. This should have already been done in a parent sub
    i = CheckRecords(RecordsSheet)
    If i = 2 Or i = 4 Then 'No students
        GoTo Footer
    End If
    
    'Grab information on the CoverSheet
    Set OldBook = ThisWorkbook
    Set OldCoverSheet = OldBook.Worksheets("Cover Page")
    Set c = OldCoverSheet.Range("A:A").Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Offset(0, 1) 'May need to change this in the future
    Set CopyRange = OldCoverSheet.Range("A1", c) 'We want everything from A1 to the cell right of "Center"
    
    'Make a new book, rename the sheet, and put in the cover sheet information
    Set NewBook = Workbooks.Add
    NewBook.Activate
    NewBook.Sheets("Sheet1").Name = "Cover Page"
    
    Set NewCoverSheet = NewBook.Worksheets("Cover Page")
    Set PasteRange = NewCoverSheet.Range(CopyRange.Address)
    PasteRange.Value = CopyRange.Value
    
    'Skip exporting attendance if the RosterSheet wasn't passed
    If RosterSheet Is Nothing Then
        GoTo ReturnBook
    End If
     
    'Define who is being exported. Checking that there are students, etc. should be done in a parent sub
    Set RecordsNameRange = FindRecordsName(RecordsSheet)
    Set RosterNameRange = RosterSheet.ListObjects(1).ListColumns("First").DataBodyRange
    
    'Skip matching if no range was passed or if a range of all students was passed
    If NameRange Is Nothing Then
        Set ExportNameRange = RosterNameRange
    Else
        Set ExportNameRange = NameRange
        
        'Create the Report if no name range was passed
        Call ExportReport(ReportSheet, NewBook)
    End If
    
    'Nothing else to be done for SharePoint
    If BookType = "SharePoint" Then
        GoTo ReturnBook
    End If
    
    'Create the simple and detailed  attendance
    Call ExportRoster(RosterSheet, NewBook, NameRange)
    Call ExportSimpleAttendance(RecordsSheet, NewBook, ExportNameRange)
    Call ExportDetailedAttendance(RecordsSheet, RosterSheet, NewBook, ExportNameRange)

ReturnBook:
    Set MakeNewBook = NewBook
     
Footer:

End Function

Sub ExportRoster(RosterSheet As Worksheet, NewBook As Workbook, Optional NameRange As Range)
'Creates a new sheet and puts the entire roster on it
'Only copies over the passed names is a range is passed

    Dim NewRosterSheet As Worksheet
    Dim RosterTableRange As Range
    Dim RosterNameRange As Range
    Dim RosterMatchRange As Range
    Dim RosterExportRange As Range
    Dim CopyRange As Range
    Dim PasteRange As Range
    Dim c As Range
    Dim RosterTable As ListObject

    'Grab the entire table range of the Report sheet. Checking that there is anything there should be done in a parent sub
    Set RosterTable = RosterSheet.ListObjects(1)
    Set RosterTableRange = RosterTable.Range
    
    'Create a new sheet
    With NewBook
        Set NewRosterSheet = .Sheets.Add(After:=.Sheets(.Sheets.Count))
        NewRosterSheet.Name = "Roster Page"
    End With

    'Put in the header
    Set c = RosterTable.HeaderRowRange.Resize(1, RosterTable.ListColumns.Count - 1)
    Set CopyRange = c.Offset(0, 1) 'Chop off the first column
    
    Set c = NewRosterSheet.Range("A1")
    Set PasteRange = c.Resize(CopyRange.Rows.Count, CopyRange.Columns.Count)
    
    PasteRange.Value(11) = CopyRange.Value(11)
    
    'If no range was passed, copy the entire roster
    If NameRange Is Nothing Then
        Set c = RosterTable.DataBodyRange.Resize(RosterTableRange.Rows.Count, RosterTableRange.Columns.Count - 1)
        Set CopyRange = c.Offset(0, 1) 'Chop off the first column
    
        Set c = NewRosterSheet.Range("A2")
        Set PasteRange = c.Resize(CopyRange.Rows.Count, CopyRange.Columns.Count)

        PasteRange.Value(11) = CopyRange.Value(11)
    'Only copy over passed students if a range was passed
    Else
        Set RosterNameRange = RosterTable.ListColumns("First").DataBodyRange
        Set RosterMatchRange = FindName(NameRange, RosterNameRange)
        
        If RosterMatchRange Is Nothing Then
            GoTo Footer
        End If

        For Each c In RosterMatchRange
            'This excludes the first column
            Set RosterExportRange = BuildRange(c.Resize(1, RosterTable.ListColumns.Count - 1), RosterExportRange)
        Next c
        
        Call CopyRows(RosterSheet, RosterExportRange, NewRosterSheet, NewRosterSheet.Range("A2"))
    End If

Footer:

End Sub

Sub ExportReport(ReportSheet As Worksheet, NewBook As Workbook)
'Reproduces the report into the new workbook

    Dim NewReportSheet As Worksheet
    Dim ReportTableRange As Range
    Dim CopyRange As Range
    Dim PasteRange As Range
    Dim c As Range
    Dim d As Range
    
    'Grab the entire table range of the Report sheet. Checking that there is anything there should be done in a parent sub
    Set ReportTableRange = FindTableRange(ReportSheet)
    
    'Create a new sheet
    With NewBook
        Set NewReportSheet = .Sheets.Add(After:=.Sheets(.Sheets.Count))
        NewReportSheet.Name = "Report Page"
    End With
    
    'Chop off the first column for copying and paste at the top of the new sheet
    Set CopyRange = ReportTableRange.Resize(ReportTableRange.Rows.Count, ReportTableRange.Columns.Count - 1).Offset(0, 1)
    Set c = NewReportSheet.Range("A1")
    Set PasteRange = c.Resize(CopyRange.Rows.Count, CopyRange.Columns.Count)
    
    PasteRange.Value(11) = CopyRange.Value(11)
    
    'Fit the columns up to one before the "Total" header for the first two rows
    Set c = NewReportSheet.Range("1:1").Find("Total", , xlValues, xlWhole)
    
    If Not c Is Nothing Then
        Set d = NewReportSheet.Range("A2", c.Offset(0, -1))
        d.Columns.AutoFit
    End If
    
Footer:

End Sub

Sub ExportSimpleAttendance(RecordsSheet As Worksheet, NewBook As Workbook, Optional NameRange As Range)
'Creates new sheet in passed workbook for simple Attendance
'If a range is passed, copies over information for each indicated student

    Dim SimpleSheet As Worksheet
    Dim RecordsNameRange As Range
    Dim ActivityHeaderRange As Range
    Dim NameHeaderRange As Range
    Dim ExportRange As Range
    Dim CopyRange As Range
    Dim PasteRange As Range
    Dim DateRange As Range
    Dim MatchCell As Range
    Dim LCell As Range
    Dim c As Range
    Dim i As Long
    
    'Find when activities end. Checking if there are any should be done in a parent sub
    Set LCell = RecordsSheet.Range("1:1").Find("*", SearchOrder:=xlByColumns, SearchDirection:=xlPrevious)
    
    'Find the name headers
    Set c = RecordsSheet.Range("A:A").Find("H BREAK", , xlValues, xlWhole) 'This is one past the headers
    Set NameHeaderRange = c.Resize(1, 2).Offset(-1, 0)
    
    'Find the activity headers. They're vertical
    Set c = RecordsSheet.Range("1:1").Find("V BREAK", , xlValues, xlWhole) 'This is one to the right of the headers
    Set ActivityHeaderRange = c.Resize(NameHeaderRange.Row - 1, 1).Offset(0, -1) 'It will end one row above the name headers
    
    'Set ActivityHeaderRange = RecordsSheet.Range(ActivityHeaderRange, LCell).Resize(NameHeaderRange.Row - 1, LCell.Column - ActivityHeaderRange.Column + 1)

    'Create new sheet in passed workbook for simple attendance
    With NewBook
        Set SimpleSheet = .Sheets.Add(After:=.Sheets(.Sheets.Count))
        SimpleSheet.Name = "Simple Attendance"
    End With
    
    'Add headers to the same location in the new sheet
    Set CopyRange = NameHeaderRange
    Set PasteRange = SimpleSheet.Range(CopyRange.Address)
    PasteRange.Value = CopyRange.Value
    
    Set CopyRange = ActivityHeaderRange.Resize(NameHeaderRange.Row - 1, LCell.Column - ActivityHeaderRange.Column + 1)
    Set PasteRange = SimpleSheet.Range(CopyRange.Address)
    PasteRange.Value = CopyRange.Value
    
    'Format the date row and autofit
    Set c = ActivityHeaderRange.Find("Date", , xlValues, xlWhole)
    If c Is Nothing Then
        GoTo Footer
    End If
    
    c.EntireRow.NumberFormat = "mm/dd/yyyy"
    
    'Skip copying students if no names were passed
    If NameRange Is Nothing Then
        GoTo Footer
    End If
    
    'Define the row of first names
    Set RecordsNameRange = FindRecordsName(RecordsSheet)
    
    'Copy over simple attendance
    i = 0
    For Each c In NameRange
        Set MatchCell = FindName(c, RecordsNameRange)
        If Not MatchCell Is Nothing Then 'This would only matter if the names didn't match, which shouldn't happen
            Set CopyRange = RecordsSheet.Range(MatchCell, MatchCell.Offset(0, LCell.Column - 1))
            Set PasteRange = SimpleSheet.Range(RecordsNameRange(1).Address, RecordsNameRange(1).Offset(0, LCell.Column - 1).Address).Offset(i, 0)
            PasteRange.Value = CopyRange.Value
            
            i = i + 1
SkipLoop:
        End If
    Next c

Footer:
    'Remove the empty row and column where the padding cells were
    SimpleSheet.Range(NameHeaderRange.Address).Offset(1, 0).EntireRow.Delete
    SimpleSheet.Range(ActivityHeaderRange.Address).Offset(0, 1).EntireColumn.Delete
    
End Sub

Sub ExportDetailedAttendance(RecordsSheet As Worksheet, RosterSheet As Worksheet, NewBook As Workbook, Optional NameRange As Range)
'Creates a new sheet for detailed Attendance information
'Populates with each activity done by the students in the passed range

    Dim DetailedSheet As Worksheet
    Dim CopyNameRange As Range
    Dim PasteNameRange As Range
    Dim CopyActivityRange As Range
    Dim PasteActivityRange As Range
    Dim CopyDemoRange As Range
    Dim PasteDemoRange As Range
    Dim RosterNameRange As Range
    Dim RecordsNameRange As Range
    Dim AttendanceRange As Range
    Dim DateRange As Range
    Dim MatchCell As Range
    Dim FCell As Range
    Dim LCell As Range
    Dim CopyRange As Range
    Dim PasteRange As Range
    Dim c As Range
    Dim d As Range
    Dim i As Long
    Dim j As Long
    Dim RosterTable As ListObject
    
    'Create new sheets in the workbook for simple attendance
    With NewBook
        Set DetailedSheet = .Sheets.Add(After:=.Sheets(.Sheets.Count))
        DetailedSheet.Name = "Detailed Attendance"
    End With
    
    '''Names'''
    'Insert headers in row 1, starting with First and Last name
    DetailedSheet.Range("A1").Value = "First"
    DetailedSheet.Range("B1").Value = "Last"
    
    'Define where to paste names
    Set PasteNameRange = DetailedSheet.Range("A1:B1")
    
    '''Activity information'''
    'The headers for the activity information will need to be transposed. Doing it programmatically in case the number or order of headers change
    Set FCell = RecordsSheet.Range("1:1").Find("V BREAK", , xlValues, xlWhole).Offset(0, -1) 'Headers are one column before
    Set LCell = FCell.EntireColumn.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    Set CopyActivityRange = RecordsSheet.Range(FCell, LCell)
    
    'Define where to paste activity headers and paste
    Set PasteActivityRange = DetailedSheet.Range("C1")
    
    For i = FCell.Row To LCell.Row
        PasteActivityRange.Offset(0, i - 1).Value = FCell.Offset(i - 1, 0).Value
    Next i
    
    '''Demographic and other information'''
    'Grab the rest of the headers on the RosterSheet, including custom ones added
    Set RosterTable = RosterSheet.ListObjects(1)
    Set FCell = RosterTable.HeaderRowRange.Find("Last", , xlValues, xlWhole).Offset(0, 1)
    Set LCell = RosterSheet.Cells(FCell.Row, RosterTable.ListColumns.Count)
    Set CopyDemoRange = RosterSheet.Range(FCell, LCell)
    
    'Define where to paste
    Set FCell = PasteActivityRange.Offset(0, i - 1) 'i will be one larger than the number of activity headers
    Set PasteDemoRange = FCell.Resize(1, CopyDemoRange.Columns.Count)
    
    PasteDemoRange.Value = CopyDemoRange.Value
    
    'If a range of names was passed, fill each column for each time a student was present
    If NameRange Is Nothing Then
        GoTo Footer
    End If
    
    'For each passed name, find the row of Attendance
    j = 1
    For Each c In NameRange
        Set AttendanceRange = FindRecordsAttendance(RecordsSheet, c)
        If AttendanceRange Is Nothing Then 'This shouldn't happen
            GoTo NextName
        End If
        
        'Go across and find each 1
        For Each d In AttendanceRange
            If d.Value = 1 Then
                'Copy name
                Set CopyRange = c.Resize(1, 2)
                Set PasteRange = PasteNameRange.Offset(j, 0)
                PasteRange.Value = CopyRange.Value
                
                'Copy activity information
                Set CopyRange = CopyActivityRange.Offset(0, d.Column - CopyActivityRange.Column)
                Set PasteRange = PasteActivityRange.Offset(j, 0)

                For i = 1 To CopyActivityRange.Rows.Count
                    PasteRange.Offset(0, i - 1).Value = RecordsSheet.Cells(1, CopyRange.Column).Offset(i - 1, 0).Value
                Next i
                
                'Copy demographics
                Set CopyRange = CopyDemoRange.Offset(c.Row - CopyDemoRange.Row, 0)
                Set PasteRange = PasteDemoRange.Offset(j, 0)
                PasteRange.Value = CopyRange.Value
                
                j = j + 1
            End If
        Next d
NextName:
    Next c

    'Autofit names
    PasteNameRange.EntireColumn.AutoFit

    'Format the date column and autofit
    Set c = DetailedSheet.Range("1:1").Find("Date")
    c.NumberFormat = "mm/dd/yyyy"
    c.EntireColumn.AutoFit

Footer:

End Sub
