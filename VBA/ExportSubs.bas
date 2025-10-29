Attribute VB_Name = "ExportSubs"
Option Explicit

Sub extest()

    Dim OldBook As Workbook
    Dim NewBook As Workbook
    Dim ws As Worksheet
    Dim rng As Range
    Dim i As Long
    Dim arr As Variant
    
    Set OldBook = ThisWorkbook
    Set ws = Worksheets("Roster Page")
    
    ReDim arr(1 To 5)
        arr(1) = "Cover Page"
        arr(2) = "Report Page"
        arr(3) = "Roster Page"
        arr(4) = "Simple Attendance"
        arr(5) = "Detailed Attendance"
        
    'Set NewBook = ExportMakeBook(, arr)
    
    Set rng = FindChecks(ws.ListObjects(1).ListColumns("Select").DataBodyRange)
    Set NewBook = ExportMakeBook(rng, arr)

End Sub

Function ExportFromRecords(Optional ExportRange As Range) As Workbook
'A helper function for whenever students are deleted from the RecordsSheet
'Returns a finished workbook that can be saved
'Return nothing on error

    Dim NewBook As Workbook
    Dim SheetNameArray As Variant
    
    'Which sheets to make
    ReDim SheetNameArray(1 To 2)
        SheetNameArray(1) = "Cover Page"
        SheetNameArray(2) = "Simple Attendance"
        
    Set NewBook = ExportMakeBook(ExportRange, SheetNameArray)
        If NewBook Is Nothing Then
            GoTo Footer
        End If
    
    Set ExportFromRecords = NewBook
    
Footer:

End Function

Function ExportFromRoster(Optional ExportRange As Range) As Workbook
'A helper function for whenever students are deleted from the RosterSheet
'Returns a finished workbook that can be saved
'Return nothing on error

    Dim NewBook As Workbook
    Dim SheetNameArray As Variant
    
    'Which sheets to make
    ReDim SheetNameArray(1 To 4)
        SheetNameArray(1) = "Cover Page"
        SheetNameArray(2) = "Roster Page"
        SheetNameArray(3) = "Simple Attendance"
        SheetNameArray(4) = "Detailed Attendance"

    Set NewBook = ExportMakeBook(ExportRange, SheetNameArray)
        If NewBook Is Nothing Then
            GoTo Footer
        End If
    
    Set ExportFromRoster = NewBook
    
Footer:

End Function

Function ExportCoverSheet(OldBook As Workbook, NewBook As Workbook) As Long
'Makes a simple cover sheet for the new book
'Passing OldBook to make sure we are pulling from the correct workbook, even if the NewBook is active
'Returns 1 if successful

    Dim NewSheet As Worksheet
    Dim i As Long
    Dim CoverInfoArray() As Variant

    ExportCoverSheet = 0
    OldBook.Activate
    
    'Grab information from cover sheet. Avoiding copy and paste so we don't run into problems with MacOS
    CoverInfoArray = GetCoverInfo()
    
    'Make a new sheet and insert information
    With NewBook
        Set NewSheet = .Sheets.Add(After:=.Sheets(.Sheets.Count))
        NewSheet.Name = "Cover Page"
    End With
    
    For i = 1 To UBound(CoverInfoArray, 2)
        NewSheet.Cells(i, 1).Value = CoverInfoArray(1, i)
        NewSheet.Cells(i, 2).Value = CoverInfoArray(2, i)
        
        'Format the date
        If CoverInfoArray(1, i) = "Date" Then
            NewSheet.Cells(i, 2).NumberFormat = "mm/dd/yyyy"
        End If
    Next i
    
    NewSheet.Range("A1").EntireColumn.AutoFit
    
    'Return
    ExportCoverSheet = 1

Footer:

End Function

Function ExportDetailedSheet(OldBook As Workbook, NewBook As Workbook, Optional ExportRange As Range) As Long
'Exports a line for each time a student was marked present for an activity
'Passing a range only exports for those students. Range should be from the RecordsSheet
'Returns 1 if successful, 0 otherwise

    Dim OldRecordsSheet As Worksheet
    Dim OldRosterSheet As Worksheet
    Dim NewSheet As Worksheet
    Dim OldRecordsNameRange As Range
    Dim RosterHeaderRange As Range
    Dim ActivityHeaderRange As Range
    Dim CopyRange As Range
    Dim PasteRange As Range
    Dim c As Range
    Dim d As Range
    Dim e As Range
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim HeaderArray As Variant
    
    ExportDetailedSheet = 0

    Set OldRecordsSheet = OldBook.Worksheets("Records Page")
    Set OldRosterSheet = OldBook.Worksheets("Roster Page")

    'Make a new sheet
    With NewBook
        Set NewSheet = .Sheets.Add(After:=.Sheets(.Sheets.Count))
        NewSheet.Name = "Detailed Attendance"
    End With

    'Grab the headers for the Roster and activity headers
    OldBook.Activate
    
    Set RosterHeaderRange = Range("RosterHeadersList")
    Set ActivityHeaderRange = Range("ActivityHeadersList")
        If RosterHeaderRange Is Nothing Or ActivityHeaderRange Is Nothing Then
            GoTo Footer
        End If
    
    j = RosterHeaderRange.Cells.Count
    k = ActivityHeaderRange.Cells.Count
        
    ReDim HeaderArray(1 To 2, 1 To j + k)
    i = 1
    
    For Each c In RosterHeaderRange
        HeaderArray(1, i) = c.Value
        
        i = i + 1
    Next c
    
    For Each d In ActivityHeaderRange
        HeaderArray(1, i) = d.Value
    
        i = i + 1
    Next d

    'Put in headers on the new sheet
    Set PasteRange = NewSheet.Range("A1")
    
    For i = 1 To UBound(HeaderArray, 2)
        PasteRange.Offset(0, i - 1).Value = HeaderArray(1, i)
    Next i

    'If no export range was passed, grab ALL attendance
    Set OldRecordsNameRange = FindRecordsName(OldRecordsSheet)
    
    If ExportRange Is Nothing Then
        Set CopyRange = OldRecordsNameRange
    'Match on the Roster sheet, if needed
    ElseIf ExportRange.Worksheet.Name = "Roster Page" Then
        Set CopyRange = FindName(ExportRange.Offset(0, 1), OldRecordsNameRange)
    Else
        Set CopyRange = ExportRange
    End If
    
    'Go through each student and pass to search attendance and copy over information
    For Each c In CopyRange
        Set PasteRange = NewSheet.Range("B:B").Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Offset(1, -1)
    
        Call ExportDetailedSheetHelper(NewSheet, OldRosterSheet, OldRecordsSheet, c, HeaderArray)
    Next c
    
    'Make a table
    Call MakeTable(NewSheet)
    
    'Delete the first column
    NewSheet.Range("A1").EntireColumn.Delete
    
    ExportDetailedSheet = 1

Footer:

End Function

Function ExportDetailedSheetHelper(NewSheet As Worksheet, RosterSheet As Worksheet, RecordsSheet As Worksheet, RecordsNameCell As Range, HeaderArray As Variant) As Long
'Takes a name on the RecordsSheet, searches across the roster and records sheets for demographic and activity information
'Copies the information onto the passed NewSheet
'Returns 1 if successful

    Dim RecordsLabelRange As Range
    Dim RosterNameRange As Range
    Dim RosterMatchCell As Range
    Dim AttendanceRange As Range
    Dim LabelCell As Range
    Dim CopyRange As Range
    Dim PasteRange As Range
    Dim c As Range
    Dim d As Range
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim RosterTable As ListObject

    ExportDetailedSheetHelper = 0

    Set RosterTable = RosterSheet.ListObjects(1)
    Set RosterNameRange = RosterTable.ListColumns("First").DataBodyRange
    Set RecordsLabelRange = FindRecordsLabel(RecordsSheet)
    
    'Find where to start pasting
    Set c = NewSheet.Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    Set PasteRange = NewSheet.Cells(c.Row, 1)
    
    'Find the name on the RosterSheet
    Set RosterMatchCell = FindName(RecordsNameCell, RosterNameRange)
        If RosterMatchCell Is Nothing Then
            GoTo Footer
        End If
        
     'Grab demographics
     Set CopyRange = FindTableRow(RosterSheet, RosterMatchCell)
    
     i = 1
     For Each d In CopyRange
         HeaderArray(2, i) = d.Value
     
         i = i + 1
     Next d
    
    j = i
    k = 1
    
    'Take the passed name on the RecordsSheet, search across for each "1"
    Set AttendanceRange = RecordsLabelRange.Offset(RecordsNameCell.Row - 1, 0)
    
    For Each c In AttendanceRange
        If c.Value <> 1 Then
            GoTo NextActivity
        End If
        
        'Grab activity headers
        Set LabelCell = RecordsSheet.Cells(1, c.Column)
        Set CopyRange = FindRecordsActivityHeaders(RecordsSheet, LabelCell)
        
        i = j
        For Each d In CopyRange
            HeaderArray(2, i) = d.Value
            
            i = i + 1
        Next d

        'Copy over
        For i = 1 To UBound(HeaderArray, 2)
            PasteRange.Offset(k, i - 1).Value = HeaderArray(2, i)
        Next i
        
        k = k + 1
NextActivity:
    Next c
    
    ExportDetailedSheetHelper = 1

Footer:

End Function

Function ExportLocalSave(OldBook As Workbook, NewBook As Workbook) As Long
'For making a local save
'Returns 1 if successful, 2 if canceled
'Returns 0 on error

    Dim CoverSheet As Worksheet
    Dim CenterString As String
    Dim FileName As String
    Dim LocalPath As String
    Dim SaveName As String
    Dim SubDate As String
    Dim SubTime As String

    ExportLocalSave = 0

    Set CoverSheet = Worksheets("Cover Page")

    'Pull in center from the CoverSheet, date, and time
    CenterString = CoverSheet.Range("A:A").Find("Center", , xlValues, xlWhole).Offset(0, 1).Value
    SubDate = Format(Date, "yyyy-mm-dd")
    SubTime = Format(Time, "hh-nn AM/PM")

    'Make a file name using the center, date, and time of submission
    FileName = CenterString & " " & SubDate & "." & SubTime & ".xlsm"

    'Where the OldBook is stored
    OldBook.Activate
    LocalPath = GetLocalPath(ThisWorkbook.path)
    'For Win and Mac
    If Application.OperatingSystem Like "*Mac*" Then
        SaveName = Application.GetSaveAsFilename(LocalPath & "/" & FileName) ', "Excel Files (*.xlsm), *.xlsm")  MacOS sandboxing can't use file filters
        
        If SaveName = "False" Then
            NewBook.Close savechanges:=False
            ExportLocalSave = 2
            
            GoTo Footer
        End If
        
        NewBook.SaveAs FileName:=LocalPath & "/" & FileName, FileFormat:=xlOpenXMLWorkbookMacroEnabled
        'ActiveWorkbook.Close SaveChanges:=False
    Else
        SaveName = Application.GetSaveAsFilename(LocalPath & "\" & FileName, "Excel Files (*.xlsm), *.xlsm")
        
        If SaveName = "False" Then
            NewBook.Close savechanges:=False
            ExportLocalSave = 2
            
            GoTo Footer
        End If
        
        NewBook.SaveAs FileName:=SaveName, FileFormat:=xlOpenXMLWorkbookMacroEnabled
        'ActiveWorkbook.Close SaveChanges:=False
    End If
    
    'Everything worked
    ExportLocalSave = 1

Footer:

End Function

Function ExportMakeBook(Optional ExportRange As Range, Optional SheetArray As Variant, Optional TotalsOnly As String) As Workbook
'Sub exportest(SheetArray As Variant, Optional ExportRange As Range)
'Not passing a range exports all students
'Container function for each section of exporting
'Every sheet passed gets included in the returned book

    Dim NewBook As Workbook
    Dim OldBook As Workbook
    Dim NewSheet As Worksheet
    Dim RosterSheet As Worksheet
    Dim RecordsSheet As Worksheet
    Dim ReportSheet As Worksheet
    Dim i As Long
    Dim j As Long
    Dim SheetValue As Long
    Dim ErrorMessage As String
    Dim SheetName As String
    Dim ReadyArray() As Variant
    
    Set OldBook = ThisWorkbook
    Set RosterSheet = Worksheets("Roster Page")
    Set RecordsSheet = Worksheets("Records Page")
    Set ReportSheet = Worksheets("Report Page")
    
    'Check the pages that we need to, based on the passed sheet array
    ReadyArray = GetReadyToExport(SheetArray)
        If Not IsArray(ReadyArray) Or IsEmpty(ReadyArray) Then
            GoTo Footer
        End If
    
    For i = 1 To UBound(ReadyArray, 2)
        SheetName = ReadyArray(1, i)
        SheetValue = ReadyArray(2, i)
        
        If SheetValue = 0 Then
            Select Case SheetName
            
                Case "Cover Page"
                    ErrorMessage = "- Please enter your name, date, and center on the Cover Page"
                
                Case "Roster Page"
                    ErrorMessage = ErrorMessage & vbCr & "- You have no students on your roster. Please add your students and parse the roster."
                
                Case "Records Page"
                    ErrorMessage = ErrorMessage & vbCr & "- You have no saved attendance information. Please parse your roster and add an activity."
                
                Case "Report Page"
                    ErrorMessage = ErrorMessage & vbCr & "- There are no totals on the Report Page. Please tabulate your student totals."
                    
            End Select
        End If
    Next i
    
    'If there was an error
    If Len(ErrorMessage) > 0 Then
        MsgBox (ErrorMessage)
        
        GoTo Footer
    End If
    
    'Create new book and add sheets
    Set NewBook = Workbooks.Add
    
    For i = 1 To UBound(SheetArray)
        SheetName = SheetArray(i)
        
        Select Case SheetName
        
            Case "Cover Page"
                j = ExportCoverSheet(OldBook, NewBook)

            Case "Report Page" 'Not created when only exporting some students
                j = ExportReportSheet(OldBook, NewBook)
            
            Case "Roster Page"
                j = ExportRosterSheet(OldBook, NewBook, ExportRange)
            
            Case "Simple Attendance"
                j = ExportSimpleSheet(OldBook, NewBook, ExportRange)
            
            Case "Detailed Attendance"
                j = ExportDetailedSheet(OldBook, NewBook, ExportRange)
        
        End Select
        
        If j <> 1 Then
            GoTo ErrorMessage
        End If
    Next i
    
    'Delete "Sheet1"
    Set NewSheet = NewBook.Worksheets("Sheet1")
    NewSheet.Delete
    
    'Return
    Set ExportMakeBook = NewBook
    
    GoTo Footer
    
ErrorMessage:
    MsgBox ("Something has gone wrong, please close and reopen this file, then try again." & vbCr _
        & "If the problem persists, please contact the state office.")
            
    NewBook.Close savechanges:=False
    
Footer:
    
End Function

Function ExportReportSheet(OldBook As Workbook, NewBook As Workbook) As Long
'Grabs the entire Report sheet
'Returns 1 if successful, 0 otherwise

    Dim OldReportSheet As Worksheet
    Dim NewSheet As Worksheet
    Dim CopyRange As Range
    Dim PasteRange As Range
    Dim c As Range
    Dim OldReportTable As ListObject
    Dim NewTable As ListObject
    
    Set OldReportSheet = OldBook.Worksheets("Report Page")
    Set OldReportTable = OldReportSheet.ListObjects(1)
    
    ExportReportSheet = 0
    
    'Make a new sheet
    With NewBook
        Set NewSheet = .Sheets.Add(After:=.Sheets(.Sheets.Count))
        NewSheet.Name = "Report Page"
    End With
    
    'Copy and paste the entire table data
    Set CopyRange = OldReportTable.Range
    Set c = NewSheet.Range("A1")
    Set PasteRange = c.Resize(CopyRange.Rows.Count, CopyRange.Columns.Count)
    
    CopyRange.Copy Destination:=PasteRange
    
    'Chop off the first row, autofit
    c.EntireColumn.Delete
    PasteRange.EntireColumn.AutoFit
    
    ExportReportSheet = 1

Footer:

End Function

Function ExportRosterSheet(OldBook As Workbook, NewBook As Workbook, Optional ExportRange As Range) As Long
'Reproduces the roster in the NewBook. Grabs all students by default
'Passing a range restricts to only those students
'Range should only come from the RosterSheet

    Dim OldRosterSheet As Worksheet
    Dim NewSheet As Worksheet
    Dim CopyRange As Range
    Dim PasteRange As Range
    Dim c As Range
    Dim d As Range
    Dim OldRosterTable As ListObject
    Dim NewTable As ListObject
    
    Set OldRosterSheet = OldBook.Worksheets("Roster Page")
    Set OldRosterTable = OldRosterSheet.ListObjects(1)
    
    ExportRosterSheet = 0
    
    'Make a new sheet
    With NewBook
        Set NewSheet = .Sheets.Add(After:=.Sheets(.Sheets.Count))
        NewSheet.Name = "Roster Page"
    End With
    
    'Copy the headers and autofit
    Set CopyRange = OldRosterTable.HeaderRowRange
    Set c = NewSheet.Range("A1")
    Set PasteRange = c.Resize(1, CopyRange.Columns.Count)
    
    PasteRange.Value = CopyRange.Value
    PasteRange.EntireColumn.AutoFit
    
    'Match student names on Roster, if needed
    If ExportRange Is Nothing Then
        Set ExportRange = OldRosterTable.ListColumns("First").DataBodyRange
    ElseIf ExportRange.Worksheet.Name = "Records Sheet" Then
        Set c = FindName(ExportRange, OldRosterTable.ListColumns("First").DataBodyRange)
        
        If c Is Nothing Then
            GoTo Footer
        End If
        
        Set ExportRange = c 'Bad form, change later
    End If
    
    'If there is no passed ranged, reproduce the entire table
    If ExportRange Is Nothing Then
        Set CopyRange = OldRosterTable.Range
        Set PasteRange = NewSheet.Range("A1").Resize(CopyRange.Rows.Count, CopyRange.Columns.Count)
        
        PasteRange.Value = CopyRange.Value
    'If a range was passed, copy over each row
    Else
        Set PasteRange = NewSheet.Range("A2")

        Call CopyRow(OldRosterSheet, ExportRange, NewSheet, PasteRange)
    End If
        
    'Make a table and remove the first column
    Call MakeTable(NewSheet)
    NewSheet.Range("A1").EntireColumn.Delete
    
    ExportRosterSheet = 1

Footer:

End Function

Function ExportSimpleSheet(OldBook As Workbook, NewBook As Workbook, Optional ExportRange As Range) As Long
'Exports if a student was present, absent, or unrecorded for every activity
    '1 - present
    '0 - absent
    '[nothing] - N/A
'Passing a range only exports for those students. Range should be from the RecordsSheet
'Returns 1 if successful, 0 otherwise

    Dim OldRecordsSheet As Worksheet
    Dim NewSheet As Worksheet
    Dim CopyRange As Range
    Dim PasteRange As Range
    Dim OldRecordsHeaderRange As Range
    Dim OldRecordsFoundNames As Range
    Dim OldRecordsNameRange As Range
    Dim c As Range
    Dim d As Range
    Dim LRow As Long
    Dim LCol As Long
    
    ExportSimpleSheet = 0
    
    Set OldRecordsSheet = OldBook.Worksheets("Records Page")
    
    'Make a new sheet
    With NewBook
        Set NewSheet = .Sheets.Add(After:=.Sheets(.Sheets.Count))
        NewSheet.Name = "Simple Attendance"
    End With
    
    'If there is no passed range, simply copy over the entire sheet
    If ExportRange Is Nothing Then
        LRow = OldRecordsSheet.Range("A:A").Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
        LCol = OldRecordsSheet.Range("1:1").Find("*", SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
        
        Set c = OldRecordsSheet.Range("A1")
        Set d = OldRecordsSheet.Cells(LRow, LCol)
        Set CopyRange = OldRecordsSheet.Range(c, d)
        Set PasteRange = NewSheet.Range(CopyRange.Address)
    
        PasteRange.Value = CopyRange.Value
        
        GoTo RemovePadding
    End If
    
    'If a range of names was passed
    'Grab the headers and copy over
    Set c = OldRecordsSheet.Range("A:A").Find("H BREAK", , xlValues, xlWhole)
    Set d = OldRecordsSheet.Range("1:1").Find("*", SearchOrder:=xlByColumns, SearchDirection:=xlPrevious)
    Set CopyRange = OldRecordsSheet.Range("A1", Cells(c.Row, d.Column).Address)
    Set PasteRange = NewSheet.Range(CopyRange.Address)
        PasteRange.Value = CopyRange.Value
    
    'Find the students on the RecordsSheet, if needed
    Set OldRecordsNameRange = FindRecordsName(OldRecordsSheet)
    
    If ExportRange.Worksheet.Name = "Roster Page" Then
        OldBook.Activate
        
        Set CopyRange = FindName(ExportRange.Offset(0, 1), OldRecordsNameRange)
        
        If CopyRange Is Nothing Then
            GoTo Footer
        End If
    Else
        Set CopyRange = ExportRange
    End If
            
    'Grab each row and copy over
    Set PasteRange = NewSheet.Range(c.Offset(1, 0).Address)
    
    Call CopyRow(OldRecordsSheet, CopyRange, NewSheet, PasteRange)
        
RemovePadding:
    'Bold the headers, then delete the padding cells
    Set c = NewSheet.Range("A:A").Find("H BREAK", , xlValues, xlWhole)
    Set d = NewSheet.Range("1:1").Find("V BREAK", , xlValues, xlWhole)
    
    NewSheet.Range(c, d).EntireRow.Font.Bold = True
    
    c.EntireRow.Delete
    d.EntireColumn.Delete
    
    ExportSimpleSheet = 1
    
Footer:

End Function

Function ExportSharePoint(OldBook As Workbook, NewBook As Workbook) As Long
'Sends the cover sheet and report to SharePoint
'Returns 1 if successful
'Returns 0 on error
    Dim CoverSheet As Worksheet
    Dim CenterString As String
    Dim FileName As String
    Dim SaveName As String
    Dim SubDate As String
    Dim SubTime As String
    Dim SpPath As String
    Dim TempArray() As Variant

    ExportSharePoint = 0

    Set CoverSheet = Worksheets("Cover Page")

    'Pull in center from the CoverSheet, date, and time
    CenterString = CoverSheet.Range("A:A").Find("Center", , xlValues, xlWhole).Offset(0, 1).Value
    SubDate = Format(Date, "yyyy-mm-dd")
    SubTime = Format(Time, "hh-nn AM/PM")

    'Make a file name using the center, date, and time of submission
    FileName = CenterString & " " & SubDate & "." & SubTime & ".xlsm"

    'The address where the new book will be save in SharePoint
    SpPath = "https://uwnetid.sharepoint.com/sites/partner_university_portal/Data%20Portal/Report%20Submissions/"

    'Upload
    NewBook.SaveAs FileName:=SpPath & "/" & FileName, FileFormat:=xlOpenXMLWorkbookMacroEnabled
    NewBook.Close savechanges:=False
    
    'Everything worked
    ExportSharePoint = 1
    
Footer:

End Function
