Attribute VB_Name = "DeletedSubs"

Function CreateReportTable2() As ListObject
'Grabs headers from reference page, unmakes and remakes the table
'Called when adding or deleting rows, tabulating totals

    Dim ReportSheet As Worksheet
    Dim CoverSheet As Worksheet
    Dim ReportTableStart As Range
    Dim ReportLabelRange As Range
    Dim ReportTableRange As Range
    Dim HeaderRange As Range
    Dim BoxRange As Range
    Dim DelRange As Range
    Dim i As Long
    Dim HeaderArray() As Variant
    Dim TotalsArray() As Variant
    Dim CenterInfoArray() As Variant
    Dim ReportTable As ListObject
    
    Set CoverSheet = Worksheets("Cover Page")
    Set ReportSheet = Worksheets("Report Page")
    Set ReportTableStart = ReportSheet.Range("A:A").Find("Select", , xlValues, xlWhole)
       
    Call UnprotectSheet(ReportSheet)
    
    'Remove any existing filters, unlist the table and remove formatting
    If ReportSheet.AutoFilterMode = True Then
        ReportSheet.AutoFilterMode = False
    End If
    
    Call RemoveTable(ReportSheet)
    
    'Reset headers. This creates two rows
    HeaderArray = Application.Transpose(ActiveWorkbook.Names("ReportHeadersList").RefersToRange.Value)
    TotalsArray = Application.Transpose(ActiveWorkbook.Names("ReportTotalsRowList").RefersToRange.Value)
    Call TableResetHeaders(ReportSheet, ReportTableStart, HeaderArray)
    
    Set ReportLabelRange = ReportTableStart.EntireRow.Find("Label", , xlValues, xlWhole)
    Call TableResetHeaders(ReportSheet, ReportLabelRange.Offset(1, 0), TotalsArray) 'The two columns before this are pulled from the cover sheet
    
    'Define where to put information and pull in in values from the cover sheet
    ReDim CenterInfoArray(1 To 3, 1 To 2)
        Set CenterInfoArray(1, 1) = ReportTableStart.EntireRow.Find("Center", , xlValues, xlWhole)
        Set CenterInfoArray(2, 1) = ReportTableStart.EntireRow.Find("Name", , xlValues, xlWhole)
        Set CenterInfoArray(3, 1) = ReportTableStart.EntireRow.Find("Date", , xlValues, xlWhole)
        
        Set CenterInfoArray(1, 2) = CoverSheet.Range("A:A").Find("Center", , xlValues, xlWhole)
        Set CenterInfoArray(2, 2) = CoverSheet.Range("A:A").Find("Name", , xlValues, xlWhole)
        Set CenterInfoArray(3, 2) = CoverSheet.Range("A:A").Find("Date", , xlValues, xlWhole)
    
    For i = 1 To UBound(CenterInfoArray)
        CenterInfoArray(i, 1).Offset(1, 0).Value = CenterInfoArray(i, 2).Offset(0, 1).Value
    Next i
    
    'Define table range and clear formats
    Set ReportTableRange = FindTableRange(ReportSheet)
    ReportTableRange.ClearFormats
    
    'Make a new table
    Set ReportTable = ReportSheet.ListObjects.Add(SourceType:=xlSrcRange, Source:=ReportTableRange, _
        xlListObjectHasHeaders:=xlYes)
    ReportTable.Name = "ReportTable"
    
    'Look for blank rows if there are more than two rows
    If ReportTable.DataBodyRange.Rows.Count < 2 Then
        GoTo FormatTable
    End If
        
    Set ReportLabelRange = ReportTable.ListColumns("Label").DataBodyRange
    Set DelRange = FindBlanks(ReportLabelRange)
    
    If Not DelRange Is Nothing Then
        Call RemoveRows(ReportSheet, ReportTable.DataBodyRange, ReportLabelRange, DelRange)
        Set ReportTable = ReportSheet.ListObjects(1)
    End If
    
FormatTable:
    'Format
    ReportTable.ShowTableStyleRowStripes = False
    
    Call TableFormatReport(ReportSheet, ReportTable)
    
    'Add Marlett Boxes to everything but the Totals row
    Set BoxRange = ReportTable.ListColumns("Select").DataBodyRange
    Call AddMarlettBox(BoxRange)
    ReportTable.ListColumns("Select").DataBodyRange(1, 1).Font.Name = "Aptos Narrow" 'This can be anything except Marlett
    
    'Format the Date column
    ReportTable.ListColumns("Date").DataBodyRange.NumberFormat = "mm/dd/yyyy"
    
    'Autofit Description column
    ReportTable.ListColumns("Description").Range.EntireColumn.AutoFit

    'Return
    Set CreateReportTable = ReportTable

Footer:

End Function

Function FindPresent2(RecordsSheet As Worksheet, LabelCell As Range, Optional OperationString As String) As Range 'Surperfluous, the FindChecks function can do this already
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
