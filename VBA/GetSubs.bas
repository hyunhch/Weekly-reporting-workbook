Attribute VB_Name = "GetSubs"
Option Explicit

Function GetActivityInfo(RecordsSheet As Worksheet, ReportSheet As Worksheet, LabelCell As Range) As Variant
'Creates an array to put in the Report Page for the passed label
'Returns empty if the search range can't be found
'Values that can't be found are returned as ""

    Dim CoverSheet As Worksheet
    Dim InfoHeaderRange As Range
    Dim RecordsLabelRange As Range
    Dim CoverSearchRange As Range
    Dim RecordsSearchRange As Range
    Dim c As Range
    Dim d As Range
    Dim i As Long
    Dim InfoArray() As Variant
    
    Set CoverSheet = Worksheets("Cover Page")
    
    'Define how long the array needs to be by finding the bookends. It will be between "Name" and "Total"
    Set c = FindTableHeader(ReportSheet, "Name", "Total")
    If c Is Nothing Then
        GoTo Footer
    End If
    
    Set InfoHeaderRange = c.Resize(1, c.Columns.Count - 2).Offset(0, 1) 'Chop off two cells and shift one to the right
    
    'Define where to search
    Set RecordsLabelRange = FindRecordsLabel(RecordsSheet, LabelCell)
    If RecordsLabelRange Is Nothing Then
        GoTo Footer
    End If

    Set c = RecordsSheet.Range("1:1").Find("V BREAK", , xlValues, xlWhole).Offset(0, -1) 'Cell where activity headers start
    Set d = RecordsSheet.Range("A:A").Find("H BREAK", , xlValues, xlWhole).Offset(-1, 0) 'Row where the activity headers stop
    Set RecordsSearchRange = c.Resize(d.Row - 1, 1) 'Values are in the RecordsLabelRange column
    
    'Read into an array
    ReDim InfoArray(1 To InfoHeaderRange.Cells.Count, 1 To 2)
    
    i = 1
    For Each c In InfoHeaderRange
        Set d = RecordsSearchRange.Find(c.Value, , xlValues, xlWhole)
        If Not d Is Nothing Then
            InfoArray(i, 1) = c.Value
            InfoArray(i, 2) = RecordsSheet.Cells(d.Row, RecordsLabelRange.Column)
        End If
        
        i = i + 1
    Next c
        
    GetActivityInfo = InfoArray
    
Footer:
   
End Function

Function GetEdition() As String

    Dim FileName As String
    Dim TempName As String

    FileName = ThisWorkbook.Name
    TempName = Left(FileName, InStrRev(FileName, ".") - 1)
    GetEdition = Right(TempName, Len(TempName) - InStrRev(TempName, " "))
    
End Function

Function GetReadyToSave(CoverSheet As Worksheet, ReportSheet As Worksheet, Optional RecordsSheet As Worksheet, Optional RosterSheet As Worksheet) As Variant
'Checks that everything is ready to save or export
'For each sheet, returns 1 if it's filled out, 0 if it's not

    Dim SearchRange As Range
    Dim c As Range
    Dim d As Range
    Dim SearchString As String
    Dim ResultString As String
    Dim i As Long
    Dim j As Long
    Dim TempArray() As Variant
    Dim SearchArray() As Variant
    
    'Put in the two sheets that will always be use
    ReDim TempArray(1 To 2, 1 To 2)
        TempArray(1, 1) = "Cover Page"
        TempArray(1, 2) = "Report Page"
        TempArray(2, 1) = 0
        TempArray(2, 2) = 0
        
    'Redim the array depending on the number of arguments passed
    i = 2
    
    'Roster Sheet
    If RosterSheet Is Nothing Then
        GoTo CheckRecords
    End If
    
    i = i + 1
    ReDim Preserve TempArray(1 To 2, 1 To i)
        TempArray(1, i) = "Roster Page"
        TempArray(2, i) = 0
    
    j = CheckTable(RosterSheet)
    If Not j > 2 Then 'At least one student on the roster table
        TempArray(2, i) = 1
    End If
    
CheckRecords:
    'Records Sheet
    If RecordsSheet Is Nothing Then
        GoTo CheckCover
    End If

    i = i + 1
    ReDim Preserve TempArray(1 To 2, 1 To i)
        TempArray(1, i) = "Records Page"
        TempArray(2, i) = 0
        
    j = CheckRecords(RecordsSheet)
    If Not j > 1 Then 'Both students and activities
        TempArray(2, i) = 1
    End If

CheckCover:
    'Cover sheet
    ReDim SearchArray(1 To 3)
        SearchArray(1) = "Name"
        SearchArray(2) = "Date"
        SearchArray(3) = "Center"

    For j = 1 To UBound(SearchArray)
        SearchString = SearchArray(j)
        Set c = CoverSheet.Range("A:A").Find(SearchString, , xlValues, xlWhole).Offset(0, 1)
        
        If Len(c.Value) < 1 Then
            GoTo CheckReport
        End If
    Next j
    
    TempArray(2, 1) = 1
    
CheckReport:
    'Report Sheet
    j = CheckTable(ReportSheet)
    If Not j > 2 Then
        TempArray(2, 2) = 1
    End If
    
    'Return
    GetReadyToSave = TempArray

Footer:

End Function

Function GetSubmissionInfo(ReportSheet As Worksheet, Optional PullDate As String) As Variant
'Grabs the name, center, and date from the Cover Page
'Will grab the date from the CoverSheet if "Yes" is passed for PullDate

    Dim CoverSheet As Worksheet
    Dim ReportSearchRange As Range
    Dim c As Range
    Dim d As Range
    Dim i As Long
    Dim SearchString As String
    Dim FoundString As String
    Dim TempArray() As Variant
    
    Set CoverSheet = Worksheets("Cover Page")
    
    'Define what part of the headers we want. It'll be between the "Select" and "Date" headers unless "Yes" was passed
    If PullDate = "Yes" Then
        Set c = FindTableHeader(ReportSheet, "Select", "Label")
    Else
        Set c = FindTableHeader(ReportSheet, "Select", "Date")
    End If
    
    If c Is Nothing Then
        GoTo Footer
    End If

    Set ReportSearchRange = c.Resize(1, c.Columns.Count - 2).Offset(0, 1)
    If ReportSearchRange Is Nothing Then
        GoTo Footer
    End If

    'Search on the CoverSheet and find corresponding values
    ReDim TempArray(1 To ReportSearchRange.Cells.Count, 1 To 2)
    
    i = 1
    For Each c In ReportSearchRange
        SearchString = c.Value
        TempArray(i, 1) = SearchString
        
        Set d = CoverSheet.Range("A:A").Find(SearchString, , xlValues, xlWhole)
        
        If Not d Is Nothing Then
            FoundString = d.Offset(0, 1).Value
            TempArray(i, 2) = FoundString
        End If
        
        i = i + 1
    Next c

    'Return array
    GetSubmissionInfo = TempArray
    
Footer:

End Function
