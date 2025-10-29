Attribute VB_Name = "ReportSubs"
Option Explicit

Sub ClearReportTotals()

    Err.Raise vbObjectError + 513, , "Wrong function. ReportClearTotals"

End Sub

Sub ReportClear()
'Called by a button and clearing the roster

    Dim ReportSheet As Worksheet
    Dim DelRange As Range
    Dim c As Range
    Dim d As Range
    Dim HeadersArray As Variant
    
    Set ReportSheet = Worksheets("Report Page")
    
    'Define area to delete
    Set c = ReportSheet.Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlNext)
    Set d = ReportSheet.Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    
    If c Is Nothing Then
        Set DelRange = ReportSheet.Range("A6")
    Else
        Set DelRange = ReportSheet.Range(c, d)
    End If

    'Delete everything
    Call UnprotectSheet(ReportSheet)
    Call RemoveTable(ReportSheet) 'This shouldn't be necessary
    
    DelRange.ClearContents
    DelRange.EntireRow.Delete
    
    'Remake the table
    Call MakeReportTable
    
Footer:

End Sub

Sub ReportClearActivity(Optional LabelCell As Range, Optional LabelString As String)
'Removes an activity from the ReportTable
'Break if nothing is passed

    Dim ReportSheet As Worksheet
    Dim ReportTable As ListObject
    Dim ReportLabelCell As Range
    Dim DelRange As Range
    Dim SearchString As String
    
    Set ReportSheet = Worksheets("Report Page")
    
    If Not ReportSheet.ListObjects.Count > 0 Then
        Call MakeReportTable
    End If
    
    'Either a range or string needs to be passed
    If Not LabelCell Is Nothing Then
        SearchString = LabelCell.Value
    ElseIf Len(LabelString) > 0 Then
        SearchString = LabelString
    Else
        GoTo Footer
    End If
    
    'Find the row we need
    Set ReportTable = ReportSheet.ListObjects(1)
    Set ReportLabelCell = ReportTable.ListColumns("Label").DataBodyRange.Find(SearchString, , xlValues, xlWhole)
        If ReportLabelCell Is Nothing Then
            GoTo Footer
        End If

    'Clear the row if it's the total row, delete if it's for an activity
    Call UnprotectSheet(ReportSheet)
    
    If ReportLabelCell.Value = "Total" Then
        Call ClearReportTotals
    Else
        Set DelRange = ReportLabelCell
        
        Call RemoveFromReport(DelRange)
    End If
    
Footer:

End Sub

Sub ReportClearTotals()
'Only clears the totals. Called when clearing the roster and clearing the entire report

    Dim ReportSheet As Worksheet
    Dim DelRange As Range
    Dim c As Range
    Dim ReportTable As ListObject
    
    Set ReportSheet = Worksheets("Report Page")
        If CheckReport(ReportSheet) > 3 Then
            GoTo Footer
        End If
    
    Set ReportTable = ReportSheet.ListObjects(1)
    Set DelRange = ReportTable.HeaderRowRange.Offset(1, 0)
    Set c = FindTableHeader(ReportSheet, "Select")
    
    'Take everything out, put in the Totals row headers
    DelRange.ClearContents
    Call TableResetReportTotalHeaders(ReportSheet, c)
    
Footer:

End Sub




