Attribute VB_Name = "ReportSubs"
Option Explicit

Sub ClearReport()
'Called by a button and clearing the roster

    Dim ReportSheet As Worksheet
    Dim ReportTable As ListObject
    
    Set ReportSheet = Worksheets("Report Page")
    
    'Make sure there is a table. There always should be
    If CheckTable(ReportSheet) = 4 Then
        GoTo Footer
    End If
    
    Call UnprotectSheet(ReportSheet)
    
    'Make sure there is at least one activity. Only clear totals if there isn't
    Set ReportTable = ReportSheet.ListObjects(1)
    
    If ReportTable.ListRows.Count < 1 Then
        Call ClearReportTotals
        GoTo Footer
    End If
    
    'Clear everything under the header, then unlist and remake the table
    ReportTable.DataBodyRange.ClearContents
    ReportTable.DataBodyRange.EntireRow.Delete
    ReportTable.Unlist
    
    Set ReportTable = CreateReportTable
    Call FormatReportTable(ReportSheet, ReportTable)
    
Footer:

End Sub

Sub ClearReportTotals()
'Only clears the totals. Called when clearing the roster and clearing the entire report

    Dim CoverSheet As Worksheet
    Dim ReportSheet As Worksheet
    Dim DelRange As Range
    
    Set CoverSheet = Worksheets("Cover Page")
    Set ReportSheet = Worksheets("Report Page")
    
    If InStr(1, CoverSheet.Range("A1").Value, "College") > 0 Then
        Set DelRange = FindTableHeader(ReportSheet, "Select", "Other Grade")
    Else
        Set DelRange = FindTableHeader(ReportSheet, "Select", "Low Income")
    End If
    
    'Delete the row beneath the header
    DelRange.Offset(1, 0).ClearContents

End Sub

Sub ClearReportActivity(LabelCell As Range)
'Removes all previous data in a row on the Report sheet

    Dim ReportSheet As Worksheet
    Dim ReportTable As ListObject
    Dim ReportLabelCell As Range
    Dim DelRange As Range
    
    Set ReportSheet = Worksheets("Report Page")
    
    If ReportSheet.ListObjects.Count < 1 Then
        Call CreateReportTable
    End If
    
    'Find the row we need
    Set ReportTable = ReportSheet.ListObjects(1)
    Set ReportLabelCell = ReportTable.ListColumns("Label").DataBodyRange.Find(LabelCell.Value, , xlValues, xlWhole)
    
    If ReportLabelCell Is Nothing Then
        GoTo Footer
    End If
    
    Call UnprotectSheet(ReportSheet)

    'Clear the row if it's the total row, delete if it's for an activity
    If ReportLabelCell.Value = "Total" Then
        Call ClearReportTotals
    Else
        ReportLabelCell.EntireRow.Delete
    End If
    
Footer:

End Sub


