Attribute VB_Name = "RecordsSubs"
Option Explicit

Sub RecordsClear()
'Remove all activities, students, and attendance from the Records Page

    Dim OldBook As Workbook
    Dim NewBook As Workbook
    Dim RecordsSheet As Worksheet
    Dim RosterSheet As Worksheet
    Dim DelRange As Range
    Dim DelCheck As Long
    
    Set RecordsSheet = Worksheets("Records Page")
    Set RosterSheet = Worksheets("Roster Page")
    
    'Check if there are students or activities
    DelCheck = CheckRecords(RecordsSheet)
    
    'No students or activities
    If DelCheck = 4 Then
        GoTo Footer
    'Only students
    ElseIf DelCheck = 3 Then
        Set DelRange = FindRecordsName(RecordsSheet)
        DelRange.EntireRow.ClearContents
    'Only activities
    ElseIf DelCheck = 2 Then
        Set DelRange = FindRecordsLabel(RecordsSheet)
        DelRange.EntireColumn.ClearContents
    'Both
    Else
        Set DelRange = FindRecordsAttendance(RecordsSheet)
        DelRange.EntireRow.ClearContents
        DelRange.EntireColumn.ClearContents
    End If
    
Footer:

End Sub

Sub RecordsClearAttendance(RecordsSheet As Worksheet, Optional LabelString As String)
'Clears all attendance information, without clearing students and activity headers
'If a label is passed, only clears the attendance for that activity
'Also clears the corresponding activities on the ReportSheet

    Dim ReportSheet As Worksheet
    Dim RecordsLabelRange As Range
    Dim RecordsLabelCell As Range
    Dim ReportLabelCell As Range
    Dim DelRange As Range
    
    If CheckRecords(RecordsSheet) <> 1 Then
        GoTo Footer
    End If
    
    Set ReportSheet = Worksheets("Report Page")
    Set RecordsLabelRange = FindRecordsLabel(RecordsSheet)
    
    'If a label was passed
    If Len(LabelString) > 0 Then
        Set ReportLabelCell = FindReportLabel(ReportSheet, LabelString)
        Set RecordsLabelCell = FindRecordsLabel(RecordsSheet, , LabelString)
            If RecordsLabelCell Is Nothing Then
                GoTo Footer
            End If
    End If
    
    'Define the attendance range and delete
    Set DelRange = FindRecordsAttendance(RecordsSheet, , RecordsLabelCell)
        If DelRange Is Nothing Then
           GoTo Footer
        End If
        
    DelRange.ClearContents
     
    'Delete from the ReportSheet
    If ReportLabelCell Is Nothing Then
        Call ReportClear
        Call TabulateReportTotals
    Else
        Set DelRange = ReportLabelCell
        Call RemoveFromReport(DelRange)
    End If

Footer:

End Sub





