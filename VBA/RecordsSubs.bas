Attribute VB_Name = "RecordsSubs"
Option Explicit

Sub ClearRecords(Optional ExportCheck As Long)
'Remove all activities, students, and attendance from the Records Page
'ExportCheck passed from RosterClearButton()

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
    'Both students and activities, ExportCheck passed
    'ElseIf ExportCheck = vbYes Then
        'Set OldBook = ActiveWorkbook
        'Set NewBook = MakeNewBook(OldBook)
        'Set DelRange = FindRecordsName(RecordsSheet)
    
        'Call ExportSimpleAttendance(RecordsSheet, NewBook, DelRange)
        'Call ExportDetailedAttendance(RecordsSheet, RosterSheet, NewBook, DelRange)
        'Call SaveNewBook(OldBook, NewBook)
    
        'Set DelRange = FindRecordsAttendance(RecordsSheet)
        'DelRange.EntireRow.ClearContents
        'DelRange.EntireColumn.ClearContents
    'Both students and activities, ExportCheck failed
    Else
        Set DelRange = FindRecordsAttendance(RecordsSheet)
        DelRange.EntireRow.ClearContents
        DelRange.EntireColumn.ClearContents
    End If
    
    'Clear the report sheet
    'Call ClearReportButton
    
Footer:

End Sub

Sub RecordsPullAttendance(ActivitySheet As Worksheet, ActivityNameRange As Range, LabelCell As Range)
'Pulls attendance for all students marked "present" in the Records sheet to an activity Sheet

    Dim RecordsSheet As Worksheet
    Dim RecordsNameRange As Range
    Dim RecordsAttendanceRange As Range
    Dim RecordsLabelRange As Range
    Dim RecordsPresentRange As Range
    Dim RecordsAbsentRange As Range
    Dim ActivityPresentRange As Range
    Dim MatchCell As Range
    Dim c As Range
    Dim d As Range

    Set RecordsSheet = Worksheets("Records Page")
    
    'Check if there are both students and activities
    If CheckRecords(RecordsSheet) <> 1 Then
        GoTo Footer
    End If
    
    'Check that there are students
    If CheckTable(ActivitySheet) > 2 Then
        GoTo Footer
    End If
    
    Set RecordsNameRange = FindRecordsName(RecordsSheet)
    Set RecordsLabelRange = FindRecordsLabel(RecordsSheet, LabelCell)
    Set RecordsAttendanceRange = FindRecordsAttendance(RecordsSheet, , RecordsLabelRange)

    'Clear current checks on activity sheet and copy over saved Attendance
    ActivityNameRange.Offset(0, -1).Value = ""

    'Find all students marked present
    Set d = FindChecks(RecordsAttendanceRange)
    If d Is Nothing Then
        GoTo Footer
    End If
    
    Set RecordsPresentRange = d.Offset(0, -d.Column + 1)
    For Each c In RecordsPresentRange
        Set MatchCell = FindName(c, ActivityNameRange)
        If Not MatchCell Is Nothing Then
            MatchCell.Offset(0, -1).Value = "a"
        End If
    Next c

Footer:

End Sub



