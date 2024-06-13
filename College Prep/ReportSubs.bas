Attribute VB_Name = "ReportSubs"
Option Explicit

Sub PullReportTotals()
'Also called when roster is parsed

    Dim RosterSheet As Worksheet
    Dim ReportSheet As Worksheet
    
    Set RosterSheet = Worksheets("Roster Page")
    Set ReportSheet = Worksheets("Report Page")
    
    'Unprotect
    Call UnprotectCheck(ReportSheet)

    'Define where the totals will go. It can be a discontiguous area
    Dim TotalRange As Range
    Dim RaceRange As Range
    Dim GenderRange As Range
    Dim GradeRange As Range
                           
    Set RaceRange = FindReportRange("White", "Other Race").Offset(1, 0)
    Set GenderRange = FindReportRange("Female", "Other Gender").Offset(1, 0)
    Set GradeRange = FindReportRange("6", "Other Grade").Offset(1, 0)
    Set TotalRange = FindReportRange("Total").Offset(1, 0)
    
    'Clear the contents
    TotalRange.EntireRow.ClearContents
    
    'Grab the entire first name column from the Roster Page
    Dim TempArray As Variant
    Dim NameRange As Range
    Dim TableLength As Long
    
    Set NameRange = RosterSheet.ListObjects(1).ListColumns("First").DataBodyRange
    TableLength = RosterSheet.ListObjects(1).ListRows.Count
    
    'Pass to be tabulated and paste in values
    TempArray = DemoTabulate(TableLength, NameRange, "Race")
    RaceRange = TempArray
    Erase TempArray
    
    TempArray = DemoTabulate(TableLength, NameRange, "Gender")
    GenderRange = TempArray
    Erase TempArray
    
    TempArray = DemoTabulate(TableLength, NameRange, "Grade")
    GradeRange = TempArray
    Erase TempArray
    
    'Total
    TotalRange = TableLength
    
    'Add information from the coversheet
    Dim CoverSheet As Worksheet
    Dim CenterRange As Range
    
    Set CoverSheet = Worksheets("Cover Page")
    Set CenterRange = ReportSheet.Range("B7")
    
    CenterRange = CoverSheet.Range("B5")
    CenterRange.Offset(0, 1) = CoverSheet.Range("B3")
    CenterRange.Offset(0, 2) = "Total"
    CenterRange.Offset(0, 3) = "N/A"
    CenterRange.Offset(0, 4) = CDate(CoverSheet.Range("B4"))
    CenterRange.Offset(0, 5) = "All students on the roster"
    
    'Apply bold font
    TotalRange.EntireRow.Font.Bold = True

End Sub

Sub ClearReportTotals()
'Only clears the totals. Called when clearing the roster and clearing the entire report

    Dim ReportSheet As Worksheet
    Dim DelRange As Range
    
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    Set ReportSheet = Worksheets("Report Page")
    Set DelRange = FindReportRange("Select", "Other Grade")
    
    'Unprotect
    Call UnprotectCheck(ReportSheet)
    
    'Delete the row beneath the header
    DelRange.Offset(1, 0).ClearContents

    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

End Sub

Function FindReportRange(StartString As String, Optional EndString As String) As Range
'Find either a single column or a range of columns based the header names

    Dim ReportSheet As Worksheet
    Dim StartRange As Range
    Dim EndRange As Range
    
    Set ReportSheet = Worksheets("Report Page")
    Set StartRange = ReportSheet.Range("6:6").Find(StartString, , xlValues, xlWhole)
    
    'If a bad string has been entered
    If StartRange Is Nothing Then
        MsgBox ("Something has gone wrong. The names of the columns on the Report Page may be incorrect.")
        GoTo Footer
    End If
    
    'If only one string was provided
    If Not Len(EndString) > 0 Then
        Set FindReportRange = StartRange
        GoTo Footer
    End If
    
    'If both strings are provided
    Set EndRange = ReportSheet.Range("6:6").Find(EndString, , xlValues, xlWhole)
    
    If EndRange Is Nothing Then
        MsgBox ("Something has gone wrong. The names of the columns on the Report Page may be incorrect.")
        GoTo Footer
    End If

    Set FindReportRange = ReportSheet.Range(StartRange, EndRange)

Footer:

End Function

