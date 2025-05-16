Attribute VB_Name = "ReportButtonSubs"
Option Explicit

Sub ClearReportButton()
'Clears and resets the report sheet

    Dim ReportSheet As Worksheet
    
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    Set ReportSheet = Worksheets("Report Page")
    
    'Make sure there is a table. There always should be
    If CheckTable(ReportSheet) = 4 Then
        GoTo Footer
    End If
    
    'Clear
    Call UnprotectSheet(ReportSheet)
    Call ClearReport

    'Reprotect
    Call ResetProtection
    
Footer:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
End Sub

Sub OpenTabulateActivityButton()
'Checks that there is anything new to tabulate and open the TabulateActivityForm

    Dim ReportSheet As Worksheet
    Dim RecordsSheet As Worksheet
    Dim ReportLabelRange As Range
    Dim RecordsLabelRange As Range
    Dim AddRange As Range
    Dim c As Range

    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    Set ReportSheet = Worksheets("Report Page")
    Set RecordsSheet = Worksheets("Records Page")
    
    'Make sure there are activites on the Records Page
    If CheckRecords(RecordsSheet) > 2 Then
        GoTo NoActivities
    End If
    
    'Set RecordsLabelRange = FindRecordsLabel(RecordsSheet)
    Set AddRange = FindRecordsLabel(RecordsSheet)
    
    'Subtract any activities already present on the Report Page
    'If ReportSheet.ListObjects(1).ListRows.Count < 2 Then
        'GoTo SkipCompare
    'End If
    
    'Set ReportLabelRange = FindReportLabel(ReportSheet)
    
    'Taking this out for now. There needs to be a way to retabulate activities
    'For Each c In RecordsLabelRange
        'If ReportLabelRange.Find(c.Value, , xlValues, xlWhole) Is Nothing Then
            'If AddRange Is Nothing Then
                'Set AddRange = c
            'Else
                'Set AddRange = Union(AddRange, c)
            'End If
        'End If
    'Next c
    
    'If all activities are already tabulated
    If AddRange Is Nothing Then
        GoTo NoActivities
    End If
    
SkipCompare:
    'Unprotect and show userform
    Call UnprotectSheet(ReportSheet)
    TabulateActivityForm.Show
    
    GoTo Footer
    
NoActivities:
    'MsgBox ("You have no additional activities to tabulate")
    'GoTo Footer
    
Footer:
    Call ResetProtection 'This isn't working for some reason. Putting it at the end of TabulateActivity works, though
    ReportSheet.Activate
    
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

End Sub

Sub ReportTabulateAllButton()
'Calls the sub, here to control when screen updating happens

    Dim ReportSheet As Worksheet

    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    Set ReportSheet = Worksheets("Report Page")
    
    Call UnprotectSheet(ReportSheet)
    Call TabulateAllActivities
    
Footer:
    Call ResetProtection 'This isn't working for some reason. Putting it at the end of TabulateActivity works, though

    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

End Sub

Sub ReportTabulateTotalsButton()
'Calls the sub, here to control when screen updating happens

    Dim ReportSheet As Worksheet

    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    Set ReportSheet = Worksheets("Report Page")
    
    Call UnprotectSheet(ReportSheet)
    Call TabulateReportTotals
    
Footer:
    Call ResetProtection
    
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

End Sub
