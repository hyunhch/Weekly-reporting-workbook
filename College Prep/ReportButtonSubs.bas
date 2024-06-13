Attribute VB_Name = "ReportButtonSubs"
Option Explicit

Sub PullReportTotalsButton()
'Calls PullReportTotals()

    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    'First make sure the roster is parsed, then pull. Tabulation looks at a table object on the Roster sheet
    Call ReadRosterButton
    Call PullReportTotals
    
Footer:
    Call ResetProtection
    
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

End Sub

Sub ClearReportButton(Optional ShowWarning As Long)
'Resets the Report sheet

    Dim ReportSheet As Worksheet
    Dim LRow As Long
    Dim DelConfirm As Long
    
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    Set ReportSheet = Worksheets("Report Page")
    
    'Confirm the deletion
    If ShowWarning <> 1 Then
        DelConfirm = MsgBox("Do you wish to clear all content on this sheet?" & vbCr & _
        "This cannot be undone.", vbQuestion + vbYesNo + vbDefaultButton2)
        If DelConfirm <> vbYes Then
            GoTo Footer
        End If
    End If
    
    'Make sure there is at least one activity. We can skip the entire function if there isn't
    If CheckTableLength(ReportSheet, ReportSheet.Range("B:B").Find("Center", , xlValues, xlWhole)) = 0 Then
        GoTo Footer
    End If
    
    'Unprotect
    Call UnprotectCheck(ReportSheet)
    
    'Find the data range and clear contents
    'We aren't using ClearSheet() because it will clear formatting
    Dim StartCell As Range
    Dim EndCell As Range
    
    Set StartCell = FindReportRange("Select").Offset(1, 0)
    Set EndCell = ReportSheet.Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    ReportSheet.Range(StartCell, EndCell).EntireRow.ClearContents
    
    'Reprotect
    Call ResetProtection
    
Footer:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
End Sub

Sub TabulateButton()
'Brings up the TabulateActivityForm userform

    Dim ReportSheet As Worksheet
    Dim RecordsSheet As Worksheet
    Dim FRow As Long
    Dim LRow As Long
    Dim FCol As Long
    Dim LCol As Long
    
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    Set ReportSheet = Worksheets("Report Page")
    Set RecordsSheet = Worksheets("Records Page")
    
    'Unprotect the Report page
    Call UnprotectCheck(ReportSheet)
    
    'Make sure there are students on the Records sheet
    FRow = RecordsSheet.Range("A:A").Find("H BREAK", , xlValues, xlWhole).Row
    LRow = RecordsSheet.Range("A:A").Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    
    If FRow = LRow Then
        MsgBox ("Something has gone wrong. Please parse the roster, save the activity, and try again.")
        GoTo Footer
    End If
    
    'Make sure there are any saved activities
    FCol = RecordsSheet.Range("1:1").Find("V BREAK", , xlValues, xlWhole).Column
    LCol = RecordsSheet.Range("1:1").Find("*", SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
    
    If FCol = LCol Then
        MsgBox ("You don't have any saved activities." & vbCr & "Please save an activity and try again.")
        GoTo Footer
    End If

    'Make sure the totals are tabulated
    Call ReadRosterButton
    Call PullReportTotals
    
    'Display the userform
    TabulateActivityForm.Show
    
Footer:
    Call ResetProtection

    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
End Sub
