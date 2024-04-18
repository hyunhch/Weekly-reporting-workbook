Attribute VB_Name = "ActivityButtonSubs"
Option Explicit

Sub DeleteActivityButton()
'Deletes the sheet and any record of it

    Dim ActivitySheet As Worksheet
    Dim RecordsSheet As Worksheet
    Dim ReportSheet As Worksheet
    Dim LabelString As String
    Dim DelConfirm As Long
    
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    'First give a warning
    DelConfirm = MsgBox("Are you sure you want to delete this activity? " & vbCr & _
    "This cannot be undone.", vbQuestion + vbYesNo + vbDefaultButton2)

    If DelConfirm <> vbYes Then
        GoTo Footer
    End If
    
    Set ActivitySheet = ActiveSheet
    Set RecordsSheet = Worksheets("Records Page")
    Set ReportSheet = Worksheets("Report Page")
    LabelString = ActivitySheet.Range("H1").Value

    'Search the Records sheet
    Dim LabelMatch As Range
    
    Set LabelMatch = FindLabel(LabelString)
    
    If Not LabelMatch Is Nothing Then
        LabelMatch.EntireColumn.Delete
    End If
    
    'Search the Report sheet
    Dim LabelHeader As Range
    Dim LabelFooter As Range
    Dim SearchRange As Range
    
    Set LabelHeader = FindReportRange("Label")
    Set LabelFooter = LabelHeader.EntireColumn.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    Set SearchRange = ReportSheet.Range(LabelHeader.Offset(2, 0), LabelFooter)
    Set LabelMatch = SearchRange.Find(LabelString, , xlValues, xlWhole)
    
    If Not LabelMatch Is Nothing Then
        LabelMatch.EntireRow.Delete
    End If
    
    'Delete the activity sheet
    ActiveSheet.Delete

Footer:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

End Sub

Sub CloseActivityButton()
'Deletes the sheet, but keeps record of the activity

    Dim ActivitySheet As Worksheet
    Dim LabelString As String
    Dim SaveConfirm As Long

    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    Set ActivitySheet = ActiveSheet
    LabelString = ActivitySheet.Range("H1")
    
    'Ask if the activity should be saved before closing
    SaveConfirm = MsgBox("Would you like to save this activity before closing the sheet?", vbQuestion + vbYesNo + vbDefaultButton2)
    If SaveConfirm = vbYes Then
        If SaveActivity = False Then
            GoTo Footer
        End If
    End If

    ActiveSheet.Delete

Footer:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

End Sub

Sub SaveActivityButton()
'Only here to call the SaveActivity() function
    
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    'Make sure the totals pulled. The tabulation will go into the wrong row otherwise
    Call PullReportTotals
    
    If SaveActivity = True Then
        Call SaveActivity
        MsgBox ("Activity saved.")
    End If
    
Footer:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

End Sub


