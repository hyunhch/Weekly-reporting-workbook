Attribute VB_Name = "ActivityButtonSubs"
Option Explicit

Sub ActivitySaveButton()
'To call the ActivitySave() suc

    Dim ActivitySheet As Worksheet
    Dim RecordsSheet As Worksheet
    Dim LabelCell As Range
    Dim ActivityTable As ListObject

    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    Set ActivitySheet = ActiveSheet
    Set RecordsSheet = Worksheets("Records Page")

    'Check that there is a table. There should always be
    If ActivitySheet.ListObjects.Count < 1 Then
        MsgBox ("Something has gone wrong. Please close this activity and either load or recreate it.")
        GoTo Footer
    End If
    
    'Check that there are students
    Set ActivityTable = ActivitySheet.ListObjects(1)
    If ActivityTable.ListRows.Count < 1 Then
        GoTo Footer
    End If
    
    'Check that the label is present. It always should be
    Set LabelCell = ActivitySheet.Range("1:1").Find("Label", , xlValues, xlWhole).Offset(0, 1)
    If LabelCell Is Nothing Or Len(LabelCell.Value) < 1 Then
        MsgBox ("Something has gone wrong. Please close this activity and either load or recreate it.")
        GoTo Footer
    End If
    
    'Check that there are students on the Records Page. It's okay if there are no activities
    If CheckRecords(RecordsSheet) = 2 Or CheckRecords(RecordsSheet) = 4 Then
        GoTo Footer
    End If
    
    'Make sure the Report totals are tabulated. The activity can go into the wrong row otherwise
    Call TabulateReportTotals

    'Pass to save
    Call ActivitySave(ActivitySheet, RecordsSheet, LabelCell, "Yes")

Footer:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

End Sub

Sub ActivityCloseButton()
'Deletes the sheet, but keeps record of the activity
'Only prompts to save if there is a difference between the sheet and the Records Sheet

    Dim ActivitySheet As Worksheet
    Dim RecordsSheet As Worksheet
    Dim ActivityLabelCell As Range
    Dim RecordsLabelCell As Range
    Dim ActivityNameRange As Range
    Dim RecordsNameRange As Range
    Dim TempAttendanceRange As Range
    Dim RecordsAttendanceRange As Range
    Dim c As Range
    Dim d As Range
    Dim SaveCheck As Long
    Dim ActivityTable As ListObject
    
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    Set ActivitySheet = ActiveSheet
    Set RecordsSheet = Worksheets("Records Page")
    
    Call UnprotectSheet(ActivitySheet)
    
    'Check if there's a table, students, and label. Close without saving if these aren't present
    If CheckTable(ActivitySheet) > 2 Then
        GoTo DeleteSheet
    End If
    
    Set ActivityTable = ActivitySheet.ListObjects(1)
    Set ActivityLabelCell = FindActivityLabel(ActivitySheet)
    
    If ActivityLabelCell Is Nothing Then
        GoTo DeleteSheet
    End If
    
    'Check if there are students on the Records Page. Prompt to parse the roster if there isn't
    If CheckRecords(RecordsSheet) = 2 Or CheckRecords(RecordsSheet) = 4 Then
        'MsgBox ("Please parse the roster and try again.")
        GoTo DeleteSheet
    End If
    
    'Check if the activity has been saved at all. Prompt if it hasn't
    Set RecordsLabelCell = FindRecordsLabel(RecordsSheet, ActivityLabelCell)
    
    If RecordsLabelCell Is Nothing Then
        GoTo SavePrompt
    End If
    
    'Compare the attendance information on the activity sheet and Records Page. Prompt if different
    Set RecordsAttendanceRange = FindRecordsAttendance(RecordsSheet, , ActivityLabelCell)
    Set RecordsNameRange = FindRecordsName(RecordsSheet)
    Set ActivityNameRange = ActivityTable.ListColumns("First").DataBodyRange
    
    For Each c In ActivityNameRange
        Set d = FindName(c, RecordsNameRange)
        If Not d Is Nothing Then
            If d.Offset(0, RecordsLabelCell.Column - 1) = 0 And c.Offset(0, -1) = "a" Then
                GoTo SavePrompt
            ElseIf d.Offset(0, RecordsLabelCell.Column - 1) = 1 And c.Offset(0, -1) <> "a" Then
                GoTo SavePrompt
            End If
        End If
    Next c
    
    'Everything was the same, so to deletion
    GoTo DeleteSheet
  
SavePrompt:
        SaveCheck = MsgBox("You have unsaved changes." & vbCr & _
            "Would you like to save this activity before closing the sheet?", vbQuestion + vbYesNo + vbDefaultButton2)
            
        If SaveCheck = vbYes Then
            Call ActivitySaveButton
        End If

DeleteSheet:
    'These are being enabled somewhere along the line
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    ActivitySheet.Delete

Footer:
    Call ResetProtection

    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
End Sub

Sub ActivityPullAttendenceButton(Optional PassSheet As Worksheet)
'Reproduces attendence as stored on the Records sheet

    Dim ActivitySheet As Worksheet
    Dim RecordsSheet As Worksheet
    Dim ActivityNameRange As Range
    Dim ActivityLabelCell As Range
    Dim RecordsLabelCell As Range
    Dim i As Long
    Dim ActivityTable As ListObject
    
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    Set RecordsSheet = Worksheets("Records Page")
    
    If PassSheet Is Nothing Then
        Set ActivitySheet = ActiveSheet
    Else
        Set ActivitySheet = PassSheet
    End If
    
    Call UnprotectSheet(ActivitySheet)
    
    'Check if there's a table, students, and label. Close without saving if these aren't present
    If CheckTable(ActivitySheet) > 2 Then
        GoTo ErrorMessage
    End If
    
    Set ActivityTable = ActivitySheet.ListObjects(1)
    Set ActivityLabelCell = FindActivityLabel(ActivitySheet)
        If ActivityLabelCell Is Nothing Then
            GoTo ErrorMessage
        End If

    'Check if there are students and activities
    If CheckRecords(RecordsSheet) <> 1 Then
        GoTo Footer
    End If
    
    'Check if the activity has been saved
    Set RecordsLabelCell = FindRecordsLabel(RecordsSheet, ActivityLabelCell)
        If RecordsLabelCell Is Nothing Then
            GoTo Footer
        End If

    'Pass to pull in attendence. All existing attendence will be cleared first
    Set ActivityNameRange = ActivityTable.ListColumns("First").DataBodyRange
    
    Call ActivityPullAttendence(ActivitySheet, ActivityNameRange, ActivityLabelCell)
    GoTo Footer

ErrorMessage:
    MsgBox ("Something has gone wrong. Please close this activity and recreate it.")

Footer:
    Call ResetProtection

    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
End Sub

Sub ActivityDeleteButton()
'Deletes the activity and removes it from the Records and Report sheets

    Dim ActivitySheet As Worksheet
    Dim RecordsSheet As Worksheet
    Dim ReportSheet As Worksheet
    Dim ActivityLabelCell As Range
    Dim RecordsLabelCell As Range
    Dim ReportLabelCell As Range
    Dim DelConfirm As Long
    
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    Set ActivitySheet = ActiveSheet
    Set RecordsSheet = Worksheets("Records Page")
    Set ReportSheet = Worksheets("Report Page")
    Set ActivityLabelCell = FindActivityLabel(ActivitySheet)
        If ActivityLabelCell Is Nothing Then
            GoTo Footer
        End If
    
    'Confirm deletion
    DelConfirm = MsgBox("This activity will be permanently deleted. Do you wish to continue?", vbQuestion + vbYesNo + vbDefaultButton2)
    If DelConfirm <> vbYes Then
        GoTo Footer
    End If
    
    'Remove from the Records sheet. This also removes from the Report and closes the Activity sheet
    Call UnprotectSheet(ActivitySheet)
    Call RemoveRecordsActivity(RecordsSheet, ActivityLabelCell)

    
Footer:
    Call ResetProtection

    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
End Sub
