Attribute VB_Name = "ActivitySubs"
Option Explicit

Sub RangeBuildTest()
    
    Dim RosterSheet As Worksheet
    Dim PasteSheet As Worksheet
    Dim c As Range
    Dim d As Range
    Dim CheckRange As Range
    Dim UnionRange As Range
    Dim CopyRange As Range
    Dim PasteRange As Range
    Dim RosterTable As ListObject
    Dim PasteTable As ListObject
    
    
    Set PasteSheet = ActiveSheet
    Set RosterSheet = Worksheets("Roster Page")
    Set RosterTable = RosterSheet.ListObjects(1)
    Set PasteRange = PasteSheet.Range("A7")
    Set CheckRange = FindChecks(RosterTable.ListColumns("Select").DataBodyRange)
    
    For Each c In CheckRange.Offset(0, 1)
        Set d = c.Resize(1, RosterTable.ListColumns.Count - 1)
        Set CopyRange = BuildRange(d, CopyRange)
    Next c
    
    Call CopyRows(RosterSheet, CopyRange, PasteSheet, PasteRange)


End Sub

Function NewActivitySheet(InfoArray() As Variant) As Worksheet
'Called from the new activity form, returns a completed activity sheet
'Activates the sheet if it's already open and ends the subroutine
'The array is 2D and contains the information to be inserted and where it's to be inserted
'(1, 1) -> {What1}
'(1, 2) -> {Address1}
'(1, 3) -> {Value1}, etc.

    Dim RosterSheet As Worksheet
    Dim RecordsSheet As Worksheet
    Dim ActivitySheet As Worksheet
    Dim ActivityNameRange As Range
    Dim RosterCopyRange As Range
    Dim RecordsLabelCell As Range
    Dim ActivityLabelCell As Range
    Dim CopyRange As Range
    Dim PasteRange As Range
    Dim DelRange As Range
    Dim c As Range
    Dim d As Range
    Dim i As Long
    Dim LabelString As String
    Dim HeaderArray() As Variant
    Dim RosterTable As ListObject
    Dim ActivityTable As ListObject

    Set RosterSheet = Worksheets("Roster Page")
    Set RecordsSheet = Worksheets("Records Page")
    
    'Grab the label
    For i = 1 To UBound(InfoArray)
        If InfoArray(i, 1) = "Label" Then
            LabelString = InfoArray(i, 3)
            GoTo LabelCheck
        End If
    Next i
    
LabelCheck:
    'First check if there is an activity sheet open with the label
    Set ActivitySheet = FindSheet(LabelString)
    
    If Not ActivitySheet Is Nothing Then
        'MsgBox ("Sheet already open")
        Set NewActivitySheet = ActivitySheet
        ActivitySheet.Activate
        
        GoTo Footer
    End If

    'If it isn't, create a new sheet at the end of the workbook, add activity information and buttons
    ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)).Name = LabelString
    Set ActivitySheet = Worksheets(LabelString)
    Set PasteRange = ActivitySheet.Range("B7")
    Set RosterTable = RosterSheet.ListObjects(1)
    
    Call UnprotectSheet(ActivitySheet)
    Call NewActivityText(ActivitySheet, InfoArray)
    Call NewActivityButtons(ActivitySheet)
    
    'Make a blank table, required for copying students
    Set ActivityTable = CreateActivityTable(ActivitySheet)
    Set ActivityLabelCell = FindActivityLabel(ActivitySheet)
    
    'If the activity already exists, load it from the RecordsSheet. Otherwise, pull in students checked on RosterSheet
    Set RecordsLabelCell = FindRecordsLabel(RecordsSheet, ActivityLabelCell)
    
    If RecordsLabelCell Is Nothing Then
        Set c = FindChecks(RosterTable.ListColumns("Select").DataBodyRange)
        Set RosterCopyRange = c.Offset(0, 1)
    ElseIf RecordsLabelCell.Value = "V BREAK" Then
        Set c = FindChecks(RosterTable.ListColumns("Select").DataBodyRange)
        Set RosterCopyRange = c.Offset(0, 1)
    Else
        Set c = CopyFromRecords(RecordsSheet, ActivitySheet, ActivityLabelCell)
        Set RosterCopyRange = FindName(c, RosterTable.ListColumns("First").DataBodyRange)
    End If

    'Redefine the range to copy as the entire table row and copy over
    For Each c In RosterCopyRange
        Set d = c.Resize(1, RosterTable.ListColumns.Count - 1)
        Set CopyRange = BuildRange(d, CopyRange)
    Next c
    
    Call CopyRows(RosterSheet, CopyRange, ActivitySheet, PasteRange)
    
    'Make a table
    Set ActivityTable = CreateActivityTable(ActivitySheet)
    
    If ActivityTable.ListRows.Count < 1 Then 'This shouldn't happen
        GoTo ProtectSheet
    End If
    
    Call FormatTable(ActivitySheet, ActivityTable)
    
    'Clean up. These shouldn't be needed
    Set ActivityNameRange = ActivityTable.ListColumns("First").DataBodyRange
    Set c = FindDuplicate(ActivityNameRange)
    Set d = FindBlanks(ActivitySheet, ActivityNameRange)
    
    If Not c Is Nothing Then
        Set DelRange = c
    End If
    
    If Not d Is Nothing Then
        If DelRange Is Nothing Then
            Set DelRange = d
        Else
            Set DelRange = Union(DelRange, d)
        End If
    End If
    
    If Not DelRange Is Nothing Then
        Call RemoveRows(ActivitySheet, ActivityTable.DataBodyRange, ActivityNameRange, DelRange)
    End If

    'Save the activity. This can create an empty activity on the Records Sheet [Disabling this for now. I'm not sure how useful it will be]
    'Set ActivityLabelCell = ActivitySheet.Range("1:1").Find(LabelString, , xlValues, xlWhole)
    
    'Call SaveActivity(ActivitySheet, RecordsSheet, ActivityLabelCell)
    
ProtectSheet:
    'Apply protection to the first five rows and activate
    With ActivitySheet
       .Protect , userinterfaceonly:=True, AllowSorting:=True, AllowFiltering:=True, AllowFormattingColumns:=True
       .Cells.Locked = False
       .Range("A1:A5").EntireRow.Locked = True
    
       .Activate
    End With

    Set NewActivitySheet = ActivitySheet

Footer:

End Function

Sub NewActivityText(ActivitySheet As Worksheet, InfoArray() As Variant)
'Puts in text and formatting on a new activity sheet

    Dim RefSheet As Worksheet
    Dim c As Range
    Dim i As Long
    Dim PracticeString As String
    Dim PracticeTable As ListObject

    Set RefSheet = Worksheets("Ref Tables")
    Set PracticeTable = RefSheet.ListObjects("ActivitiesTable")

    'Put in all text from the passed array
    For i = 1 To UBound(InfoArray)
        Set c = ActivitySheet.Range(InfoArray(i, 2))
        
        With c
            .Value = InfoArray(i, 1)
            .Font.Bold = True
            .HorizontalAlignment = xlRight
            .Offset(0, 1).Value = InfoArray(i, 3)

            .Resize(1, 2).Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Resize(1, 2).Borders(xlEdgeBottom).Weight = xlMedium
            .Resize(1, 2).WrapText = False
        End With
        
        If InfoArray(i, 1) = "Practice" Then
            PracticeString = InfoArray(i, 3)
        End If
    Next i
    
    'Put in the category for the passed practice
    Set c = PracticeTable.ListColumns("Practice").DataBodyRange.Find(PracticeString, , xlValues, xlWhole) 'One column to the right of the category
    
    If Not c Is Nothing Then
        ActivitySheet.Range("A:A").Find("Category", , xlValues, xlWhole).Offset(0, 1).Value = c.Offset(0, -1).Value
    End If
    
    'Autofit the first column
    ActivitySheet.Range("A1").EntireColumn.AutoFit
    
Footer:

End Sub

Sub SaveActivity(ActivitySheet As Worksheet, RecordsSheet As Worksheet, LabelCell As Range, Optional ShowMessage As String)
'Saves the indicated activity. Called when an activity is first made and when it's manually saved
'Silent by default, shows a message if passed

    Dim RecordsNameRange As Range
    Dim RecordsLabelRange As Range
    Dim RecordsAttendanceRange As Range
    Dim ActivityNameRange As Range
    Dim ActivitySearchRange As Range
    Dim c As Range
    Dim d As Range
    Dim i As Long
    Dim j As Long
    Dim ActivityTable As ListObject
    
    'Checking that there are students and tables should be done in parent sub
    Set ActivityTable = ActivitySheet.ListObjects(1)
    Set ActivityNameRange = ActivityTable.ListColumns("First").DataBodyRange
    Set RecordsNameRange = FindRecordsName(RecordsSheet)
    Set RecordsLabelRange = FindRecordsLabel(RecordsSheet, LabelCell)
    
    If RecordsLabelRange Is Nothing Then 'The label hasn't been recorded yet, so place one after the last filled column
        Set RecordsLabelRange = RecordsSheet.Range("1:1").Find("*", SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Offset(0, 1)
    ElseIf RecordsLabelRange.Value = "V BREAK" Then 'No labels at all
        Set RecordsLabelRange = RecordsLabelRange.Offset(0, 1)
    End If
    
    If ActivityNameRange Is Nothing Then
        GoTo TabulateActivity
    End If
    
    'Enter in or update activity information. The headers are listed in order on the Ref Tables sheet
    j = Range("ActivityHeadersList").Rows.Count
    Set ActivitySearchRange = ActivitySheet.Range("A1", LabelCell.Offset(j - 1))
    
    For i = 1 To j
        Set c = Range("ActivityHeadersList").Rows(i)
        Set d = ActivitySearchRange.Find(c.Value, , xlValues, xlWhole).Offset(0, 1)
    
        RecordsLabelRange.Offset(i - 1, 0).Value = d.Value
    Next i
    
    'Clear all existing attendance information and update
    Set RecordsAttendanceRange = FindRecordsAttendance(RecordsSheet, , LabelCell)
    
    RecordsAttendanceRange.ClearContents
    For Each c In ActivityNameRange
        Set d = FindName(c, RecordsNameRange)
        If Not d Is Nothing Then
            If c.Offset(0, -1).Value = "a" Then
                d.Offset(0, RecordsLabelRange.Column - 1).Value = 1
            Else
                d.Offset(0, RecordsLabelRange.Column - 1).Value = 0
            End If
        End If
    Next c
    
TabulateActivity: 'Consider removing and moving up to parent subs
    'Pass to retabulate the activity
    Call TabulateActivity(LabelCell)
    
    'Confirmation message
    If ShowMessage = "Yes" Then
        MsgBox ("Activity saved.")
    End If

Footer:

End Sub

Sub DeleteActivity(RecordsSheet As Worksheet, LabelCell As Range)
'Deletes the activity from the Records Page, removes it from the Report Page, and deletes any open Activity Sheet

    Dim ReportSheet As Worksheet
    Dim ActivitySheet As Worksheet
    Dim RecordsLabelRange As Range
    Dim ReportLabelRange As Range
    Dim ActivityLabelRange As Range
    
    Set ReportSheet = Worksheets("Report Page")
    
    'Make sure the label exists on the Records sheet
    Set RecordsLabelRange = FindRecordsLabel(RecordsSheet, LabelCell)
    
    If RecordsLabelRange Is Nothing Then
        GoTo CheckReport
    ElseIf RecordsLabelRange.Value <> LabelCell.Value Then
        GoTo CheckReport
    End If
    
    RecordsLabelRange.EntireColumn.Delete
    
CheckReport:
    'Make sure the label exists on the Report sheet
    Set ReportLabelRange = FindReportLabel(ReportSheet, LabelCell)
    
    If ReportLabelRange Is Nothing Then
        GoTo CheckSheets
    ElseIf ReportLabelRange.Value <> LabelCell.Value Then
        GoTo CheckSheets
    End If
    
    ReportLabelRange.EntireRow.Delete

CheckSheets:
    'If there is an open sheet, delete it
    For Each ActivitySheet In ThisWorkbook.Sheets
        If ActivitySheet.Range("A1").Value = "Practice" And Not ActivitySheet.Range("1:1").Find(LabelCell.Value, , xlValues, xlWhole) Is Nothing Then
            ActivitySheet.Delete
        End If
    Next ActivitySheet
    
Footer:

End Sub

Sub NewActivityButtons(ActivitySheet As Worksheet)
'Called when an activity is created or loaded

    Dim NewButton As Button
    Dim NewButtonRange As Range

    'Select All
    Set NewButtonRange = ActivitySheet.Range("A5:B5")
    Set NewButton = ActivitySheet.Buttons.Add(NewButtonRange.Left, NewButtonRange.Top, _
        NewButtonRange.Width, NewButtonRange.Height)
    
    With NewButton
        .OnAction = "SelectAllButton"
        .Caption = "Select All"
    End With

    'Delete Row
    Set NewButtonRange = ActivitySheet.Range("C5:D5")
    Set NewButton = ActivitySheet.Buttons.Add(NewButtonRange.Left, NewButtonRange.Top, _
        NewButtonRange.Width, NewButtonRange.Height)
    
    With NewButton
        .OnAction = "RemoveSelectedButton"
        .Caption = "Delete Row"
    End With

    'Delete activity button
    Set NewButtonRange = ActivitySheet.Range("J2:K2")
    Set NewButton = ActivitySheet.Buttons.Add(NewButtonRange.Left, NewButtonRange.Top, _
        NewButtonRange.Width, NewButtonRange.Height)
    
    With NewButton
        .OnAction = "DeleteActivityButton"
        .Caption = "Delete Activity"
    End With
    
    'Save Activity button
    Set NewButtonRange = ActivitySheet.Range("G2:H3")
    Set NewButton = ActivitySheet.Buttons.Add(NewButtonRange.Left, NewButtonRange.Top, _
        NewButtonRange.Width, NewButtonRange.Height)
    
    With NewButton
        .OnAction = "SaveActivityButton"
        .Caption = "Save Activity"
    End With
    
    'Close Activity button
    Set NewButtonRange = ActivitySheet.Range("G5:H5")
    Set NewButton = ActivitySheet.Buttons.Add(NewButtonRange.Left, NewButtonRange.Top, _
        NewButtonRange.Width, NewButtonRange.Height)
    
    With NewButton
        .OnAction = "CloseActivityButton"
        .Caption = "Close Sheet"
    End With
    
    'Pull attendence button
    Set NewButtonRange = ActivitySheet.Range("E5:F5")
    Set NewButton = ActivitySheet.Buttons.Add(NewButtonRange.Left, NewButtonRange.Top, _
        NewButtonRange.Width, NewButtonRange.Height)
    
    With NewButton
        .OnAction = "PullAttendenceButton"
        .Caption = "Pull Attendence"
    End With

End Sub


