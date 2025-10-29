Attribute VB_Name = "ActivitySubs"
Option Explicit

Sub ActivityNewButtons(ActivitySheet As Worksheet)
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
        .OnAction = "ActivityDeleteButton"
        .Caption = "Delete Activity"
    End With
    
    'Save Activity button
    Set NewButtonRange = ActivitySheet.Range("G2:H3")
    Set NewButton = ActivitySheet.Buttons.Add(NewButtonRange.Left, NewButtonRange.Top, _
        NewButtonRange.Width, NewButtonRange.Height)
    
    With NewButton
        .OnAction = "ActivitySaveButton"
        .Caption = "Save Activity"
    End With
    
    'Close Activity button
    Set NewButtonRange = ActivitySheet.Range("G5:H5")
    Set NewButton = ActivitySheet.Buttons.Add(NewButtonRange.Left, NewButtonRange.Top, _
        NewButtonRange.Width, NewButtonRange.Height)
    
    With NewButton
        .OnAction = "ActivityCloseButton"
        .Caption = "Close Sheet"
    End With
    
    'Pull attendence button
    Set NewButtonRange = ActivitySheet.Range("E5:F5")
    Set NewButton = ActivitySheet.Buttons.Add(NewButtonRange.Left, NewButtonRange.Top, _
        NewButtonRange.Width, NewButtonRange.Height)
    
    With NewButton
        .OnAction = "ActivityPullAttendenceButton"
        .Caption = "Pull Attendence"
    End With

End Sub

Function ActivityNewSheet(InfoArray As Variant) As Worksheet
'Called from the new activity form, returns a completed activity sheet
'Activates the sheet if it's already open and ends the subroutine
'The array is 2D and contains the information to be inserted and where it's to be inserted
    '(1, i) -> Header
    '(2, i) -> Value
    '(3, i) -> Address

    Dim RosterSheet As Worksheet
    Dim RecordsSheet As Worksheet
    Dim ActivitySheet As Worksheet
    Dim RecordsLabelCell As Range
    Dim RecordsAttendanceRange As Range
    Dim RosterNameRange As Range
    Dim ActivityNameRange As Range
    Dim CopyRange As Range
    Dim PasteRange As Range
    Dim c As Range
    Dim d As Range
    Dim i As Long
    Dim LabelString As String
    Dim RosterTable As ListObject
    Dim ActivityTable As ListObject
    
    Set RosterSheet = Worksheets("Roster Page")
    Set RecordsSheet = Worksheets("Records Page")
    Set RosterTable = RosterSheet.ListObjects(1)
    
    'Grab the label
    For i = 1 To UBound(InfoArray)
        If InfoArray(1, i) = "Label" Then
            LabelString = InfoArray(2, i)
            
            If Not Len(LabelString) > 0 Then
                GoTo Footer
            End If
            
            GoTo LabelCheck
        End If
    Next i
    
LabelCheck:
    'First check if there is an activity sheet open with the label
    Set ActivitySheet = FindSheet(LabelString)
    
    If Not ActivitySheet Is Nothing Then
        'MsgBox ("Sheet already open")
        ActivitySheet.Activate
        
        GoTo Footer
    End If

    'If it isn't, create a new sheet at the end of the workbook, add activity information and buttons
    ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)).Name = LabelString
    Set ActivitySheet = Worksheets(LabelString)
    Set PasteRange = ActivitySheet.Range("A7")

    Call UnprotectSheet(ActivitySheet)
    Call ActivityNewText(ActivitySheet, InfoArray)
    Call ActivityNewButtons(ActivitySheet)
    
    'Make a blank table, required for copying students
    Set ActivityTable = MakeActivityTable(ActivitySheet)
    
    'If the activity already exists, load it from the RecordsSheet. Otherwise, pull in students checked on RosterSheet
    Set RecordsLabelCell = FindRecordsLabel(RecordsSheet, , LabelString)
    
    If RecordsLabelCell Is Nothing Then 'Label not found, pull from the roster
        Set CopyRange = FindChecks(RosterTable.ListColumns("Select").DataBodyRange)
    Else 'Label found
        Set RosterNameRange = RosterTable.ListColumns("First").DataBodyRange
        Set d = FindRecordsAttendance(RecordsSheet, , RecordsLabelCell)
             If d Is Nothing Then
                GoTo MakeTable
            End If

        Set RecordsAttendanceRange = FindStudentAttendance(RecordsSheet, d)
            If RecordsAttendanceRange Is Nothing Then
                GoTo MakeTable
            End If
            
        'Match names to the Roster
        Set c = FindName(RecordsAttendanceRange.Offset(0, -RecordsAttendanceRange.Column + 1), RosterNameRange)
        
            If c Is Nothing Then
                GoTo MakeTable
            End If
        
        Set CopyRange = c '.Offset(0, -1)
    End If
    
    'Label not found and no students checked, break. This shouldn't happen
    If CopyRange Is Nothing Then
        GoTo Footer
    End If
    
    'Copy from roster if it's a new activity or there are no matches
    If RecordsLabelCell Is Nothing Then
        Call CopyRow(RosterSheet, CopyRange, ActivitySheet, PasteRange)
    Else
        Call CopyFromRecords(ActivitySheet, RecordsLabelCell)
    End If
        
MakeTable:
    Set ActivityTable = MakeActivityTable(ActivitySheet)
    
    If ActivityTable.ListRows.Count < 1 Then 'This shouldn't happen
        GoTo ProtectSheet
    End If
    
    Call TableFormat(ActivitySheet, ActivityTable)
    
    'Clean up. These shouldn't be needed
    Set ActivityNameRange = ActivityTable.ListColumns("First").DataBodyRange
    Call RemoveBadRows(ActivitySheet, ActivityTable.DataBodyRange, ActivityNameRange)
    
ProtectSheet:
    'Apply protection to the first five rows and activate
    With ActivitySheet
       .Protect , userinterfaceonly:=True, AllowSorting:=True, AllowFiltering:=True, AllowFormattingColumns:=True
       .Cells.Locked = False
       .Range("A1:A5").EntireRow.Locked = True
    
       .Activate
    End With

    Set ActivityNewSheet = ActivitySheet

Footer:

End Function

Sub ActivityNewText(ActivitySheet As Worksheet, InfoArray As Variant)
'Puts in text and formatting on a new activity sheet
    '(1, i) -> Header
    '(2, i) -> Value
    '(3, i) -> Address
    
    Dim RefSheet As Worksheet
    Dim CategoryRange As Range
    Dim c As Range
    Dim i As Long
    Dim PracticeString As String
    Dim PracticeTable As ListObject

    Set RefSheet = Worksheets("Ref Tables")
    Set PracticeTable = RefSheet.ListObjects("ActivitiesTable")

    'Put in all text from the passed array. Everything in the fist x dimension is place at the address listed in the third. The second dimension goes directly to the right
    For i = 1 To UBound(InfoArray, 2)
        Set c = ActivitySheet.Range(InfoArray(3, i))
        
        With c
            .Value = InfoArray(1, i)
            .Font.Bold = True
            .HorizontalAlignment = xlRight
            .Offset(0, 1).Value = InfoArray(2, i)

            .Resize(1, 2).Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Resize(1, 2).Borders(xlEdgeBottom).Weight = xlMedium
            .Resize(1, 2).WrapText = False
        End With
        
        If InfoArray(1, i) = "Practice" Then
            PracticeString = InfoArray(2, i)
        ElseIf InfoArray(1, i) = "Category" Then
            Set CategoryRange = c.Offset(0, 1)
        End If
    Next i
    
    'Put in the category for the passed practice
    Set c = PracticeTable.ListColumns("Practice").DataBodyRange.Find(PracticeString, , xlValues, xlWhole) 'One column to the right of the category
    
    If Not c Is Nothing Then
        CategoryRange = c.Offset(0, -1).Value
    End If
    
    'Autofit the first column
    ActivitySheet.Range("A1").EntireColumn.AutoFit
    
Footer:

End Sub

Sub ActivityPullAllAttendance()
'Helper function that loops through all sheets and calls to repull attendance


End Sub

Sub ActivityPullAttendence(ActivitySheet As Worksheet, ActivityNameRange As Range, LabelCell As Range)
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
    Set d = FindStudentAttendance(RecordsSheet, RecordsAttendanceRange, "Present")
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

Sub ActivitySave(ActivitySheet As Worksheet, RecordsSheet As Worksheet, LabelCell As Range, Optional ShowMessage As String)
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






