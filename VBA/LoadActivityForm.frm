VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} LoadActivityForm 
   Caption         =   "Load Activity"
   ClientHeight    =   5352
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8670.001
   OleObjectBlob   =   "LoadActivityForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "LoadActivityForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Image1_BeforeDragOver(ByVal Cancel As MSForms.ReturnBoolean, ByVal Data As MSForms.DataObject, ByVal x As Single, ByVal Y As Single, ByVal DragState As MSForms.fmDragState, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)

End Sub

Private Sub LoadActivityCancelButton_Click()
'Hide the form

    LoadActivityForm.Hide

    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

End Sub

Private Sub LoadActivityConfirmButton_Click()
'Recreate an activity sheet with the activity information and attendance
    
    Dim RecordsSheet As Worksheet
    Dim ActivitySheet As Worksheet
    Dim LabelString As String
    Dim PracticeString As String
    Dim DateString As String
    Dim DescriptionString As String
    Dim i As Long
    Dim InfoArray() As Variant
    
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    'Make sure an activity has been selected
    If LoadActivityListBox.ListIndex = -1 Then
        'MsgBox ("Please select an activity")
        GoTo Footer
    End If

    Set RecordsSheet = Worksheets("Records Page")

    For i = 0 To Me.LoadActivityListBox.ListCount - 1
        If Me.LoadActivityListBox.Selected(i) Then
            LabelString = Me.LoadActivityListBox.List(i, 0)
            PracticeString = Me.LoadActivityListBox.List(i, 1)
            DateString = Me.LoadActivityListBox.List(i, 2)
            DescriptionString = Me.LoadActivityListBox.List(i, 3)
            
            'Check if the sheet is already open
            Set ActivitySheet = FindSheet(LabelString)
            
            If Not ActivitySheet Is Nothing Then
                ActivitySheet.Activate
                GoTo NextActivity
            End If
            
            'Create the array to pass
            ReDim InfoArray(1 To 5, 1 To 3)
            
            InfoArray(1, 1) = "Label"
            InfoArray(2, 1) = "Practice"
            InfoArray(3, 1) = "Category"
            InfoArray(4, 1) = "Date"
            InfoArray(5, 1) = "Description"
            
            InfoArray(1, 2) = "G1"
            InfoArray(2, 2) = "A1"
            InfoArray(3, 2) = "A2"
            InfoArray(4, 2) = "A3"
            InfoArray(5, 2) = "A4"
            
            InfoArray(1, 3) = LabelString
            InfoArray(2, 3) = PracticeString
            InfoArray(3, 3) = "" 'Will be overwritten
            InfoArray(4, 3) = CDate(DateString)
            InfoArray(5, 3) = DescriptionString
            
            'Make a new sheet and copy over attendance information
            Set ActivitySheet = NewActivitySheet(InfoArray)
            Call CopyFromRecords(RecordsSheet, ActivitySheet, FindActivityLabel(ActivitySheet))
        End If
NextActivity:
    Next i
        
    LoadActivityForm.Hide
   
Footer:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

End Sub

Private Sub LoadActivityDeleteAllButton_Click()
'Clears everything from the Records and Report sheets

    Dim RecordsSheet As Worksheet
    Dim ActivitySheet As Worksheet
    Dim ActivityLabelRange As Range
    Dim DelConfirm As Long
    
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    Set RecordsSheet = Worksheets("Records Page")
    
    'If nothing is saved, nothing needs to be done
    If Me.LoadActivityListBox.ListCount = 0 Then
        GoTo Footer
    End If
    
    DelConfirm = MsgBox("Are you sure you want to delete these activities? " & vbCr & _
    "This cannot be undone.", vbQuestion + vbYesNo + vbDefaultButton2)
    
    If DelConfirm <> vbYes Then
        GoTo Footer
    End If
    
    'Delete from the Records page if there are any activities. There always should be
    If CheckRecords(RecordsSheet) > 2 Then
        GoTo ClearReport
    End If
    
    Set ActivityLabelRange = FindRecordsLabel(RecordsSheet)
    ActivityLabelRange.EntireColumn.Delete
    
ClearReport:
    'Delete from the Report page
    Call ClearReportButton
    
    'The previous sub turns these back on
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    'In this case we'll leave any open Activity sheets
    For Each ActivitySheet In ThisWorkbook.Sheets
        If ActivitySheet.Range("A1").Value = "Practice" Then
            ActivitySheet.Delete
        End If
    Next ActivitySheet
    
    'Hide the userform
    LoadActivityForm.Hide
    
Footer:

    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

End Sub

Private Sub LoadActivityDeleteButton_Click()
'Delete the selected activities, removing it from the attendance and label sheets
    
    Dim RecordsSheet As Worksheet
    Dim TempLabelRange As Range
    Dim DelConfirm As Long
    Dim i As Long
    Dim j As Long
    Dim SelectedLabel As String
    
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    Set RecordsSheet = Worksheets("Records Page")
    Set TempLabelRange = RecordsSheet.Range("A1")

    'Make sure an activity is selected
    If LoadActivityListBox.ListIndex = -1 Then
        MsgBox ("Please select an activity")
        GoTo Footer
    End If
    
    'Count if one or multiple are selected
    j = 0
    For i = 0 To Me.LoadActivityListBox.ListCount - 1
        If Me.LoadActivityListBox.Selected(i) Then
            j = j + 1
        End If
    Next i
    
    'Give a warning
    If j = 1 Then
        DelConfirm = MsgBox("Are you sure you want to delete this activity? " & vbCr & _
        "This cannot be undone.", vbQuestion + vbYesNo + vbDefaultButton2)
    Else
        DelConfirm = MsgBox("Are you sure you want to delete these activities? " & vbCr & _
        "This cannot be undone.", vbQuestion + vbYesNo + vbDefaultButton2)
    End If
    
    If DelConfirm <> vbYes Then
        GoTo Footer
    End If
    
    'Loop throughthe listbox and delete all selected item
    j = Me.LoadActivityListBox.ListCount - 1
    For i = j To 0 Step -1
        If Me.LoadActivityListBox.Selected(i) Then
            SelectedLabel = Me.LoadActivityListBox.List(i, 0)
            TempLabelRange.Value = SelectedLabel
            
            Call DeleteActivity(RecordsSheet, TempLabelRange)
            TempLabelRange.ClearContents
        End If
    Next i

ListboxRemove:
    'Refresh the listbox items
    Call UserForm_Activate
    
Footer:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

End Sub

Private Sub LoadActivityFilterTextBox_Change()
'Dynamic filter for the activity list

    Dim i As Long
    Dim testString As String
    
    testString = LCase("*" & LoadActivityFilterTextBox.Text & "*")
    Call UserForm_Activate
    
    With LoadActivityListBox
        For i = .ListCount - 1 To 0 Step -1
            If (Not (LCase(.List(i, 0)) Like testString)) _
            And (Not (LCase(.List(i, 1)) Like testString)) _
            And (Not (LCase(.List(i, 2)) Like testString)) _
            Then
                .RemoveItem i
            End If
        Next i
    End With
    
End Sub

Private Sub UserForm_Activate()
'Populate the list box with all saved activities

    Dim RecordsSheet As Worksheet
    Dim LabelRange As Range
    Dim PracticeRange As Range
    Dim DateRange As Range
    Dim DescriptionRange As Range
    Dim LabelHeaderRange As Range
    Dim c As Range
    Dim i As Long
 
    Set RecordsSheet = Worksheets("Records Page")
    
    'Clear out anything that's already in the list box
    If LoadActivityListBox.ListCount > 0 Then
        LoadActivityListBox.Clear
    End If
    
    'Make columns in the list box
    With LoadActivityListBox
        .ColumnCount = 3
        .ColumnWidths = "150, 150, 30, 0"
    End With
    
    'Checking that there are activities happens in parent sub
    Set LabelRange = FindRecordsLabel(RecordsSheet)
    
    'If all activities are deleted, it will pull in the padding cell
    If LabelRange.Cells.Count = 1 Then
        If LabelRange.Value = "V BREAK" Then
            GoTo Footer
        End If
    End If
    
    'Find where the values we need are
    Set LabelHeaderRange = FindRecordsActivityHeaders(RecordsSheet)
    Set PracticeRange = LabelHeaderRange.Find("Practice", , xlValues, xlWhole)
    Set DateRange = LabelHeaderRange.Find("Date", , xlValues, xlWhole)
    Set DescriptionRange = LabelHeaderRange.Find("Description", , xlValues, xlWhole)
    
    'Copy over the label information
    i = 0
    For Each c In LabelRange
        With LoadActivityListBox
            .AddItem c.Value
            .List(i, 1) = RecordsSheet.Cells(PracticeRange.Row, c.Column)
            .List(i, 2) = CDate(RecordsSheet.Cells(DateRange.Row, c.Column))
            .List(i, 3) = RecordsSheet.Cells(DescriptionRange.Row, c.Column)
        End With
        
        i = i + 1
    Next c
    
Footer:

End Sub


