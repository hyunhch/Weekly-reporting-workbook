VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} LoadActivityForm 
   Caption         =   "Load Activity"
   ClientHeight    =   4860
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

Private Sub Image1_BeforeDragOver(ByVal Cancel As MSForms.ReturnBoolean, ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal DragState As MSForms.fmDragState, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)

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
    
    Dim LabelString As String
    Dim i As Long
    
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    'Make sure an activity has been selected
    If LoadActivitySelectBox.ListIndex = -1 Then
        MsgBox ("Please select an activity")
        GoTo Footer
    End If
   
    For i = 0 To Me.LoadActivitySelectBox.ListCount - 1
        If Me.LoadActivitySelectBox.Selected(i) Then
            LabelString = Me.LoadActivitySelectBox.List(i, 0)
            Call LoadActivity(LabelString)
        End If
    Next i
        
    LoadActivityForm.Hide
   
Footer:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

End Sub

Private Sub LoadActivityDeleteAllButton_Click()
'Clears everything from the Records and Report sheets

    Dim DelConfirm As Long
    
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    'If nothing is saved, nothing needs to be done
    If Me.LoadActivitySelectBox.ListCount = 0 Then
        MsgBox ("You have no saved activities to delete.")
        GoTo Footer
    End If
    
    DelConfirm = MsgBox("Are you sure you want to delete these activities? " & vbCr & _
    "This cannot be undone.", vbQuestion + vbYesNo + vbDefaultButton2)
    
    If DelConfirm <> vbYes Then
        GoTo Footer
    End If
    
    'Delete from the Records page
    Call ClearRecords("Labels")
    
    'Delete from the Report page
    Call ClearReportButton(1)
    
    'In this case we'll leave any open Activity sheets
    
    'Hide the userform
    LoadActivityForm.Hide
    
Footer:

    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

End Sub

Private Sub LoadActivityDeleteButton_Click()
'Delete the selected activities, removing it from the attendance and label sheets
    
    Dim SelectedLabel As String
    Dim DelConfirm As Long
    Dim i As Long
    Dim j As Long
    
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    'Make sure an activity is selected
    If LoadActivitySelectBox.ListIndex = -1 Then
        MsgBox ("Please select an activity")
        GoTo Footer
    End If
    
    'Count if one or multiple are selected
    j = 0
    For i = 0 To Me.LoadActivitySelectBox.ListCount - 1
        If Me.LoadActivitySelectBox.Selected(i) Then
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
    j = Me.LoadActivitySelectBox.ListCount - 1
    For i = j To 0 Step -1
        If Me.LoadActivitySelectBox.Selected(i) Then
            SelectedLabel = Me.LoadActivitySelectBox.List(i, 0)
            Call DeleteActivity(SelectedLabel)
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

Private Sub LoadActivitySelectBox_Click()

End Sub

Private Sub UserForm_Activate()
'Populate the list box with all saved activities

    Dim RecordsSheet As Worksheet
    Dim LabelName As String
    Dim LabelPractice As String
    Dim LabelDate As Date
    Dim LabelRange As Range
    Dim c As Range
    Dim FCol As Long
    Dim LCol As Long
    Dim i As Long
    
    Set RecordsSheet = Worksheets("Records Page")
    
    'Clear out anything that's already in the list box
    If LoadActivitySelectBox.ListCount > 0 Then
        LoadActivitySelectBox.Clear
    End If
    
    'Make columns in the list box
    With LoadActivitySelectBox
        .ColumnCount = 3
        .ColumnWidths = "150, 150, 30"
    End With
    
    'Define the range with labels
    FCol = RecordsSheet.Range("1:1").Find("V BREAK", , xlValues, xlWhole).Column
    LCol = RecordsSheet.Range("1:1").Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Column
    
    'Check if anything has been saved. This should be taken care of previously
    If LCol = FCol Then
        MsgBox ("You have no saved activities.")
        LoadActivityForm.Hide
        GoTo Footer
    End If
    
    Set LabelRange = RecordsSheet.Range(Cells(1, FCol + 1).Address, Cells(1, LCol).Address)
    
    'Copy over the label information
    i = 0
    For Each c In LabelRange
        LabelName = c.Value
        LabelPractice = c.Offset(1, 0).Value
        LabelDate = CDate(c.Offset(2, 0).Value)
        
        With LoadActivitySelectBox
            .AddItem LabelName
            .List(i, 1) = LabelPractice
            .List(i, 2) = LabelDate
        End With
        i = i + 1
    Next c
    
Footer:

End Sub

