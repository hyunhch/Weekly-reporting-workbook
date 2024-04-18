VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} NewActivityForm 
   Caption         =   "New Activity Form"
   ClientHeight    =   4860
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8040
   OleObjectBlob   =   "NewActivityForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "NewActivityForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub NewActivityCancelButton_Click()

    NewActivityForm.Hide

End Sub

Private Sub NewActivityConfirmButton_Click()
'Create a new sheet with the information given
'Checking for students and checked students comes previously

    'First check that all three fields on the form have a value
    Dim PracticeCheck As String
    Dim DateCheck As String
    Dim LabelCheck As String
    Dim DescriptionCheck As String
    
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    PracticeCheck = NewActivityPracticeBox.Value
    DateCheck = NewActivityDateBox.Value
    LabelCheck = NewActivityLabelBox.Value
    DescriptionCheck = NewActivityDescriptionBox.Value
    
    If Len(Trim(PracticeCheck)) = 0 Then
        MsgBox ("Please select an activity")
        GoTo Footer
    ElseIf Len(Trim(DateCheck)) = 0 Then
        MsgBox ("Please enter a date")
        GoTo Footer
    ElseIf Len(Trim(LabelCheck)) = 0 Then
        MsgBox ("Please enter a label for the activity")
        GoTo Footer
    ElseIf Len(Trim(DescriptionCheck)) = 0 Then
        MsgBox ("Please briefly describe the activity")
        GoTo Footer
    End If
        
    'Make sure the label is unique
    Dim RecordsSheet As Worksheet
    Dim MatchCell As Range
    
    Set RecordsSheet = Worksheets("Records Page")
    Set MatchCell = FindLabel(NewActivityLabelBox.Value)
    If MatchCell.Value = "V BREAK" Then
        MsgBox ("All labels must be unique. Please choose a different one")
        GoTo Footer
    End If
        
    'Create a new sheet
    Call NewActivitySheet(PracticeCheck, CDate(DateCheck), LabelCheck, DescriptionCheck, Checked)

    NewActivityForm.Hide
    
Footer:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

End Sub

Private Sub NewActivityDateBox_Exit(ByVal Cancel As MSForms.ReturnBoolean)
'Require a date to be entered

    'Skip if empty
    If Len(Trim(NewActivityDateBox.Value)) = 0 Then
        GoTo EmptyCheck
    End If

    If Not IsDate(NewActivityDateBox.Text) Then
        MsgBox ("Please enter a date")
        Cancel = True
    End If
        
    'For testing
    'NewActivityDescriptionBox.Text = Format(CDate(NewActivityDateBox.Text), "dd/mm/yyyy")
    
EmptyCheck:

End Sub

Private Sub UserForm_Activate()
'Clear anything in the date and description boxes when activated

    NewActivityPracticeBox = ""
    NewActivityDateBox.Value = ""
    NewActivityLabelBox.Value = ""
    NewActivityDescriptionBox.Value = ""

End Sub

Private Sub UserForm_Initialize()
'Populate the activities dropdown list the first time it's opened

    Dim RefSheet As Worksheet
    Dim c As Range
    
    For Each c In Range("ActivitiesList")
        Me.NewActivityPracticeBox.AddItem c.Value
    Next c
    
End Sub



