VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ChooseProgramForm 
   Caption         =   "Select Program"
   ClientHeight    =   3435
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   4455
   OleObjectBlob   =   "ChooseProgramForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ChooseProgramForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ChooseProgramNoButton_Click()
    
    ChooseProgramForm.Hide
    
    If Application.Workbooks.Count = 1 Then
        Application.Quit
    Else
        ActiveWorkbook.Close SaveChanges:=False
    End If

End Sub

Private Sub ChooseProgramYesButton_Click()
'Set up worksheet for selected program

    Dim SelectionString As String
    Dim ProgramString As String
    Dim i As Long
    
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    'Make sure an activity has been selected
    If ChooseProgramListBox.ListIndex = -1 Then
        MsgBox ("Please select an activity")
        GoTo Footer
    End If
    
    For i = 0 To Me.ChooseProgramListBox.ListCount - 1
        If Me.ChooseProgramListBox.Selected(i) Then
            SelectionString = Me.ChooseProgramListBox.List(i, 0)
        End If
    Next i
    
    'Change to the strings we need
    If SelectionString = "College Prep" Then
        ProgramString = "College Ref"
    ElseIf SelectionString = "Transfer Prep" Then
        ProgramString = "Transfer Ref"
    ElseIf SelectionString = "MESA University" Then
        ProgramString = "University Ref"
    End If
    
    'Pass to the setup sub and close the form
    Call ChooseProgram(ProgramString)
    
    ChooseProgramForm.Hide

    'Bring up the form to enter submitter information
    EnterInfoForm.Show

Footer:

    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

End Sub

Private Sub UserForm_Initialize()

    ChooseProgramForm.Height = 205
    ChooseProgramForm.Width = 230
    
    ChooseProgramListBox.List = Array("College Prep", "Transfer Prep", "MESA University")

End Sub
