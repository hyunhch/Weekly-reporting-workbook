VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} RemoveStudentsForm 
   Caption         =   "Remove Students"
   ClientHeight    =   4935
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6600
   OleObjectBlob   =   "RemoveStudentsForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "RemoveStudentsForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub SetArray(NameArray)
'Passing an array of names to populate the select box

    Me.RemoveStudentsSelectBox.List = NameArray
    
End Sub

Private Sub RemoveStudentsAllButton_Click()
'Delete all names in the listbox

    Dim SelectedNames() As String
    Dim i As Long
    Dim j As Long

    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    'Loop through the list
    j = 0
    For i = 0 To Me.RemoveStudentsSelectBox.ListCount - 1
        j = j + 1
        ReDim Preserve SelectedNames(1 To 2, 1 To j)
        SelectedNames(1, j) = Me.RemoveStudentsSelectBox.List(i, 0)
        SelectedNames(2, j) = Me.RemoveStudentsSelectBox.List(i, 1)
    Next i
    
    'Pass names to match and delete selected students
    Call RemoveFromRecords(SelectedNames, j)
    
    Me.Hide
    
Footer:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True


End Sub

Private Sub RemoveStudentsCancelButton_Click()

    Me.Hide

    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

End Sub

Private Sub RemoveStudentsConfirmButton_Click()
'Remove all students selected in the list box. Also retabulate all activities where those students were marked present

    Dim SelectedNames() As String
    Dim i As Long
    Dim j As Long

    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    'Make sure a student is selected
    If RemoveStudentsSelectBox.ListIndex = -1 Then
        MsgBox ("Please select one or more students")
        GoTo Footer
    End If
    
    'Loop through the list for selected names
    j = 0
    For i = 0 To Me.RemoveStudentsSelectBox.ListCount - 1
        If Me.RemoveStudentsSelectBox.Selected(i) Then
            j = j + 1
            ReDim Preserve SelectedNames(1 To 2, 1 To j)
            SelectedNames(1, j) = Me.RemoveStudentsSelectBox.List(i, 0)
            SelectedNames(2, j) = Me.RemoveStudentsSelectBox.List(i, 1)
        End If
    Next i
    
    'Pass names to match and delete selected students
    Call RemoveFromRecords(SelectedNames, j)
    
    Me.Hide
    
Footer:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

End Sub

Private Sub UserForm_Activate()
'Make the columns we need for the list box

    With RemoveStudentsSelectBox
        .ColumnCount = 2
        .ColumnWidths = "60, 60"
    End With

End Sub

