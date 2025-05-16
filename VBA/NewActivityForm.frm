VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} NewActivityForm 
   Caption         =   "New Activity Form"
   ClientHeight    =   7335
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7215
   OleObjectBlob   =   "NewActivityForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "NewActivityForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Image1_BeforeDragOver(ByVal Cancel As MSForms.ReturnBoolean, ByVal Data As MSForms.DataObject, ByVal x As Single, ByVal Y As Single, ByVal DragState As MSForms.fmDragState, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)

End Sub

Private Sub NewActivityCancelButton_Click()

    NewActivityForm.Hide

End Sub

Private Sub NewActivityConfirmButton_Click()
'Create a new sheet with the information given
'Checking for students and checked students comes previously

    Dim RecordsSheet As Worksheet
    Dim MatchCell As Range
    Dim TempCell As Range
    Dim PracticeString As String
    Dim CategoryString As String
    Dim DateString As String
    Dim LabelString As String
    Dim DescriptionString As String
    Dim InfoArary() As Variant
    
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    'First check that all four fields on the form have a value
    PracticeString = NewActivitySelectListBox.Value
    DateString = NewActivityDateBox.Value
    LabelString = NewActivityLabelBox.Value
    DescriptionString = NewActivityDescriptionBox.Value
    
    If Len(Trim(PracticeString)) = 0 Then
        MsgBox ("Please select an activity")
        GoTo Footer
    ElseIf Not IsDate(DateString) Then
        MsgBox ("Please enter a date in the form mm/dd/yyyy")
        GoTo Footer
    ElseIf Len(Trim(LabelString)) = 0 Then
        MsgBox ("Please enter a label for the activity")
        GoTo Footer
    ElseIf Len(Trim(DescriptionString)) = 0 Then
        MsgBox ("Please briefly describe the activity")
        GoTo Footer
    End If
        
    'Make sure the label is unique
    Set RecordsSheet = Worksheets("Records Page")
    Set TempCell = RecordsSheet.Range("B2")
    
    TempCell.Value = NewActivityLabelBox.Value
    Set MatchCell = FindRecordsLabel(RecordsSheet, TempCell)
    TempCell.ClearContents
    
    'Function returns nothing if no match was found, the padding cell if there are not activities
    If Not MatchCell Is Nothing Then
        If Not MatchCell.Value = "V BREAK" Then
            MsgBox ("All labels must be unique. Please choose a different one")
            GoTo Footer
        End If
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
        
    'Create a new sheet
    Call NewActivitySheet(InfoArray)

    NewActivityForm.Hide
    
Footer:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

End Sub

Private Sub NewActivityFilterBox_Change()
'Dynamic filter for the activity list

    Dim i As Long
    Dim testString As String
    
    testString = LCase("*" & NewActivityFilterBox.Text & "*")
    NewActivitySelectListBox.Clear
    Call UserForm_Initialize
    
    With NewActivitySelectListBox
        For i = .ListCount - 1 To 0 Step -1
            If (Not (LCase(.List(i, 0)) Like testString)) Then
                .RemoveItem i
            End If
        Next i
    End With
    
End Sub

Private Sub UserForm_Activate()
'Clear anything in the date and description boxes when activated

    NewActivityFilterBox.Value = ""
    NewActivityDateBox.Value = ""
    NewActivityLabelBox.Value = ""
    NewActivityDescriptionBox.Value = ""

    NewActivityForm.Height = 395
    NewActivityForm.Width = 371

End Sub

Private Sub UserForm_Initialize()
'Populate the activities dropdown list the first time it's opened

    Dim RefSheet As Worksheet
    Dim c As Range
    
    NewActivitySelectListBox.Clear
    
    For Each c In Range("ActivitiesList")
        Me.NewActivitySelectListBox.AddItem c.Value
    Next c
    
End Sub



