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

Private Sub CommandButton1_Click()

    Dim i As Long
    Dim InfoArray As Variant
    
    InfoArray = GetFormInfo(NewActivityForm)
    
    For i = 1 To UBound(InfoArray, 2)
    
        Debug.Print InfoArray(1, i) & " - " & InfoArray(2, i) & " - " & InfoArray(3, i)
    
    Next i


End Sub

Private Sub NewActivityCancelButton_Click()

    NewActivityForm.Hide

    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

End Sub

Private Sub NewActivityConfirmButton_Click()
'Create a new sheet with the information given
'Checking for students and checked students comes previously
'Passes a 3xi array
    '(1, i) - Header
    '(2, i) - Value
    '(3, i) - Address

    Dim RecordsSheet As Worksheet
    Dim MatchCell As Range
    Dim i As Long
    Dim PracticeString As String
    Dim CategoryString As String
    Dim DateString As String
    Dim LabelString As String
    Dim DescriptionString As String
    Dim HeaderArray As Variant
    Dim AddressArray As Variant
    Dim PassArray As Variant
    
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    Set RecordsSheet = Worksheets("Records Page")
    
    'Make sure a practice is selected
    If NewActivitySelectListBox.ListIndex = -1 Then
        MsgBox ("Please select a practice")
        GoTo Footer
    End If
    
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
    LabelString = NewActivityLabelBox.Value
    
    Set MatchCell = FindRecordsLabel(RecordsSheet, , LabelString)
    
    'Function returns nothing if no match was found, the padding cell if there are not activities
    If Not MatchCell Is Nothing Then
        If Not MatchCell.Value = "V BREAK" Then
            MsgBox ("All labels must be unique. Please choose a different one")
            GoTo Footer
        End If
    End If
        
    'Grab the headers and values
    PassArray = GetFormInfo(NewActivityForm)
        If IsEmpty(PassArray) Or Not IsArray(PassArray) Then
            GoTo Footer
        End If
        
    'Create a new sheet
    Call ActivityNewSheet(PassArray)

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

Private Sub NewActivityLabelBox_Exit(ByVal Cancel As MSForms.ReturnBoolean)
'Check for label length and invalid characters
'Must be 31 characters or fewer, not contain : \ / ? * [ or ]

    Dim i As Long
    Dim LabelString As String
    Dim InvalidArray As Variant
    
    Me.NewActivityLabelBox.BackColor = RGB(255, 255, 255)
    
    'If there is nothing, break
    If Not Len(NewActivityLabelBox.Value) > 0 Then
        GoTo Footer
    End If
    
    'Check length
    LabelString = NewActivityLabelBox.Value
    
    If Len(LabelString) > 31 Then
        MsgBox ("Labels can only be 31 characters or shorter")
        Me.NewActivityLabelBox.BackColor = RGB(255, 198, 198)
        
        GoTo Footer
    End If
    
    'Check for invalid characters
    InvalidArray = Split(": \ / ? * [ ]", " ")
    
    For i = LBound(InvalidArray) To UBound(InvalidArray)
        If InStr(1, LabelString, InvalidArray(i)) > 0 Then
            MsgBox ("Labels cannot use any of the following characters: " & vbCr _
                & ": \ / ? * [ or ]")
            Me.NewActivityLabelBox.BackColor = RGB(255, 198, 198)
        
            GoTo Footer
        End If
    Next i
        
Footer:

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

Private Function NewActivityGetInfo() As Variant
'Returns a 2D array with the information used to create a new ActivitySheet
    '(1, i) - header
    '(2, i) - value
    
    Dim TempRange As Range
    Dim c As Range
    Dim i As Long
    Dim LabelString As String
    Dim PracticeString As String
    Dim CategoryString As String
    Dim DateString As String
    Dim DescriptionString As String
    Dim ReturnArray As Variant

    'Grab the values for headers
    Set TempRange = Range("ActivityHeadersList")
        
    ReDim ReturnArray(1 To 2, 1 To TempRange.Cells.Count)
    
    i = 1
    For Each c In TempRange
        ReturnArray(1, i) = c.Value
    
        i = i + 1
    Next c
    
    'Read in the values from the userform
    PracticeString = NewActivitySelectListBox.Value
    DateString = NewActivityDateBox.Value
    LabelString = NewActivityLabelBox.Value
    DescriptionString = NewActivityDescriptionBox.Value
    
    'Find the category associated with the practice
    Set TempRange = Range("ActivitiesList")
    Set c = TempRange.Find(PracticeString, , xlValues, xlWhole)
        If c Is Nothing Then
            GoTo Footer
        End If

    CategoryString = c.Offset(0, -1).Value

AddValues:
    ReturnArray(2, 1) = LabelString
    ReturnArray(2, 2) = PracticeString
    ReturnArray(2, 3) = CategoryString
    ReturnArray(2, 4) = CDate(DateString)
    ReturnArray(2, 5) = DescriptionString

    'Return
    NewActivityGetInfo = ReturnArray

Footer:

End Function

