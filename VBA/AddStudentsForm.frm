VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AddStudentsForm 
   Caption         =   "Add Students"
   ClientHeight    =   5355
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8670.001
   OleObjectBlob   =   "AddStudentsForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "AddStudentsForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Image1_BeforeDragOver(ByVal Cancel As MSForms.ReturnBoolean, ByVal Data As MSForms.DataObject, ByVal x As Single, ByVal Y As Single, ByVal DragState As MSForms.fmDragState, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)

End Sub

Private Sub AddStudentsCancelButton_Click()
'Hide the form

    AddStudentsForm.Hide

    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

End Sub

Private Sub AddStudentsConfirmButton_Click()
'Recreate an activity sheet with the activity information and attendance
    
    Dim ActivitySheet As Worksheet
    Dim RosterSheet As Worksheet
    Dim RosterNameRange As Range
    Dim RosterCheckedRange As Range
    Dim ActivityNameRange As Range
    Dim AddRange As Range
    Dim c As Range
    Dim LabelString As String
    Dim PracticeString As String
    Dim DateString As String
    Dim DescriptionString As String
    'Dim AddedStudents As String
    Dim i As Long
    Dim InfoArray() As Variant
    Dim RosterTable As ListObject
    
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    'Make sure an activity has been selected
    If AddStudentsListBox.ListIndex = -1 Then
        MsgBox ("Please select an activity")
        GoTo Footer
    End If
    
    Set RosterSheet = Worksheets("Roster Page")
    
    'Grab activity information
    For i = 0 To Me.AddStudentsListBox.ListCount - 1
        If Me.AddStudentsListBox.Selected(i) Then
            LabelString = Me.AddStudentsListBox.List(i, 0)
            PracticeString = Me.AddStudentsListBox.List(i, 1)
            DateString = Me.AddStudentsListBox.List(i, 2)
            DescriptionString = Me.AddStudentsListBox.List(i, 3)
            
            GoTo MakeArray 'shouldn't be necessary since it's a single-select listbox
        End If
    Next i
    
MakeArray:
    'Populate an array for later
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
    
    'Check to see if there is an open sheet
    Set ActivitySheet = FindSheet(LabelString)
    
    If ActivitySheet Is Nothing Then
        Set ActivitySheet = NewActivitySheet(InfoArray)
    End If

DuplicateCheck:
    Set RosterTable = RosterSheet.ListObjects(1)
    Set RosterNameRange = RosterTable.ListColumns("First").DataBodyRange
    Set RosterCheckedRange = FindChecks(RosterNameRange.Offset(0, -1)) 'Ensuring there's a checked student happens in the parent sub
    
    'If there are no existing students, copy them all over
    Set ActivityNameRange = ActivitySheet.ListObjects(1).ListColumns("First").DataBodyRange
    
    If ActivityNameRange Is Nothing Then
        Set AddRange = RosterCheckedRange.Offset(0, 1)

        GoTo AddNew
    End If
    
    'Find all checked students not already on the ActivitySheet
    Set AddRange = FindUnique(RosterCheckedRange.Offset(0, 1), ActivityNameRange)
    
    If AddRange Is Nothing Then
        MsgBox ("All checked students were already added to the activity.")
        
        GoTo Footer
    End If
    
    'Listing individual students gets too unwiedly
    'For Each c In AddRange
        'AddedStudents = AddedStudents & c.Offset(0, 1) & " " & c.Offset(0, 2) & vbCr
    'Next c
    
    'MsgBox ("The following students have been added:" & vbCr & AddedStudents)
    
AddNew:
    'Add non-duplicative students and list those added
    Call CopyToActivity(RosterSheet, ActivitySheet, AddRange)
    ActivitySheet.Activate
    
    'Reset checks on the RosterSheet
    RosterCheckedRange.Value = "a"
    
    'Show how many were added
    If RosterCheckedRange.Cells.Count = AddRange.Cells.Count Then
        MsgBox ("All selected students added.")
    Else
        MsgBox (AddRange.Cells.Count & " students added.")
    End If
   
Footer:
    AddStudentsForm.Hide
    
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

End Sub

Private Sub AddStudentsFilterTextBox_Change()
'Dynamic filter for the activity list

    Dim i As Long
    Dim testString As String
    
    testString = LCase("*" & AddStudentsFilterTextBox.Text & "*")
    Call UserForm_Activate
    
    With AddStudentsListBox
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
    If AddStudentsListBox.ListCount > 0 Then
        AddStudentsListBox.Clear
    End If
    
    'Make columns in the list box
    With AddStudentsListBox
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
        With AddStudentsListBox
            .AddItem c.Value
            .List(i, 1) = RecordsSheet.Cells(PracticeRange.Row, c.Column)
            .List(i, 2) = CDate(RecordsSheet.Cells(DateRange.Row, c.Column))
            .List(i, 3) = RecordsSheet.Cells(DescriptionRange.Row, c.Column)
        End With
        
        i = i + 1
    Next c
    
Footer:

End Sub


