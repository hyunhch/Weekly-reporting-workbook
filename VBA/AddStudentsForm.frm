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
    Dim CopyRange As Range
    Dim PasteRange As Range
    Dim c As Range
    Dim LabelString As String
    Dim i As Long
    Dim InfoArray() As Variant
    Dim RosterTable As ListObject
    Dim ActivityTable As ListObject
    
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    Set RosterSheet = Worksheets("Roster Page")
    Set RosterTable = RosterSheet.ListObjects(1)
    Set RosterNameRange = RosterTable.ListColumns("First").DataBodyRange

    'Make sure an activity has been selected
    If AddStudentsListBox.ListIndex = -1 Then
        MsgBox ("Please select an activity")
        GoTo Footer
    End If
    
    Set RosterSheet = Worksheets("Roster Page")
    Set RosterCheckedRange = FindChecks(RosterTable.ListColumns("Select").DataBodyRange)
        If RosterCheckedRange Is Nothing Then
            GoTo Footer
        End If
    
    'Loop through
    For i = 0 To Me.AddStudentsListBox.ListCount - 1
        If Me.AddStudentsListBox.Selected(i) Then
            GoTo GetInfo
        End If
    Next i
    
    GoTo Footer
    
GetInfo:
    'Grab activity information
    InfoArray = GetFormInfo(AddStudentsForm, i)
    LabelString = Me.AddStudentsListBox.Value

    'Check to see if there is an open sheet
    Set ActivitySheet = FindSheet(LabelString)
        If ActivitySheet Is Nothing Then
            Set ActivitySheet = ActivityNewSheet(InfoArray)
        ElseIf Not ActivitySheet.ListObjects.Count > 0 Then 'Make sure there is a table
            Call MakeTable(ActivitySheet)
        End If

    Set ActivityTable = ActivitySheet.ListObjects(1)

DuplicateCheck:
    'Check if there are any students already on the ActivitySheet
    If CheckTable(ActivitySheet) > 2 Then
        Set CopyRange = RosterNameRange 'Copy over everyone
        
        GoTo AddNew
    End If

    'Find all checked students on the RosterTable and compare to existing students on the ActivityTable
    Set ActivityNameRange = ActivityTable.ListColumns("First").DataBodyRange
    
    Set CopyRange = FindUnique(RosterCheckedRange.Offset(0, 1), ActivityNameRange) '.Offset(0, -1)
        If CopyRange Is Nothing Then
            MsgBox ("All checked students were already added to the activity.")
            
            GoTo Footer
        End If
    
AddNew:
    'Find the last row on the ActivitySheet
    Set PasteRange = FindLastRow(ActivitySheet, "Select").Offset(1, 0)
        If PasteRange Is Nothing Then
            GoTo Footer
        End If
        
    'Add non-duplicative students and list those added and remake the table
    Call CopyRow(RosterSheet, CopyRange, ActivitySheet, PasteRange)
    
    Set ActivityTable = MakeActivityTable(ActivitySheet)
    Call TableFormat(ActivitySheet, ActivityTable)
    Call ActivityPullAttendenceButton(ActivitySheet)
    
    'Reset checks on the RosterSheet
    RosterCheckedRange.Value = "a"
    
    'Show how many were added
    If RosterCheckedRange.Cells.Count = CopyRange.Cells.Count Then
        MsgBox ("All selected students added.")
    Else
        MsgBox (CopyRange.Cells.Count & " students added.")
    End If
    
    ActivitySheet.Activate
   
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
    Dim LabelArray As Variant
 
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
    
    'Put into an array and pass to fill the listbox
    ReDim LabelArray(1 To LabelRange.Cells.Count)
    
    i = 1
    For Each c In LabelRange
        LabelArray(i) = c.Value
        
        i = i + 1
    Next c
    
    Call PopulateListBox(AddStudentsForm, Me.AddStudentsListBox, LabelArray)
    
Footer:

End Sub


