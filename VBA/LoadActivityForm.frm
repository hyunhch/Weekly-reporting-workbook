VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} LoadActivityForm 
   Caption         =   "Load Activity"
   ClientHeight    =   5355
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
    Dim HeaderRefRange As Range
    Dim LabelCell As Range
    Dim c As Range
    Dim LabelString As String
    Dim PracticeString As String
    Dim DateString As String
    Dim DescriptionString As String
    Dim i As Long
    Dim RowIndex As Long
    Dim InfoArray() As Variant
    
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    'Make sure an activity has been selected
    If LoadActivityListBox.ListIndex = -1 Then
        GoTo Footer
    End If

    'Loop through
    For i = 0 To Me.LoadActivityListBox.ListCount - 1
        If Not Me.LoadActivityListBox.Selected(i) Then
            GoTo NextActivity
        End If

    'Get activity information to pass
    InfoArray = GetFormInfo(LoadActivityForm, i)
        If IsEmpty(InfoArray) Or Not IsArray(InfoArray) Then
            GoTo Footer
        End If
        
    'Make a new sheet and pull in attendance
    Set ActivitySheet = ActivityNewSheet(InfoArray)
        If ActivitySheet Is Nothing Then
            GoTo Footer
        End If
        
        Call ActivityPullAttendenceButton(ActivitySheet)
NextActivity:
    Next i
     
    LoadActivityForm.Hide
   
Footer:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

End Sub

Function LoadActivityGetInfo(RowIndex As Long) As Variant
'Prepares a 3 x i array for passing when creating a new activity sheet
    '(1, i) - Header
    '(2, i) - Value
    '(3, i) - Address

    Dim HeaderRefRange As Range
    Dim c As Range
    Dim i As Long
    Dim CategoryString As String
    Dim PracticeString As String
    Dim ReturnArray As Variant
    Dim AddressArray As Variant
    Dim TempArray As Variant
    Dim ValueArray As Variant

    

    'Grab the activity headers from the RefSheet
    Set HeaderRefRange = Range("ActivityHeadersList")
        If HeaderRefRange Is Nothing Then
            GoTo Footer
        End If


    'Make an array to pass
    ReDim ReturnArray(1 To 3, 1 To HeaderRefRange.Cells.Count)
    
    i = 1
    For Each c In HeaderRefRange
        ReturnArray(1, i) = c.Value
        
        i = i + 1
    Next c
    
    'Grab values from the form and where they go. In the future, make this programmatic
    TempArray = Split("G1,A1,A2,A3,A4", ",") 'This is in base 0
    
    ReDim AddressArray(1 To UBound(TempArray) + 1)
    
    For i = 1 To UBound(AddressArray)
        AddressArray(i) = TempArray(i - 1)
    Next i
    
    'Grab values
    ReDim ValueArray(1 To HeaderRefRange.Cells.Count)
    
    'This starts with an index of 0
    For i = 1 To HeaderRefRange.Cells.Count
        ValueArray(i) = Me.LoadActivityListBox.List(RowIndex, i - 1)
    Next i
    
    'Category isn't here. The cell will be overwritten later and the array needs to have a blank inserted
    i = UBound(ValueArray)
    
    'Make programmatic
    ValueArray(i) = ValueArray(i - 1)
    ValueArray(i - 1) = ValueArray(i - 2)
    ValueArray(i - 2) = ""
    
    'Put them all together
    For i = 1 To UBound(ReturnArray, 2) 'This should be exactly the same for all three
        'This doesn't grab the category
        If ReturnArray(1, i) = "Practice" Then 'This should always come before the category
            PracticeString = ValueArray(i)
            
            Set c = Range("ActivitiesList").Find(PracticeString, , xlValues, xlWhole)
            
            If Not c Is Nothing Then
                CategoryString = c.Offset(0, -1).Value
            End If
        End If

        ReturnArray(2, i) = ValueArray(i)
        ReturnArray(3, i) = AddressArray(i)
        
        'Overwrite category
        If ReturnArray(1, i) = "Category" Then
            ReturnArray(2, i) = CategoryString
        End If
    Next i
        
    'Return
    If IsEmpty(ReturnArray) Or Not IsArray(ReturnArray) Then
        GoTo Footer
    End If

    LoadActivityGetInfo = ReturnArray

Footer:

End Function

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
    Call ReportClearButton
    
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
    Call LoadActivityCancelButton_Click
    
Footer:

End Sub

Private Sub LoadActivityDeleteButton_Click()
'Delete the selected activities, removing it from the attendance and label sheets
    
    Dim RecordsSheet As Worksheet
    Dim LabelCell As Range
    Dim DelConfirm As Long
    Dim DeletePrompt As String
    Dim i As Long
    Dim j As Long
    Dim SelectedLabel As String
    
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    Set RecordsSheet = Worksheets("Records Page")

    'Make sure an activity is selected
    If LoadActivityListBox.ListIndex = -1 Then
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
        DeletePrompt = "Are you sure you want to delete this activity? " & vbCr & _
        "This cannot be undone."
    ElseIf j > 1 Then
        DeletePrompt = "Are you sure you want to delete these activities? " & vbCr & _
        "This cannot be undone."
    End If

    DelConfirm = MsgBox(DeletePrompt, vbQuestion + vbYesNo + vbDefaultButton2)
        If DelConfirm <> vbYes Then
            GoTo Footer
        End If
    
    'Loop throughthe listbox and delete all selected item
    j = Me.LoadActivityListBox.ListCount - 1
    For i = j To 0 Step -1
        If Not Me.LoadActivityListBox.Selected(i) Then
            GoTo NextRow
        End If
        
        'Find the activity on the RecordsSheet
        SelectedLabel = Me.LoadActivityListBox.List(i, 0)
        Set LabelCell = FindRecordsLabel(RecordsSheet, , SelectedLabel)
        
        If Not LabelCell Is Nothing Then
            Call RemoveRecordsActivity(RecordsSheet, LabelCell)
        End If
NextRow:
    Next i

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
    Dim ReportSheet As Worksheet
    Dim LabelCell As Range
    Dim LabelRange As Range
    Dim LabelHeaderRange As Range
    Dim c As Range
    Dim i As Long
    Dim j As Long
    Dim PracticeString As String
    Dim DateString As String
    Dim DescriptionString As String
    Dim HeaderArray() As Variant
 
    Set RecordsSheet = Worksheets("Records Page")
    Set ReportSheet = Worksheets("Report Page")
    
    'Clear out anything that's already in the list box
    If LoadActivityListBox.ListCount > 0 Then
        LoadActivityListBox.Clear
    End If
    
    'Make columns in the list box
    With LoadActivityListBox
        .ColumnCount = 3
        .ColumnWidths = "150, 150, 30, 0"
    End With
    
    'Populate the listbox
    If LoadActivityPopulate <> 1 Then
        'Call LoadActivityCancelButton_Click
    End If
    
Footer:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True


End Sub

Function LoadActivityPopulate() As Long
'Grabs the activities that are not open and puts them into the listbox
'Returns 1 on succeess, 0 otherwise

    Dim RecordsSheet As Worksheet
    Dim RecordsLabelRange As Range
    Dim c As Range
    Dim i As Long
    Dim j As Long
    Dim LabelString As String
    Dim ValueArray As Variant
    
    LoadActivityPopulate = 0
    
    Set RecordsSheet = Worksheets("Records Page")
    
    'Ensure that there are activities to pull in
    If CheckRecords(RecordsSheet) <> 1 Then
        GoTo Footer
    End If

    'Define the list of activities
    Set RecordsLabelRange = FindRecordsLabel(RecordsSheet)

    ReDim LabelArray(1 To RecordsLabelRange.Cells.Count)
    
    i = 1
    j = 0
    For Each c In RecordsLabelRange
        LabelString = c.Value
        
        'Loop through open sheets. If there is a matching activity sheet, leave the label blank
        If Not FindSheet(LabelString) Is Nothing Then
            LabelArray(i) = ""
        Else
            LabelArray(i) = LabelString
            
            j = j + 1
        End If
        
        i = i + 1
    Next c
    
    'Break if we don't have anything
    If Not j > 0 Then
        GoTo Footer
    End If
    
    'Populate
    Call PopulateListBox(LoadActivityForm, Me.LoadActivityListBox, LabelArray)
    
    LoadActivityPopulate = 1

Footer:

End Function
