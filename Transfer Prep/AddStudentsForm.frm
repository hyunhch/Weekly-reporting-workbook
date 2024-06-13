VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AddStudentsForm 
   Caption         =   "Add Selected Students"
   ClientHeight    =   4830
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8655.001
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
'Hides the userform

    AddStudentsForm.Hide

    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

End Sub

Private Sub AddStudentsConfirmButton_Click()

    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    'Make sure an activity is selected
    If AddStudentsSelectBox.ListIndex = -1 Then
        MsgBox ("Please select an activity")
        GoTo Footer
    End If

    'Make sure the roster is parsed
    If Worksheets("Roster Page").ListObjects.Count < 1 Then
        MsgBox ("Please parse the roster and try again.")
        GoTo Footer
    End If

    'Grab the label and find it on the Records sheet
    Dim LabelMatch As Range
    Dim LabelString As String
    
    LabelString = Me.AddStudentsSelectBox.List(AddStudentsSelectBox.ListIndex, 0)
    Set LabelMatch = FindLabel(LabelString, "RecordsSheet")
    
    'If the label can't be found. This shouldn't happen
    If LabelMatch Is Nothing Then
        MsgBox ("Something has gone wrong. Please save the activity and try again.")
        GoTo Footer
    End If
    
    'See if the activity sheet is open
    Dim sh As Worksheet
    Dim ActivitySheet As Worksheet
    Dim ActivityTableStart As Range
    
    For Each sh In ActiveWorkbook.Sheets
        If sh.Range("H1").Value = LabelString Then
            Set ActivitySheet = sh
            Set ActivityTableStart = ActivitySheet.Range("A:A").Find("Select", , xlValues, xlWhole)
            GoTo SearchNames
        End If
    Next sh
    
    'Load the activity into a new sheet
    Call LoadActivity(LabelString)
    Set ActivitySheet = ActiveWorkbook.Sheets(ActiveWorkbook.Sheets.Count)
    
SearchNames:
    'Make sure there are students on the activity sheet
    Set ActivityTableStart = ActivitySheet.Range("A:A").Find("Select", , xlValues, xlWhole)
    If Not CheckTableLength(ActivitySheet, ActivityTableStart) > 0 Then
        
        Call CopySelected(ActivitySheet, "Checked")
        
        'If it's an empty activity, which can happen if all participating students are removed from the roster
        If CheckTableLength(ActivitySheet, ActivityTableStart) > 0 Then
            Call UnprotectCheck(ActivitySheet)
            ActivitySheet.ListObjects(1).Unlist
            Call TableCreate(ActivitySheet, ActivityTableStart)
            Call AddMarlettBox(ActivitySheet.ListObjects(1).ListColumns("Select").DataBodyRange, ActivitySheet)
            Call ResetProtection
        End If
        
        ActivitySheet.Activate
        AddStudentsForm.Hide
        MsgBox ("All selected students added.")
        GoTo Footer
    End If

    'Grab the selected names from the Roster sheet and Activity sheet
    Dim RosterSheet As Worksheet
    Dim RosterChecks As Range
    Dim RosterNames As Range
    Dim ActivityNames As Range
    Dim PasteCell As Range
    Dim CopyCell As Range
    Dim NewStudents As String
    Dim LRow As Long
    
    Set RosterSheet = Worksheets("Roster Page")
    Set RosterChecks = RosterSheet.ListObjects("RosterTable").ListColumns("Select").DataBodyRange.SpecialCells(xlCellTypeVisible)
    Set RosterNames = FindChecks(RosterChecks)
    Set ActivityNames = ActivitySheet.ListObjects(1).ListColumns("First").DataBodyRange
    
    'Unprotect
    Call UnprotectCheck(ActivitySheet)
    
    'Compare names and add new ones. Record the ones that were added
    LRow = ActivitySheet.Range("B:B").Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    For Each CopyCell In RosterNames.Offset(0, 1)
        Set PasteCell = NameMatch(CopyCell, ActivityNames)
        If PasteCell Is Nothing Then
            LRow = LRow + 1
            CopyCell.EntireRow.Copy
            ActivitySheet.Cells(LRow, 1).PasteSpecial xlPasteValues
            NewStudents = NewStudents + CopyCell.Value + " " + CopyCell.Offset(0, 1).Value + vbCr
        End If
    Next CopyCell
            
    'Delete and remake the table
    ActivitySheet.ListObjects(1).Unlist
    Call TableCreate(ActivitySheet, ActivityTableStart)
    
    'Change the font of the first column to Marlett
    ActivitySheet.ListObjects(1).ListColumns("Select").DataBodyRange.Font.Name = "Marlett"
    
    'Show added students
    If Not Len(NewStudents) > 0 Then
        MsgBox ("All selected students were already saved to the activity.")
    Else
        MsgBox ("The following students were added: " & vbCr & NewStudents)
    End If
    
    'Reprotect the Activity sheet
    ActivitySheet.Protect , userinterfaceonly:=True, AllowSorting:=True, AllowFiltering:=True, AllowFormattingColumns:=True
    ActivitySheet.Cells.Locked = False
    ActivitySheet.Range("A1:A5").EntireRow.Locked = True
    ActivitySheet.Range("B3:B4").Locked = False 'Allow the date and decription to be editable
    
    AddStudentsForm.Hide
    ActivitySheet.Activate
    
Footer:

    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
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
    If AddStudentsSelectBox.ListCount > 0 Then
        AddStudentsSelectBox.Clear
    End If
    
    'Make columns in the list box
    With AddStudentsSelectBox
        .ColumnCount = 3
        .ColumnWidths = "150, 150, 30"
    End With
    
    'Define the range with labels
    FCol = RecordsSheet.Range("1:1").Find("V BREAK", , xlValues, xlWhole).Column
    LCol = RecordsSheet.Range("1:1").Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Column
    
    'Check if anything has been saved. This should be taken care of previously
    If LCol = FCol Then
        MsgBox ("You have no saved activities.")
        AddStudentsForm.Hide
        GoTo Footer
    End If
    
    Set LabelRange = RecordsSheet.Range(Cells(1, FCol + 1).Address, Cells(1, LCol).Address)
    
    'Copy over the label information
    i = 0
    For Each c In LabelRange
        LabelName = c.Value
        LabelPractice = c.Offset(1, 0).Value
        LabelDate = CDate(c.Offset(2, 0).Value)
        
        With AddStudentsSelectBox
            .AddItem LabelName
            .List(i, 1) = LabelPractice
            .List(i, 2) = LabelDate
        End With
        i = i + 1
    Next c
    
Footer:

End Sub
