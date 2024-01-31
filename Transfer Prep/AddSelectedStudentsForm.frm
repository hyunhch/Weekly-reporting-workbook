VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AddSelectedStudentsForm 
   Caption         =   "Add selected students to an activity"
   ClientHeight    =   4005
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7455
   OleObjectBlob   =   "AddSelectedStudentsForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "AddSelectedStudentsForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SelectedSheet As String

Private Sub CancelButton_Click()

    AddSelectedStudentsForm.Hide

End Sub

Private Sub StudentsAddSelectedButton_Click()
'Add students checked on cover sheet to the selected activity sheet
'Exclude duplicates

    Dim CoverSheet As Worksheet
    Dim PasteSheet As Worksheet
    Dim LRow As Long
    Dim LCol As Long
    Dim DestLRow As Long
    Dim TableEnd As Range
    Dim TableStart As Range
    Dim c As Range
    Dim BoxHere As Range
    Dim CopyRange As Range
    Dim PasteRange As Range
    
    Set CoverSheet = Worksheets("Cover Page")
    Set PasteSheet = Worksheets(SelectedSheet)
    Set TableStart = CoverSheet.Range("A10")
    
    LRow = CoverSheet.Cells(Rows.Count, 1).End(xlUp).Row
    LCol = CoverSheet.Cells(TableStart.Row, Columns.Count).End(xlToLeft).Column
    
    DestLRow = PasteSheet.Cells(Rows.Count, 1).End(xlUp).Row
    
    For Each c In CoverSheet.Range(Cells(TableStart.Row + 1, 1), Cells(LRow, 1))
        If c.Value <> "" Then
            Set CopyRange = CoverSheet.Range(Cells(c.Row, 2).Address, Cells(c.Row, LCol).Address)
            Set PasteRange = PasteSheet.Range(Cells(DestLRow + 1, 2).Address)
            CopyRange.Copy PasteRange
            Set BoxHere = PasteSheet.Range(Cells(DestLRow + 1, 1).Address)
            Call AddMarlettBox(BoxHere, PasteSheet)
            DestLRow = DestLRow + 1
        End If
    Next
    
    AddSelectedStudentsForm.Hide
    PasteSheet.Activate

End Sub

Private Sub SheetNameListBox_AfterUpdate()

    SelectedSheet = Me.SheetNameListBox

End Sub

Private Sub UserForm_Activate()

'Put each activity, with its date, into a listbox

    Dim ActivitySheet As Worksheet
    Dim SheetName As String
    Dim i As Long
    
    'We have to clear the list before we update it
    If SheetNameListBox.ListCount > 0 Then
        SheetNameListBox.Clear
    End If
    
    With SheetNameListBox
        .ColumnCount = 3
        .ColumnWidths = "60;60;90"
    End With
    i = -1
    
    For Each ActivitySheet In ThisWorkbook.Worksheets
        SheetName = ActivitySheet.Name
        If InStr(SheetName, "Activity") > 0 Then
            i = i + 1
            With SheetNameListBox
                .AddItem SheetName
                .List(i, 1) = ActivitySheet.Range("B3").Value
                .List(i, 2) = ActivitySheet.Range("F1").Value
            End With
        End If
    Next

End Sub
