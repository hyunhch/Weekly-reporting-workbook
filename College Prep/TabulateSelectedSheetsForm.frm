VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} TabulateSelectedSheetsForm 
   Caption         =   "UserForm1"
   ClientHeight    =   3990
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7440
   OleObjectBlob   =   "TabulateSelectedSheetsForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "TabulateSelectedSheetsForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SelectedSheet As String

Private Sub CancelButton_Click()

    TabulateSelectedSheetsForm.Hide

End Sub

Private Sub SheetNameListBox_AfterUpdate()

    SelectedSheet = Me.SheetNameListBox

End Sub

Private Sub TabulateAddSelectedButton_Click()

    Dim i As Long
    Dim SheetName As String
    Dim ActivitySheet As Worksheet
    Dim ReportSheet As Worksheet
    
    Set ReportSheet = Worksheets("Report Page")
    With TabulateSelectedSheetsForm
        For i = 0 To .SheetNameListBox.ListCount - 1
            If .SheetNameListBox.Selected(i) = True Then
                SheetName = .SheetNameListBox.List(i, 0)
                Set ActivitySheet = Worksheets(SheetName)
                Call TabulateActivities(ActivitySheet)
            End If
        Next i
    End With
    
    TabulateSelectedSheetsForm.Hide
    ReportSheet.Activate
    
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
