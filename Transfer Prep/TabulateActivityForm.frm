VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} TabulateActivityForm 
   Caption         =   "Tabulate Activities"
   ClientHeight    =   4665
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8640.001
   OleObjectBlob   =   "TabulateActivityForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "TabulateActivityForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub DeleteActivityCancelButton_Click()
'Hide the form

    TabulateActivityForm.Hide
    
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

End Sub

Private Sub TabulateActivityAllConfirmButton_Click()
'Tabualte every saved activity

    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    Call TabulateAll
    
    TabulateActivityForm.Hide
    
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

End Sub

Private Sub TabulateActivityConfirmButton_Click()
'Tabulate any of the selected activites
    
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    'Make sure an activity is selected
    If TabulateActivitySelectBox.ListIndex = -1 Then
        MsgBox ("Please select an activity.")
        GoTo Footer
    End If
    
    'Loop through listbox for all selected items
    Dim RecordsSheet As Worksheet
    Dim ReportSheet As Worksheet
    Dim LabelMatch As Range
    Dim LabelString As String
    Dim NumRows As Long
    Dim i As Long
    
    Set RecordsSheet = Worksheets("Records Page")
    Set ReportSheet = Worksheets("Report Page")
    
    NumRows = Me.TabulateActivitySelectBox.ListCount - 1
    For i = 0 To NumRows
        If Me.TabulateActivitySelectBox.Selected(i) Then
            'Make sure the label is on the Records sheet
            LabelString = Me.TabulateActivitySelectBox.List(i, 0)
            Set LabelMatch = RecordsSheet.Range("1:1").Find(LabelString, , xlValues, xlWhole)
            
            'If there is no match. This shouldn't happen
            If LabelMatch Is Nothing Then
                MsgBox ("Something has gone wrong, the activity labeled " & LabelString & "cannot be found.")
                GoTo NextIteration
            End If
            
            'Pass for tabulation
            Call TabulateActivity(LabelString)
        End If
NextIteration:
    Next i
            
    'Close the form and bring up the Report sheet
    TabulateActivityForm.Hide
    ReportSheet.Activate
    
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
    If TabulateActivitySelectBox.ListCount > 0 Then
        TabulateActivitySelectBox.Clear
    End If
    
    'Make columns in the list box
    With TabulateActivitySelectBox
        .ColumnCount = 3
        .ColumnWidths = "150, 150, 30"
    End With
    
    'Define the range with labels
    FCol = RecordsSheet.Range("1:1").Find("V BREAK", , xlValues, xlWhole).Column
    LCol = RecordsSheet.Range("1:1").Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Column
    
    'Check if anything has been saved. This should be taken care of previously
    If LCol = FCol Then
        MsgBox ("You have no saved activities.")
        TabulateActivityForm.Hide
        GoTo Footer
    End If
    
    Set LabelRange = RecordsSheet.Range(Cells(1, FCol + 1).Address, Cells(1, LCol).Address)
    
    'Copy over the label information
    i = 0
    For Each c In LabelRange
        LabelName = c.Value
        LabelPractice = c.Offset(1, 0).Value
        LabelDate = CDate(c.Offset(2, 0).Value)
        
        With TabulateActivitySelectBox
            .AddItem LabelName
            .List(i, 1) = LabelPractice
            .List(i, 2) = LabelDate
        End With
        i = i + 1
    Next c
    
Footer:

End Sub

