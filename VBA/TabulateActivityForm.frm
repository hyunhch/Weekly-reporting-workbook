VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} TabulateActivityForm 
   Caption         =   "Tabulate Activity"
   ClientHeight    =   5352
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8670.001
   OleObjectBlob   =   "TabulateActivityForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "TabulateActivityForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Private Sub Image1_BeforeDragOver(ByVal Cancel As MSForms.ReturnBoolean, ByVal Data As MSForms.DataObject, ByVal x As Single, ByVal Y As Single, ByVal DragState As MSForms.fmDragState, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)

End Sub

Private Sub TabulateActivityCancelButton_Click()
'Hide the form

    TabulateActivityForm.Hide

    'Application.EnableEvents = True
    'Application.ScreenUpdating = True
    'Application.DisplayAlerts = True

End Sub

Private Sub TabulateActivityConfirmAllButton_Click()
'Tabulate everything displayed, regardless of selection

    Dim RecordsSheet As Worksheet
    Dim LabelCell As Range
    Dim i As Long
    Dim LabelString As String

    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    Set RecordsSheet = Worksheets("Records Page")
    Set LabelCell = RecordsSheet.Range("A1")

    'First tabualate the totals
    Call TabulateReportTotals

    'Pass each selected activity for tabulation
    For i = 0 To Me.TabulateActivityListBox.ListCount - 1
        LabelString = Me.TabulateActivityListBox.List(i, 0)
        LabelCell.Value = LabelString
        
        Call TabulateActivity(LabelCell)
        LabelCell.ClearContents
    Next i
        
    TabulateActivityForm.Hide
   
Footer:

End Sub

Private Sub TabulateActivityConfirmButton_Click()
'Recreate an activity sheet with the activity information and attendance
    
    Dim RecordsSheet As Worksheet
    Dim LabelCell As Range
    Dim LabelString As String
    Dim i As Long
    
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    Set RecordsSheet = Worksheets("Records Page")
    Set LabelCell = RecordsSheet.Range("A1")

    'Make sure an activity has been selected
    If TabulateActivityListBox.ListIndex = -1 Then
        MsgBox ("Please select an activity")
        GoTo Footer
    End If

    For i = 0 To Me.TabulateActivityListBox.ListCount - 1
        If Me.TabulateActivityListBox.Selected(i) Then
            LabelString = Me.TabulateActivityListBox.List(i, 0)
            LabelCell.Value = LabelString
            
            Call TabulateActivity(LabelCell)
            LabelCell.ClearContents
        End If
    Next i
    
    'Tabulate the totals
    Call TabulateReportTotals
    
    TabulateActivityForm.Hide
   
Footer:
    'Application.EnableEvents = True
    'Application.ScreenUpdating = True
    'Application.DisplayAlerts = True

End Sub

Private Sub TabulateActivityFilterTextBox_Change()
'Dynamic filter for the activity list

    Dim i As Long
    Dim testString As String
    
    testString = LCase("*" & TabulateActivityFilterTextBox.Text & "*")
    Call UserForm_Activate
    
    With TabulateActivityListBox
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
    Dim RecordsLabelRange As Range
    Dim ReportLabelRange As Range
    Dim AddRange As Range
    Dim PracticeRange As Range
    Dim DateRange As Range
    Dim DescriptionRange As Range
    Dim LabelHeaderRange As Range
    Dim c As Range
    Dim i As Long
 
    Set RecordsSheet = Worksheets("Records Page")
    Set ReportSheet = Worksheets("Report Page")
    
    'Clear out anything that's already in the list box
    If TabulateActivityListBox.ListCount > 0 Then
        TabulateActivityListBox.Clear
    End If
    
    'Make columns in the list box
    With TabulateActivityListBox
        .ColumnCount = 3
        .ColumnWidths = "150, 150, 30, 0"
    End With
    
    'Checking that there are activities happens in parent sub
    Set RecordsLabelRange = FindRecordsLabel(RecordsSheet)
    
    'If all activities are removed from the view, it will pull in the padding cell
    If RecordsLabelRange.Cells.Count = 1 Then
        If RecordsLabelRange.Value = "V BREAK" Then
            GoTo Footer
        End If
    End If
    
    'Grab all activities on the RecordsSheet
    Set AddRange = RecordsLabelRange
    
    'Find all labels on the Report Page to see which are already tabulated
    'Not doing this for now, we want all activities
    'If ReportSheet.ListObjects(1).ListRows.Count < 2 Then
        'Set AddRange = RecordsLabelRange
        'GoTo EmptyReport
    'End If
    
    'Create a range of all activities
    'Set ReportLabelRange = FindReportLabel(ReportSheet)
    
    'Find all activites on the Records page and NOT on the Report page
    'Not doing this for now, we want all activities
    'For Each c In RecordsLabelRange
        'If ReportLabelRange.Find(c.Value, , xlValues, xlWhole) Is Nothing Then
            'If AddRange Is Nothing Then
                'Set AddRange = c
            'Else
                'Set AddRange = Union(AddRange, c)
            'End If
        'End If
    'Next c
    
    If AddRange Is Nothing Then
        GoTo Footer
    End If
    
EmptyReport:
    'Find where the values we need are
    Set LabelHeaderRange = FindRecordsActivityHeaders(RecordsSheet)
    Set PracticeRange = LabelHeaderRange.Find("Practice", , xlValues, xlWhole)
    Set DateRange = LabelHeaderRange.Find("Date", , xlValues, xlWhole)
    Set DescriptionRange = LabelHeaderRange.Find("Description", , xlValues, xlWhole)
    
    'Copy over the label information
    i = 0
    For Each c In AddRange
        With TabulateActivityListBox
            .AddItem c.Value
            .List(i, 1) = RecordsSheet.Cells(PracticeRange.Row, c.Column)
            .List(i, 2) = CDate(RecordsSheet.Cells(DateRange.Row, c.Column))
            .List(i, 3) = RecordsSheet.Cells(DescriptionRange.Row, c.Column)
        End With
        
        i = i + 1
NextRow:
    Next c
    
Footer:

End Sub

Private Sub UserForm_Deactivate()
'Bring up the Report Page and enable events

    Dim ReportSheet As Worksheet
    
    Set ReportSheet = Worksheets("Report Page")
    
    ReportSheet.Activate
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

End Sub
