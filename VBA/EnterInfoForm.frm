VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} EnterInfoForm 
   Caption         =   "Submission Information"
   ClientHeight    =   3240
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4590
   OleObjectBlob   =   "EnterInfoForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "EnterInfoForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub EnterInfoNoButton_Click()

    EnterInfoForm.Hide
    ActiveWorkbook.Close SaveChanges:=False
    
End Sub

Private Sub EnterInfoYesButton_Click()
'Verifies information entered, enter into the Cover Page

    Dim CoverSheet As Worksheet
    Dim PasteRange As Range
    Dim i As Long
    Dim j As Long
    Dim NameCheck As String
    Dim DateCheck As String
    Dim CenterCheck As String
    Dim CopyArray() As Variant
    
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    'Verify information on form
    NameCheck = Me.EnterInfoNameBox.Value
    DateCheck = Me.EnterInfoDateBox.Value
    CenterCheck = Me.EnterInfoCenterComboBox.Value
    
    If Len(Trim(NameCheck)) = 0 Then
        MsgBox ("Please enter your name")
        GoTo Footer
    ElseIf Not IsDate(DateCheck) Then
        MsgBox ("Please enter a date in the form mm/dd/yyyy")
        GoTo Footer
    ElseIf Len(Trim(CenterCheck)) = 0 Then
        MsgBox ("Please select your center")
        GoTo Footer
    End If
    
    'Put into an array
    ReDim CopyArray(0 To 2, 0 To 1)
    
    CopyArray(0, 0) = "Name"
    CopyArray(1, 0) = "Date"
    CopyArray(2, 0) = "Center"
    CopyArray(0, 1) = NameCheck
    CopyArray(1, 1) = DateCheck
    CopyArray(2, 1) = CenterCheck
    
    'Find where to paste on the Cover Page. It will be one cell to the right of the searched word
    Set CoverSheet = Worksheets("Cover Page")
    
    Call UnprotectSheet(CoverSheet)
    
    For i = LBound(CopyArray) To UBound(CopyArray)
        Set PasteRange = CoverSheet.Range("A:A").Find(CopyArray(i, 0), , xlValues, xlWhole).Offset(0, 1)
        
        If Not PasteRange Is Nothing Then
            PasteRange.Value = CopyArray(i, 1)
        End If
    Next i

    EnterInfoForm.Hide
    CoverSheet.Activate

Footer:

    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

End Sub

Private Sub UserForm_Activate()

    EnterInfoForm.Height = 191
    EnterInfoForm.Width = 242

End Sub


Private Sub UserForm_Initialize()
'To populate the Cover Page, triggered after the program is selected

    Dim c As Range
    
    For Each c In Range("CentersList")
        Me.EnterInfoCenterComboBox.AddItem c.Value
    Next c

End Sub
