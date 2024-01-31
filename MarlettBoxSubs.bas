Attribute VB_Name = "MarlettBoxSubs"
Sub AddMarlettBox(BoxHere As Range, TargetSheet As Worksheet)
'Doing this instead of actual checkboxes to deal with sorting issues
'This only changes the font of a range to Marlett

    With BoxHere
        .Font.Name = "Marlett"
        .Value = ""
    End With

End Sub

Sub AddSelectAll(BoxHere As Range, TargetSheet As Worksheet)
'Insert a button

    Dim NewButton As Button
    
    Set NewButton = TargetSheet.Buttons.Add(BoxHere.Left, BoxHere.Top, _
        BoxHere.Width, BoxHere.Height)
    
    With NewButton
        .OnAction = "SelectAll"
        .Caption = "Select All"
    End With

End Sub

Sub AddDeselectAll(BoxHere As Range, TargetSheet As Worksheet)
'Insert a button

    Dim NewButton As Button
    
    Set NewButton = TargetSheet.Buttons.Add(BoxHere.Left, BoxHere.Top, _
        BoxHere.Width, BoxHere.Height)
    
    With NewButton
        .OnAction = "DeselectAll"
        .Caption = "Deselect All"
    End With

End Sub

Sub SelectAll()
'For current sheet. Looks for font

    Dim FRow As Long
    Dim LRow As Long
    Dim CheckRange As Range
    Dim i As Long
    
    FRow = ActiveSheet.Range("A:A").Find("Select", LookIn:=xlValues).Row
    
    'In case the column name was changed or there is some other problem
    If Not FRow > 0 Then
        MsgBox ("There is a problem with the table." & vbCr & _
            "Please make sure the first column is named ""Select""")
        Exit Sub
    End If
    
    LRow = ActiveSheet.Cells(Rows.Count, 2).End(xlUp).Row
    
    'Check that there is at least row of students
    'The report sheet needs to be offset an additional row
    If ActiveSheet.Name = "Report Page" Then
        GoTo ReportProcedure
    End If
    
    If Not LRow > FRow Then
        MsgBox ("Please add at least one row to the table.")
        Exit Sub
    End If
    
    Set CheckRange = ActiveSheet.Range(Cells(FRow + 1, 1).Address, Cells(LRow, 1).Address)
    CheckRange.Font.Name = "Marlett"
    
    'Check all if any are blank, uncheck all if none are blank
    If Application.CountIf(CheckRange, "a") = LRow - (FRow) Then
        CheckRange.Value = ""
    Else
        CheckRange.Value = "a"
    End If
    
    Exit Sub
    
    'Report Page
ReportProcedure:
    If Not LRow > FRow + 1 Then
        MsgBox ("Please add at least one row to the table.")
        Exit Sub
    End If

    Set CheckRange = ActiveSheet.Range(Cells(FRow + 2, 1).Address, Cells(LRow, 1).Address)
    CheckRange.Font.Name = "Marlett"

    If Application.CountIf(CheckRange, "a") = LRow - (FRow + 1) Then
        CheckRange.Value = ""
    Else
        CheckRange.Value = "a"
    End If
    
End Sub

Sub DeselectAll()
'For current sheet. Looks for font
'Currently not using this, using "Select All" button to do both

    Dim FRow As Long
    Dim LRow As Long
    Dim i As Long
    
    FRow = ActiveSheet.Range("A:A").Find("Select", LookIn:=xlValues).Row
    
    'In case the column name was changed or there is some other problem
    If Not FRow > 0 Then
        MsgBox ("There is a problem with the table." & vbCr & _
            "Please make sure the first column is named ""Select""")
        Exit Sub
    End If
    
    LRow = ActiveSheet.Cells(Rows.Count, 1).End(xlUp).Row
    
    'Check that there is at least row of students
    If Not LRow > FRow Then
        MsgBox ("Please add at least one student to the table.")
        Exit Sub
    End If
    
    For i = FRow + 1 To LRow
        ActiveSheet.Cells(i, 1).Font.Name = "Marlett"
        ActiveSheet.Cells(i, 1).Value = ""
    Next i

End Sub

Function AnyChecked(StartRow As Long, StopRow As Long, TargetSheet As Worksheet) As Boolean
'Check to see if any students have been checked

    Dim CheckRange As Range
    Dim CheckCell As Range
    
    AnyChecked = False
    Set CheckRange = TargetSheet.Range(Cells(StartRow, 1).Address, Cells(StopRow, 1).Address)
    
    For Each CheckCell In CheckRange
        If CheckCell.Value = "a" Then
            AnyChecked = True
            Exit Function
        End If
    Next CheckCell

End Function

