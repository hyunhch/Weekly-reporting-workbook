Attribute VB_Name = "MarlettBoxSubs"
Option Explicit

Sub AddMarlettBox(BoxHere As Range, TargetSheet As Worksheet)
'Doing this instead of actual checkboxes to deal with sorting issues
'This only changes the font of a range to Marlett
    
    Dim c As Range

    With BoxHere
        .Font.Name = "Marlett"
        .HorizontalAlignment = xlRight
    End With
    
    'Preserve checks, but get rid of anything other than an "a"
    For Each c In BoxHere
        If c.Value <> "a" Then
            c.Value = ""
        End If
    Next c

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

Function FindChecks(TargetRange As Range) As Range
'Returns a range that contains all cells that are not empty

    Dim CheckedRange As Range
    Dim c As Range
    
    For Each c In TargetRange
        If c.Value <> "" Then
            If Not CheckedRange Is Nothing Then
                Set CheckedRange = Union(CheckedRange, c)
            Else
                Set CheckedRange = c
            End If
        End If
    Next c
    
    Set FindChecks = CheckedRange

End Function

Function CountChecks(TargetRange As Range) As Long
'Returns a range that contains all cells with a checkmark

    Dim CheckedRange As Range
    Dim c As Range
    
    CountChecks = 0
    
    For Each c In TargetRange
        If c.Value <> "" Then
            CountChecks = CountChecks + 1
        End If
    Next c
    
End Function

Sub CopySelected(ActivitySheet As Worksheet, Optional HowMany As String)
'Copying all students who had been checked on the Roster sheet or marked on the Records sheet
    
    Dim RosterSheet As Worksheet
    Dim RosterTableStart As Range
    Dim SearchRange As Range
    Dim CopyRange As Range
    Dim PasteRange As Range
    Dim i As Long
    
    Set RosterSheet = Worksheets("Roster Page")
    Set RosterTableStart = RosterSheet.Range("A6")
    Set SearchRange = RosterSheet.ListObjects("RosterTable").ListColumns("Select").Range
    Set PasteRange = ActivitySheet.Range("A6")
    
    'Make sure the roster isn't empty
    If Not CheckTableLength(RosterSheet, RosterTableStart) > 0 Then
        MsgBox ("You don't have any students on this page.")
        GoTo Footer
    End If
    
    'If all students or checked students are copied
    If Not HowMany = "All" Then
        GoTo OnlyChecked
    End If
    
    'All students
    For i = 0 To RosterSheet.ListObjects("RosterTable").Range.Rows.Count
        RosterSheet.ListObjects("RosterTable").Range.Rows(i).EntireRow.Copy
        PasteRange.Offset(i - 1, 0).PasteSpecial xlPasteValues
    Next i
    GoTo Footer
    
OnlyChecked:
    'Make sure at least one student is selected
    If FindChecks(SearchRange) Is Nothing Then
        MsgBox ("Please select at least one student")
        GoTo Footer
    Else
        Set CopyRange = FindChecks(SearchRange)
    End If

    'Copy each checked row
    'The header and rows are copied as well
    Dim c As Range

    i = 0
    For Each c In CopyRange
        c.EntireRow.Copy
        PasteRange.Offset(i, 0).PasteSpecial xlPasteValues
        i = i + 1
    Next c
        
Footer:

End Sub
