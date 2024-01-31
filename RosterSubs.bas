Attribute VB_Name = "RosterSubs"
Option Explicit

Sub ClearRoster(PasteRange As Range, Repull As Long, TargetSheet As Worksheet)
'Repull = 1 avoids warning message

    Dim DelRange As Range
    Dim ClearAll As Long
    Dim OldTable As ListObject

    Set DelRange = TargetSheet.Range(Cells(PasteRange.Row, PasteRange.Column), Cells(TargetSheet.Rows.Count, TargetSheet.Columns.Count))
    
    If Repull <> 1 Then
            ClearAll = MsgBox("Are you sure you want to clear all content?" & vbCrLf & "This cannot be undone.", vbQuestion + vbYesNo + vbDefaultButton2, "")
    Else
        ClearAll = vbYes
    End If
    
    If ClearAll = vbYes Then
        For Each OldTable In TargetSheet.ListObjects
            OldTable.Unlist
        Next OldTable
        
        With DelRange
            .ClearContents
            .ClearFormats
            .Validation.Delete
        End With
    End If
    
End Sub

Sub CopyRoster(PasteRange As Range)

    Dim CoverSheet As Worksheet
    Dim RosterSheet As Worksheet
    Dim LRow As Long
    Dim LCol As Long
    Dim TableRange As Range
    Dim RosterTableStart As Range
    
    Set CoverSheet = Worksheets("Cover Page")
    Set RosterSheet = Worksheets("Roster Page")
    Set RosterTableStart = RosterSheet.Range("A6")
    
    LRow = RosterSheet.Cells(Rows.Count, 2).End(xlUp).Row
    LCol = RosterSheet.Cells(RosterTableStart.Row, Columns.Count).End(xlToLeft).Column
    
    If LRow = RosterTableStart.Row Then
        MsgBox ("Your roster is empty." & vbCr & _
        "Please paste in your student list")
        Exit Sub
    End If
    
    'Should replace this with checking against an array of column names
    'If LCol < 5 Then
    '    MsgBox ("Your roster is missing headers." & vbCr & _
    '    "Please use a fresh copy of this file")
    'End If
    
    Set TableRange = RosterSheet.Range(Cells(RosterTableStart.Row, 1).Address, Cells(LRow, LCol).Address)
    TableRange.Copy
    PasteRange.PasteSpecial xlPasteValues
    
End Sub

Sub TableFormat(NewTable As ListObject, TargetSheet As Worksheet)
  
    'Blanks
    With NewTable.DataBodyRange
        .FormatConditions.Delete
        .FormatConditions.Add Type:=xlBlanksCondition
    End With
    'Clear from the first column
    NewTable.ListColumns(1).DataBodyRange.FormatConditions.Delete
    
    With NewTable.DataBodyRange.FormatConditions(1)
        .StopIfTrue = False
        .Interior.Color = 49407
    End With
    
    'Demographics
    Dim RaceSource As String
    Dim GenderSource As String
    Dim GradeSource As String
    Dim RefSheet As Worksheet
    Dim RaceRange As Range
    Dim GenderRange As Range
    Dim GradeRange As Range
    Dim c As Range
    Dim FormulaString1 As String
    Dim FormulaString2 As String
    
    Set RefSheet = Worksheets("Ref Tables")
    RaceSource = RefSheet.ListObjects("EthnicityTable").DataBodyRange.Columns(1).Address
    GenderSource = RefSheet.ListObjects("GenderTable").DataBodyRange.Columns(1).Address
    GradeSource = RefSheet.ListObjects("GradeTable").DataBodyRange.Columns(1).Address

    Set RaceRange = NewTable.ListColumns("Ethnicity").DataBodyRange
    Set GenderRange = NewTable.ListColumns("Gender").DataBodyRange
    Set GradeRange = NewTable.ListColumns("Grade").DataBodyRange
    
    For Each c In RaceRange
        FormulaString1 = "=AND(ISERROR(MATCH(TRIM(" + c.Address + ")," + "'" + "Ref Tables" + "'" + "!"
        FormulaString2 = ",0)), NOT(ISBLANK(" + c.Address + ")))"
        c.FormatConditions.Add Type:=xlExpression, Formula1:=FormulaString1 & RaceSource & FormulaString2
        With c.FormatConditions(2)
            .StopIfTrue = False
            .Interior.Color = vbRed
        End With
    Next c
    
    For Each c In GenderRange
        FormulaString1 = "=AND(ISERROR(MATCH(TRIM(" + c.Address + ")," + "'" + "Ref Tables" + "'" + "!"
        FormulaString2 = ",0)), NOT(ISBLANK(" + c.Address + ")))"
        c.FormatConditions.Add Type:=xlExpression, Formula1:=FormulaString1 & GenderSource & FormulaString2
        With c.FormatConditions(2)
            .StopIfTrue = False
            .Interior.Color = vbRed
        End With
    Next c

    'Add VALUE() since we're looking at numbers
    For Each c In GradeRange
        FormulaString1 = "=AND(ISERROR(MATCH(VALUE(TRIM(" + c.Address + "))," + "'" + "Ref Tables" + "'" + "!"
        FormulaString2 = ",0)), NOT(ISBLANK(" + c.Address + ")))"
        c.FormatConditions.Add Type:=xlExpression, Formula1:=FormulaString1 & GradeSource & FormulaString2
        With c.FormatConditions(2)
            .StopIfTrue = False
            .Interior.Color = vbRed
        End With
    Next c
    
End Sub



