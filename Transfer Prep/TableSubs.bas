Attribute VB_Name = "TableSubs"
Option Explicit

Sub ResetColumns(TargetSheet As Worksheet, TargetCell As Range, TargetNames As Variant)
'The default column names
    Dim HeaderRange As Range
    
    Set HeaderRange = TargetSheet.Range(Cells(TargetCell.Row, TargetCell.Column).Address, Cells(TargetCell.Row, UBound(TargetNames)).Address)
    HeaderRange.Value = TargetNames
    
End Sub

Function CheckTableLength(CheckSheet As Worksheet, CheckStart As Range) As Long
'Small sub to make sure there's at least one student

    Dim i As Long
    
    i = CheckSheet.Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    CheckTableLength = i - CheckStart.Row

End Function

Sub TableCreate(NewSheet As Worksheet, NewStartRange As Range, Optional TableName As String)
'Make a table from the top-left corner

    Dim NewTableRange As Range
    Dim LRow As Long
    Dim LCol As Long
    Dim NewTable As ListObject
    
    'The first column will be empty
    LRow = NewSheet.Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    LCol = NewSheet.Cells.Find("*", SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
    
    'Create the table object
    Set NewTableRange = NewSheet.Range(Cells(NewStartRange.Row, NewStartRange.Column).Address, Cells(LRow, LCol).Address)
    NewTableRange.ClearFormats
    Set NewTable = NewSheet.ListObjects.Add(SourceType:=xlSrcRange, Source:=NewTableRange, _
        xlListObjectHasHeaders:=xlYes)
        
    'If specified, apply table name
    If Not Len(TableName) > 0 Then
        GoTo SkipName
    End If
    
    NewSheet.ListObjects(1).Name = TableName
    
SkipName:
    NewSheet.ListObjects(1).ShowTableStyleRowStripes = False
    
    'Pass the table for formatting
    Call TableFormat(NewTable, NewSheet)

End Sub

Sub TableFormat(NewTable As ListObject, TargetSheet As Worksheet)

    'Flag blank cells
    With NewTable.DataBodyRange
        .Cells.FormatConditions.Delete
        .Cells.FormatConditions.Add Type:=xlBlanksCondition
    End With
    
    'Clear from the first column
    NewTable.ListColumns(1).DataBodyRange.FormatConditions.Delete
    
    With NewTable.DataBodyRange.FormatConditions(1)
        .StopIfTrue = False
        .Interior.ColorIndex = 36
    End With
    
    'Validate demographics using tables on Ref Tables sheet
    Dim RaceSource As String
    Dim GenderSource As String
    Dim GradeSource As String
    Dim MajorSource As String
    Dim RefSheet As Worksheet
    Dim RaceRange As Range
    Dim GenderRange As Range
    Dim GradeRange As Range
    Dim MajorRange As Range
    Dim c As Range
    Dim FormulaString1 As String
    Dim FormulaString2 As String
    
    Set RefSheet = Worksheets("Ref Tables")
    RaceSource = RefSheet.ListObjects("EthnicityTable").DataBodyRange.Columns(1).Address
    GenderSource = RefSheet.ListObjects("GenderTable").DataBodyRange.Columns(1).Address
    GradeSource = RefSheet.ListObjects("GradeTable").DataBodyRange.Columns(1).Address
    MajorSource = RefSheet.ListObjects("MajorTable").DataBodyRange.Columns(1).Address

    Set RaceRange = NewTable.ListColumns("Ethnicity").DataBodyRange
    Set GenderRange = NewTable.ListColumns("Gender").DataBodyRange
    Set GradeRange = NewTable.ListColumns("Credits").DataBodyRange
    Set MajorRange = NewTable.ListColumns("Major").DataBodyRange
    
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

    'Add ISNUMBER() since we're looking at numbers
    For Each c In GradeRange
        FormulaString1 = "=AND(NOT(ISNUMBER(" + c.Address + ")),"
        FormulaString2 = " NOT(ISBLANK(" + c.Address + ")))"
        c.FormatConditions.Add Type:=xlExpression, Formula1:=FormulaString1 & FormulaString2
        With c.FormatConditions(2)
            .StopIfTrue = False
            .Interior.Color = vbRed
        End With
    Next c
    
    'Flag these orange instead of red
    For Each c In MajorRange
        FormulaString1 = "=AND(ISERROR(MATCH(TRIM(" + c.Address + ")," + "'" + "Ref Tables" + "'" + "!"
        FormulaString2 = ",0)), NOT(ISBLANK(" + c.Address + ")))"
        c.FormatConditions.Add Type:=xlExpression, Formula1:=FormulaString1 & MajorSource & FormulaString2
        With c.FormatConditions(2)
            .StopIfTrue = False
            .Interior.ColorIndex = 45
        End With
    Next c
    
End Sub

Function FindTableRange(TableSheet As Worksheet, TableStart As Range) As Range
'Finds the range of a table on the sheet
    
    Dim FoundRange As Range
    Dim LRow As Long
    Dim LCol As Long
    
    'All tables have "Select" in their first column, except for the Records page, which has "Label"
    LRow = TableStart.Offset(0, 1).EntireColumn.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    LCol = TableStart.EntireRow.Find("*", SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
    
    If TableSheet.Name = "Records Page" Then
        GoTo RecordsTable
    End If
    
    'Make sure there is more than the heading
    If LRow = TableStart.Row Then
        Set FindTableRange = Nothing
    End If

    'Look to see if there is anything past two spacers, "H BREAK" and "L BREAK". If there is nothing past both, return nothing
RecordsTable:
    If TableSheet.Cells(LRow, TableStart.Column).Value = "H BREAK" And TableSheet.Cells(TableStart.Row, LCol).Value = "V BREAK" Then
        Set FindTableRange = Nothing
    End If
    
    'Define the table's range and return
    Set FoundRange = TableSheet.Range(TableStart, Cells(LRow, LCol).Address)
    Set FindTableRange = FoundRange
    
End Function
