Attribute VB_Name = "TableSubs"
Option Explicit

Sub TableFormat(TargetSheet As Worksheet, TargetTable As ListObject)
'Flag blanks and bad entries on Roster and Activity Sheets
'Different formatting on Report Page

    Dim EthnicityRange As Range
    Dim GenderRange As Range
    Dim GradeRange As Range
    Dim CreditsRange As Range
    Dim MajorRange As Range
    Dim c As Range
    Dim FormulaString1 As String
    Dim FormulaString2 As String
    
    'If there are no rows, skip
    If TargetTable.DataBodyRange Is Nothing Then
        GoTo Footer
    End If
    
    'Blank cells flagged yellow, except in the first column
    With TargetTable
        .DataBodyRange.FormatConditions.Delete
        .DataBodyRange.FormatConditions.Add Type:=xlBlanksCondition
        .ListColumns("Select").DataBodyRange.FormatConditions.Delete
        
        With .DataBodyRange.FormatConditions(1)
        .StopIfTrue = False
        .Interior.ColorIndex = 36
        End With
    End With
    
    'Validate demographics using tables on Ref Tables sheet
    Set EthnicityRange = TargetTable.ListColumns("Race").DataBodyRange
    Set GenderRange = TargetTable.ListColumns("Gender").DataBodyRange

    For Each c In EthnicityRange
        FormulaString1 = "=AND(COUNTIFS("
        FormulaString2 = "," & "Trim(" & c.Address & ")) < 1, NOT(ISBLANK(" & c.Address & ")))"
        c.FormatConditions.Add Type:=xlExpression, Formula1:=FormulaString1 & "EthnicityList" & FormulaString2
        With c.FormatConditions(2)
            .StopIfTrue = False
            .Interior.Color = vbRed
        End With
    Next c
    
    For Each c In GenderRange
        FormulaString1 = "=AND(COUNTIFS("
        FormulaString2 = "," & "Trim(" & c.Address & ")) < 1, NOT(ISBLANK(" & c.Address & ")))"
        c.FormatConditions.Add Type:=xlExpression, Formula1:=FormulaString1 & "GenderList" & FormulaString2
        With c.FormatConditions(2)
            .StopIfTrue = False
            .Interior.Color = vbRed
        End With
    Next c
    
    'Check if this is for College Prep
    If IsCollege = True Then
        GoTo CollegePrep
    End If

    'Credit hours don't need the reference table, add majors to validation
    Set CreditsRange = TargetTable.ListColumns("Credits").DataBodyRange
    Set MajorRange = TargetTable.ListColumns("Major").DataBodyRange
    
    For Each c In CreditsRange
        FormulaString1 = "=AND(NOT(ISNUMBER(" + c.Address + ")),"
        FormulaString2 = " NOT(ISBLANK(" + c.Address + ")))"
        c.FormatConditions.Add Type:=xlExpression, Formula1:=FormulaString1 & FormulaString2
        With c.FormatConditions(2)
            .StopIfTrue = False
            .Interior.Color = vbRed
        End With
    Next c
    
    'Flag majors orange instead of red
    For Each c In MajorRange
        FormulaString1 = "=AND(COUNTIFS("
        FormulaString2 = "," & "Trim(" & c.Address & ")) < 1, NOT(ISBLANK(" & c.Address & ")))"
        c.FormatConditions.Add Type:=xlExpression, Formula1:=FormulaString1 & "MajorList" & FormulaString2
        With c.FormatConditions(2)
            .StopIfTrue = False
            .Interior.ColorIndex = 45
        End With
    Next c

    GoTo Footer

CollegePrep:
    'Grades work as both strings and numbers
    Set GradeRange = TargetTable.ListColumns("Grade").DataBodyRange
    
    For Each c In GradeRange
        FormulaString1 = "=AND(COUNTIFS("
        FormulaString2 = "," & "Trim(" & c.Address & ")) < 1, NOT(ISBLANK(" & c.Address & ")))"
        c.FormatConditions.Add Type:=xlExpression, Formula1:=FormulaString1 & "GradeList" & FormulaString2
        With c.FormatConditions(2)
            .StopIfTrue = False
            .Interior.Color = vbRed
        End With
    Next c

Footer:

End Sub

Sub TableFormatReport(ReportSheet As Worksheet, ReportTable As ListObject)
'Adjust font, color, number formats for the table on the Report Page

    Dim c As Range
    Dim i As Long
    Dim RedValue As String
    Dim GreenValue As String
    Dim BlueValue As String
    Dim TempArray() As String
    Dim ColorArray() As Variant

    'I'm not sure how to prevent the totals row from sorting, so for now I'll disable the autofilter
    ReportTable.ShowAutoFilterDropDown = False
    
    'Remove all colors, then add them in the correct columns. Add vertical lines
    ColorArray = Application.Transpose(ActiveWorkbook.Names("ReportColumnRGBList").RefersToRange.Value)
    
    For i = 1 To ReportTable.ListColumns.Count
        TempArray = Split(ColorArray(i), ",")
        RedValue = TempArray(0)
        GreenValue = TempArray(1)
        BlueValue = TempArray(2)
        With ReportTable.ListColumns(i)
            .Range.Interior.Color = RGB(RedValue, GreenValue, BlueValue)
            .Range.Borders(xlEdgeLeft).LineStyle = xlContinuous
            .Range.Borders(xlEdgeRight).LineStyle = xlContinuous
            .Range.Columns.HorizontalAlignment = xlCenter
        End With
    Next i

    'Grey out the cell under "Select"
    TempArray = Split(ColorArray(2), ",")
        RedValue = TempArray(0)
        GreenValue = TempArray(1)
        BlueValue = TempArray(2)
    ReportTable.ListColumns("Select").Range.Resize(1, 1).Offset(1, 0).Interior.Color = RGB(RedValue, GreenValue, BlueValue)

    'Header and Total row, Total column text black, bold
    With ReportTable.HeaderRowRange.Resize(2, ReportTable.ListColumns.Count)
        .Font.Color = vbBlack
        .Font.Bold = True
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
    End With
    
    With ReportTable.ListColumns("Total").DataBodyRange
        .Font.Color = vbBlack
        .Font.Bold = True
    End With
    
    'Label, Pratice, description left aligned, Label bold
    With ReportSheet.Range(ReportTable.ListColumns("Label").DataBodyRange, ReportTable.ListColumns("Description").DataBodyRange)
        .HorizontalAlignment = xlLeft
    End With
    
    With ReportTable.ListColumns("Label").DataBodyRange
        .Font.Bold = True
    End With
    
    '2nd row of both back to center aligned
    ReportTable.ListRows(1).Range.HorizontalAlignment = xlCenter
    
    'Format date
    Set c = FindTableHeader(ReportSheet, "Date")
        If Not c Is Nothing Then
            c.Offset(1, 0).NumberFormat = "mm/dd/yyyy"
        End If
            
    'Autofit Category and Practice columns
    ReportTable.ListColumns("Category").Range.EntireColumn.AutoFit
    ReportTable.ListColumns("Practice").Range.EntireColumn.AutoFit
            
Footer:

End Sub

Sub TableResetHeaders(TargetSheet As Worksheet, TargetCell As Range, TargetFields As Variant)
'Insert the default column names starting at the passed cell

    Dim FieldRange As Range
    
    Set FieldRange = TargetSheet.Range(TargetCell, TargetCell.Offset(0, UBound(TargetFields) - 1))
    FieldRange.Value = TargetFields

End Sub

Sub TableResetReportTotalHeaders(ReportSheet As Worksheet, TargetRange As Range)
'Grabs information from the cover sheet and a reference table for the 2nd row of the Report Table

    Dim c As Range
    Dim TempRange As Range
    Dim i As Long
    Dim TempArray() As Variant
    Dim TotalsArray() As Variant

    'Grab the headers for the total row from a named range and the cover
    TempArray = GetCoverInfo()
    Set TempRange = Range("ReportTotalsRowList")
    i = UBound(TempArray, 2) + TempRange.Cells.Count
    
    ReDim TotalsArray(1 To 2, 1 To i)
    
    For i = 1 To UBound(TempArray, 2)
        TotalsArray(1, i) = TempArray(1, i)
        TotalsArray(2, i) = TempArray(2, i)
    Next i
    
    For Each c In TempRange
        TotalsArray(1, i) = c.Offset(0, -1).Value
        TotalsArray(2, i) = c.Value
    
    i = i + 1
    Next c
    
    'Doing this programmatically
    For i = 1 To UBound(TotalsArray, 2)
        Set c = TargetRange.EntireRow.Find(TotalsArray(1, i), , xlValues, xlWhole)
        
        If c Is Nothing Then
            GoTo NextHeader
        End If
    
        c.Offset(1, 0).Value = TotalsArray(2, i)
NextHeader:
    Next i

End Sub











