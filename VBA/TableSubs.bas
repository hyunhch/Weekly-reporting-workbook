Attribute VB_Name = "TableSubs"
Option Explicit

Sub AddMarlettBox(BoxHere As Range)
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

Function CreateActivityTable(ActivitySheet As Worksheet) As ListObject
'Puts in headers, creates a table

    Dim RosterSheet As Worksheet
    Dim RecordsSheet As Worksheet
    Dim ActivityTableStart As Range
    Dim c As Range
    Dim LRow As Long
    Dim StartOffset As Long
    Dim i As Long
    Dim HeaderArray() As Variant
    Dim RosterTable As ListObject
    Dim ActivityTable As ListObject
    
    Set RosterSheet = Worksheets("Roster Page")
    Set RecordsSheet = Worksheets("Records Page")
    Set ActivityTableStart = ActivitySheet.Range("A6")
    Set RosterTable = RosterSheet.ListObjects(1)
    
    'If there's already a table, unlist
    Call RemoveTable(ActivitySheet)
    
    'Read in the headers of the Roster Table and pass to copy headers, then make a table
    ReDim HeaderArray(1 To RosterTable.ListColumns.Count)
    i = 1
    
    For Each c In RosterTable.HeaderRowRange
        HeaderArray(i) = c.Value
        i = i + 1
    Next c
    
    Call ResetTableHeaders(ActivitySheet, ActivityTableStart, HeaderArray)
    
    'Create the table, format, add Marlett Boxes
    Set ActivityTable = CreateTable(ActivitySheet)
    
    If ActivityTable.ListRows.Count > 0 Then
        Set c = ActivityTable.ListColumns("Select").DataBodyRange
        Call AddMarlettBox(c)
    End If
    
    Set CreateActivityTable = ActivityTable
    
Footer:

End Function

Function CreateTable(TargetSheet As Worksheet, Optional NewTableName As String, Optional PartialTableRange As Range) As ListObject
'Create a table on an existing or new sheet

    Dim NewTableRange As Range
    Dim NameRange As Range
    Dim BoxRange As Range
    Dim DelRange As Range
    Dim c As Range
    Dim NewTable As ListObject
    
    'Unprotect
    Call UnprotectSheet(TargetSheet)
    
    'Find the range to use
    If PartialTableRange Is Nothing Then
        Set NewTableRange = FindTableRange(TargetSheet)
    Else
        Set NewTableRange = PartialTableRange
    End If

    If NewTableRange Is Nothing Then
        MsgBox ("There was a problem creating a table on sheet " & TargetSheet.Name)
        GoTo Footer
    End If
    
    'Unlist any existing table
    Call RemoveTable(TargetSheet)

    'Clear formats and create a table
    NewTableRange.ClearFormats
    
    Set NewTable = TargetSheet.ListObjects.Add(SourceType:=xlSrcRange, Source:=NewTableRange, _
        xlListObjectHasHeaders:=xlYes)
    
    NewTable.ShowTableStyleRowStripes = False
    
    'Assign a name if passed
    If Len(NewTableName) > 0 Then
        NewTable.Name = NewTableName
    End If
    
    'Look for blanks and remove if needed
    If NewTable.ListRows.Count < 1 Then
        GoTo RetunTable
    End If

    Set NameRange = NewTable.ListColumns("First").DataBodyRange
    Set DelRange = FindBlanks(TargetSheet, NameRange)
    
    If Not c Is Nothing Then
        Call RemoveRows(TargetSheet, NewTable.DataBodyRange, NameRange, DelRange)
        
        If NewTable.ListRows.Count < 1 Then 'In case this gets rid of all rows
            GoTo RetunTable
        End If
    End If
    
    'Put in Marlett boxes. I had taken this out but can't remember why
    Set BoxRange = NewTable.ListColumns("Select").DataBodyRange
    
    Call AddMarlettBox(BoxRange)
    
RetunTable:
    Set CreateTable = NewTable

Footer:

End Function

Function CreateReportTable() As ListObject
'Grabs headers from reference page, unmakes and remakes the table
'Called when adding or deleting rows, tabulating totals

    Dim ReportSheet As Worksheet
    Dim CoverSheet As Worksheet
    Dim ReportTableStart As Range
    Dim ReportLabelRange As Range
    Dim ReportTableRange As Range
    Dim HeaderRange As Range
    Dim BoxRange As Range
    Dim DelRange As Range
    Dim i As Long
    Dim HeaderArray() As Variant
    Dim TotalsArray() As Variant
    Dim CenterInfoArray() As Variant
    Dim ReportTable As ListObject
    
    Set CoverSheet = Worksheets("Cover Page")
    Set ReportSheet = Worksheets("Report Page")
    Set ReportTableStart = ReportSheet.Range("A:A").Find("Select", , xlValues, xlWhole)
       
    Call UnprotectSheet(ReportSheet)
    
    'Remove any existing filters, unlist the table and remove formatting
    If ReportSheet.AutoFilterMode = True Then
        ReportSheet.AutoFilterMode = False
    End If
    
    Call RemoveTable(ReportSheet)
    
    'Reset headers. This creates two rows
    HeaderArray = Application.Transpose(ActiveWorkbook.Names("ReportColumnNamesList").RefersToRange.Value)
    TotalsArray = Application.Transpose(ActiveWorkbook.Names("ReportTotalsRowList").RefersToRange.Value)
    Call ResetTableHeaders(ReportSheet, ReportTableStart, HeaderArray)
    
    Set ReportLabelRange = ReportTableStart.EntireRow.Find("Label", , xlValues, xlWhole)
    Call ResetTableHeaders(ReportSheet, ReportLabelRange.Offset(1, 0), TotalsArray) 'The two columns before this are pulled from the cover sheet
    
    'Define where to put information and pull in in values from the cover sheet
    ReDim CenterInfoArray(1 To 3, 1 To 2)
        Set CenterInfoArray(1, 1) = ReportTableStart.EntireRow.Find("Center", , xlValues, xlWhole)
        Set CenterInfoArray(2, 1) = ReportTableStart.EntireRow.Find("Name", , xlValues, xlWhole)
        Set CenterInfoArray(3, 1) = ReportTableStart.EntireRow.Find("Date", , xlValues, xlWhole)
        
        Set CenterInfoArray(1, 2) = CoverSheet.Range("A:A").Find("Center", , xlValues, xlWhole)
        Set CenterInfoArray(2, 2) = CoverSheet.Range("A:A").Find("Name", , xlValues, xlWhole)
        Set CenterInfoArray(3, 2) = CoverSheet.Range("A:A").Find("Date", , xlValues, xlWhole)
    
    For i = 1 To UBound(CenterInfoArray)
        CenterInfoArray(i, 1).Offset(1, 0).Value = CenterInfoArray(i, 2).Offset(0, 1).Value
    Next i
    
    'Define table range and clear formats
    Set ReportTableRange = FindTableRange(ReportSheet)
    ReportTableRange.ClearFormats
    
    'Make a new table
    Set ReportTable = ReportSheet.ListObjects.Add(SourceType:=xlSrcRange, Source:=ReportTableRange, _
        xlListObjectHasHeaders:=xlYes)
    ReportTable.Name = "ReportTable"
    
    'Look for blank rows if there are more than two rows
    If ReportTable.DataBodyRange.Rows.Count < 2 Then
        GoTo FormatTable
    End If
        
    Set ReportLabelRange = ReportTable.ListColumns("Label").DataBodyRange
    Set DelRange = FindBlanks(ReportSheet, ReportLabelRange)
    
    If Not DelRange Is Nothing Then
        Call RemoveRows(ReportSheet, ReportTable.DataBodyRange, ReportLabelRange, DelRange)
        Set ReportTable = ReportSheet.ListObjects(1)
    End If
    
FormatTable:
    'Format
    ReportTable.ShowTableStyleRowStripes = False
    Call FormatReportTable(ReportSheet, ReportTable)
    
    'Add Marlett Boxes to everything but the Totals row
    Set BoxRange = ReportTable.ListColumns("Select").DataBodyRange
    Call AddMarlettBox(BoxRange)
    ReportTable.ListColumns("Select").DataBodyRange(1, 1).Font.Name = "Aptos Narrow" 'This can be anything except Marlett
    
    'Format the Date column
    ReportTable.ListColumns("Date").DataBodyRange.NumberFormat = "mm/dd/yyyy"
    
    'Autofit Description column
    ReportTable.ListColumns("Description").Range.EntireColumn.AutoFit

    'Return
    Set CreateReportTable = ReportTable

Footer:

End Function

Sub FormatTable(TargetSheet As Worksheet, NewTable As ListObject)
'Flag blanks and bad entries on Roster and Activity Sheets
'Different formatting on Report Page

    Dim CoverSheet As Worksheet
    Dim EthnicityRange As Range
    Dim GenderRange As Range
    Dim GradeRange As Range
    Dim CreditsRange As Range
    Dim MajorRange As Range
    Dim c As Range
    Dim FormulaString1 As String
    Dim FormulaString2 As String
    Dim IsCollegePrep As Boolean
    
    'If there are no rows, skip
    If CheckTable(TargetSheet) > 2 Then
        GoTo Footer
    End If

    'Check if this is for College Prep
    Set CoverSheet = Worksheets("Cover Page")
    
    IsCollegePrep = IsCollege

    'Blank cells flagged yellow, except in the first column
    With NewTable
        .DataBodyRange.FormatConditions.Delete
        .DataBodyRange.FormatConditions.Add Type:=xlBlanksCondition
        .ListColumns("Select").DataBodyRange.FormatConditions.Delete
        
        With .DataBodyRange.FormatConditions(1)
        .StopIfTrue = False
        .Interior.ColorIndex = 36
        End With
    End With

    'Validate demographics using tables on Ref Tables sheet
    Set EthnicityRange = NewTable.ListColumns("Ethnicity").DataBodyRange
    Set GenderRange = NewTable.ListColumns("Gender").DataBodyRange

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

    If IsCollegePrep = True Then
        'Grades work as both strings and numbers
        Set GradeRange = NewTable.ListColumns("Grade").DataBodyRange
        
        For Each c In GradeRange
            FormulaString1 = "=AND(COUNTIFS("
            FormulaString2 = "," & "Trim(" & c.Address & ")) < 1, NOT(ISBLANK(" & c.Address & ")))"
            c.FormatConditions.Add Type:=xlExpression, Formula1:=FormulaString1 & "GradeList" & FormulaString2
            With c.FormatConditions(2)
                .StopIfTrue = False
                .Interior.Color = vbRed
            End With
        Next c
    Else
        'Credit hours don't need the reference table, add majors to validation
        Set CreditsRange = NewTable.ListColumns("Credits").DataBodyRange
        Set MajorRange = NewTable.ListColumns("Major").DataBodyRange
        
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
    End If

Footer:

End Sub

Sub FormatReportTable(ReportSheet As Worksheet, ReportTable As ListObject)
'Adjust font, color, number formats for the table on the Report Page

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

End Sub

Sub RemoveTable(TargetSheet As Worksheet)
'Unlists all table objects and removes formatting

    Dim OldTableRange As Range
    Dim OldTable As ListObject
    
    Call UnprotectSheet(TargetSheet)
    
    For Each OldTable In TargetSheet.ListObjects
        Set OldTableRange = OldTable.Range
        
        OldTable.Unlist
        OldTableRange.FormatConditions.Delete
        OldTableRange.Borders.LineStyle = Excel.XlLineStyle.xlLineStyleNone
    Next OldTable

End Sub

Sub ResetTableHeaders(TargetSheet As Worksheet, TargetCell As Range, TargetFields As Variant)
'Insert the default column names starting at the passed cell

    Dim FieldRange As Range
    
    Set FieldRange = TargetSheet.Range(TargetCell, TargetCell.Offset(0, UBound(TargetFields) - 1))
    FieldRange.Value = TargetFields

End Sub










