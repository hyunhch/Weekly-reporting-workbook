Attribute VB_Name = "SetupSubs"
Option Explicit

Sub Tester()

'Call ChooseProgram("University Ref")
'Call ChooseProgram("College Ref")
Call ChooseProgram("Transfer Ref")

End Sub

Sub TesterClearTables()

    Dim RefSheet As Worksheet
    Dim CoverSheet As Worksheet
    Dim RecordsSheet As Worksheet
    Dim ReportSheet As Worksheet
    Dim RosterSheet As Worksheet
    Dim i As Long
    Dim ClearTable As ListObject
    Dim btn As Button
    Dim ButtonArray As Variant
    
    Set CoverSheet = Worksheets("Cover Page")
    Set ReportSheet = Worksheets("Report Page")
    Set RosterSheet = Worksheets("Roster Page")
    Set RecordsSheet = Worksheets("Records Page")
    
    On Error Resume Next
    
    Call UnprotectSheet(CoverSheet)
    With CoverSheet
        .Cells.ClearContents
        .Cells.ClearFormats
        .Buttons.Delete
        .Columns.UseStandardWidth = True
    End With
    
    Call UnprotectSheet(ReportSheet)
    With ReportSheet
        .Cells.ClearContents
        .Cells.ClearFormats
        .Buttons.Delete
        .Columns.UseStandardWidth = True
    End With
    
    Call UnprotectSheet(RosterSheet)
    With RosterSheet
        .Cells.ClearContents
        .Cells.ClearFormats
        .Buttons.Delete
        .Columns.UseStandardWidth = True
    End With
    
    Call UnprotectSheet(RecordsSheet)
    With RecordsSheet
        .Cells.ClearContents
        .Cells.ClearFormats
        .Buttons.Delete
        .Columns.UseStandardWidth = True
    End With
    
    If Not Worksheets(1).Name = "University Ref" Then
        Worksheets(1).Name = "University Ref"
    ElseIf Not Worksheets(2).Name = "Transfer Ref" Then
        Worksheets(2).Name = "Transfer Ref"
    ElseIf Not Worksheets(3).Name = "College Ref" Then
        Worksheets(3).Name = "College Ref"
    End If
    
    Set RefSheet = Worksheets("University Ref")

    For Each ClearTable In RefSheet.ListObjects
        If Not ClearTable.Name = "UniversityTableGen" And Not ClearTable.Name = "UniversityRangeGen" Then
            ClearTable.Unlist
        End If
    Next ClearTable
    
    
    Set RefSheet = Worksheets("Transfer Ref")
        
    For Each ClearTable In RefSheet.ListObjects
        If Not ClearTable.Name = "TransferTableGen" And Not ClearTable.Name = "TransferRangeGen" Then
            ClearTable.Unlist
        End If
    Next ClearTable
    
    
    Set RefSheet = Worksheets("College Ref")
        
    For Each ClearTable In RefSheet.ListObjects
        If Not ClearTable.Name = "CollegeTableGen" And Not ClearTable.Name = "CollegeRangeGen" Then
            ClearTable.Unlist
        End If
    Next ClearTable
    
    'Remove named ranges
    For i = ThisWorkbook.Names.Count To 1 Step -1
        If Not Right(ThisWorkbook.Names(i), 3) = "Gen" Then
            ThisWorkbook.Names(i).Delete
        End If
    Next
    
    'Put the button to choose a program on the CoverSheet
    ReDim ButtonArray(1 To 4)
        ButtonArray(1) = "B2:D3"
        ButtonArray(2) = "CoverChooseProgramButton"
        ButtonArray(3) = "Choose Program"
        ButtonArray(4) = "ButtonCoverChooseProgram"

    Call MakeButton(CoverSheet, ButtonArray)

End Sub

Sub ChooseProgram(ProgramString As String)
'User selects the program from a dropdown list
'Set up table, ranges, and references specific to that program, then disable the ability to select

        Dim RefSheet As Worksheet
        Dim ReportSheet As Worksheet
        Dim RosterSheet As Worksheet
        Dim CoverSheet As Worksheet
        Dim sh As Worksheet
        Dim StartCell As Range
        Dim StopCell As Range
        Dim BotCell As Range
        Dim TableRange As Range
        Dim SearchRange As Range
        Dim CoverTitleRange As Range
        Dim CoverRefRange As Range
        Dim CoverCenterRange As Range
        Dim c As Range
        Dim TrimString As String
        Dim HeaderArray() As Variant
        Dim TotalsArray() As Variant
        Dim TableGenTable As ListObject
        Dim RangeGenTable As ListObject
        
        'Find the refence sheet for the selected program
        Set RefSheet = Worksheets(ProgramString)
        RefSheet.Name = "Ref Tables"
        
        If RefSheet Is Nothing Then 'This shouldn't happen
            GoTo Footer
        End If

        'Make and name reference tables. Each table has an empty column between it and the next
        'A table for table names and for range names/references already exist
        TrimString = Left(ProgramString, InStr(1, ProgramString, " ") - 1)
        
        With RefSheet
            Set TableGenTable = .ListObjects(TrimString & "TableGen")
            Set RangeGenTable = .ListObjects(TrimString & "RangeGen")
            
            Set SearchRange = TableGenTable.ListColumns("First Header").DataBodyRange
            
            'The TableGenTable as the names of each header in the 1st column. Find the header, first blank column after, and last row
            For Each c In SearchRange
                Set StartCell = .Range("1:1").Find(c.Value, , xlValues, xlWhole)
                If Not StartCell Is Nothing Then
                    'Define table range
                    Set StopCell = .Range(StartCell, Cells(1, Columns.Count).Address).Find("", , xlValues, xlWhole) 'This is a blank cell one past the last column
                    Set BotCell = StartCell.EntireColumn.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
                    Set TableRange = StartCell.Resize(BotCell.Row, StopCell.Column - StartCell.Column)
                    
                    'Make and name table
                    .ListObjects.Add(xlSrcRange, TableRange, , xlYes).Name = c.Offset(0, -1).Value 'Names of tables are stored one to the left
                End If
            Next c
            
            'Define named ranges
            Set SearchRange = RangeGenTable.ListColumns("Range Name").DataBodyRange
            
            For Each c In SearchRange
                ThisWorkbook.Names.Add Name:=c.Value, RefersTo:=.Range("=" & c.Offset(0, 1).Value)
            Next c

        End With

    'Populate the Cover Page
    Set CoverSheet = Worksheets("Cover Page")
    
    Call UnprotectSheet(CoverSheet)
    Call CoverSheetText(RefSheet, CoverSheet, ProgramString)
    Call CoverSheetButtons(ProgramString)

    'Make table on Report Page and add buttons
    Set ReportSheet = Worksheets("Report Page")
    Set c = ReportSheet.Range("A6") 'Where the headers begin

    Call UnprotectSheet(ReportSheet)
    c.Value = "Select" 'This is what the following sub looks for
    Call MakeReportTable
    Call ReportSheetButtons
    
    'Put in headers for Roster Page and add buttons. Making the table will happen when it's parsed
    Set RosterSheet = Worksheets("Roster Page")
    Set c = RosterSheet.Range("A6") 'Where the headers begin
    
    Call UnprotectSheet(RosterSheet)
    HeaderArray = Application.Transpose(ActiveWorkbook.Names("RosterHeadersList").RefersToRange.Value)
    Call TableResetHeaders(RosterSheet, c, HeaderArray)
    Call MakeTable(RosterSheet)
    Call RosterSheetButtons
    
    'Text on the Records Sheet
    Call RecordsSheetText
    
    'Make sure the workbook can be edited
    Call ResetProtection
    
Footer:

End Sub

Sub CoverSheetButtons(ProgramString)
    
    Dim CoverSheet As Worksheet
    Dim i As Long
    Dim j As Long
    Dim ButtonArray As Variant
    Dim TempArray As Variant
    
    Set CoverSheet = Worksheets("Cover Page")
    
    'SharePoint
    ReDim TempArray(1 To 4)
        TempArray(1) = "A7:C8"
        TempArray(2) = "CoverSharePointExportButton"
        TempArray(3) = "Submit to SharePoint"
        TempArray(4) = "ButtonCoverSharePointExport"
        
    ReDim ButtonArray(1 To 1)
    ButtonArray(1) = TempArray
    
    'Local save
    ReDim TempArray(1 To 4)
        TempArray(1) = "A10:C11"
        TempArray(2) = "CoverSaveCopyButton"
        TempArray(3) = "Save a Copy"
        TempArray(4) = "ButtonCoverSaveCopy"

    ReDim Preserve ButtonArray(1 To 2)
    ButtonArray(2) = TempArray
    
    'Import
    ReDim TempArray(1 To 4)
        TempArray(1) = "A13:C14"
        TempArray(2) = "CoverImportButton"
        TempArray(3) = "Import Records"
        TempArray(4) = "ButtonCoverImport"

    ReDim Preserve ButtonArray(1 To 3)
    ButtonArray(3) = TempArray

    For i = 1 To UBound(ButtonArray)
        TempArray = ButtonArray(i)
        Call MakeButton(CoverSheet, TempArray)
    Next i


End Sub

Sub CoverSheetInitialize()
'Puts a button to choose the program

    Dim CoverSheet As Worksheet
    Dim ButtonArray As Variant
    
    Set CoverSheet = Worksheets("Cover Page")
    
    ReDim ButtonArray(1 To 4)
        ButtonArray(1) = "A1:C3"
        ButtonArray(2) = "CoverChooseProgramButton"
        ButtonArray(3) = "Choose Program"
        ButtonArray(4) = "ButtonCoverChooseProgram"
    
    Call MakeButton(CoverSheet, ButtonArray)
    
End Sub

Sub CoverSheetText(RefSheet As Worksheet, CoverSheet As Worksheet, ProgramString As String)
'Text, formatting, tables for CoverSheet

    Dim ChangeSheet As Worksheet
    Dim TextRange As Range
    Dim DateRange As Range
    Dim CenterRange As Range
    Dim CopyRange As Range
    Dim PasteRange As Range
    Dim c As Range
    Dim d As Range
    Dim i As Long
    Dim BookTitle As String
    Dim BookEdition As String
    Dim TextString As String
    Dim TableNameString As String
    Dim TextArray() As String
    Dim TempTable As ListObject
    Dim TableNameArray As Variant
    
    Set ChangeSheet = Worksheets("Change Log")
    Set CoverSheet = Worksheets("Cover Page")
    
    'Unprotect. This shouldn't ever be needed
    Call UnprotectSheet(CoverSheet)
    
    'Define the title and edition
    Select Case ProgramString
        Case "University Ref"
            BookTitle = "MESA University Weekly Report"
            
        Case "Transfer Ref"
            BookTitle = "Transfer Prep Weekly Report"
            
        Case "College Ref"
            BookTitle = "College Prep Weekly Report"
    End Select
    
    'BookEdition = GetEdition() borks if the file is renamed
    BookEdition = ChangeSheet.Range("A1").Value
        If Not Len(BookEdition) > 0 Then
            BookEdition = "Version Unknown"
        End If

    'Insert text
    With CoverSheet
        Set TextRange = .Range("A1:A5")
        
        TextString = BookTitle & ";" & "Version " & BookEdition & ";Name;Date;Center"
        TextArray = Split(TextString, ";")
        TextRange.Value = Application.Transpose(TextArray)
    
        'Date validation and a dropdown menu for the center
        Set DateRange = .Range("A:A").Find("Date", , xlValues, xlWhole)
        Set CenterRange = .Range("A5").Find("Center", , xlValues, xlWhole)

        Call DateValidation(CoverSheet, DateRange.Offset(0, 1))
        Call CenterDropdown(CoverSheet, CenterRange.Offset(0, 1))
    End With
    
    'Add formatting. No lines under the first two rows
    i = 1
    For Each c In TextRange
        Set d = c.Resize(1, 3)
        
        c.Font.Bold = True
        d.WrapText = False
        
        If i > 2 Then
            c.HorizontalAlignment = xlRight
            d.Borders(xlEdgeBottom).LineStyle = xlContinuous
            d.Borders(xlEdgeBottom).Weight = xlMedium
        End If
    
        i = i + 1
    Next c
    
    'Add reference tables
    Set c = CoverSheet.Range("H1")
    
    If ProgramString = "College Ref" Then
        ReDim TableNameArray(1 To 3)
            TableNameArray(3) = "GradeTable"
    Else
        ReDim TableNameArray(1 To 5)
            TableNameArray(3) = "MajorTable"
            TableNameArray(4) = "FirstGenerationTable"
            TableNameArray(5) = "LowIncomeTable"
    End If
    
    TableNameArray(1) = "RaceTable"
    TableNameArray(2) = "GenderTable"
    
    For i = 1 To UBound(TableNameArray)
        TableNameString = TableNameArray(i)
        Set TempTable = RefSheet.ListObjects(TableNameString)
        Set CopyRange = TempTable.Range
        Set PasteRange = c.Resize(TempTable.Range.Rows.Count, 1).Offset(0, (i - 1) * 2) 'Put a space between each table
        
        With PasteRange
            .Value(11) = CopyRange.Value(11)
            .HorizontalAlignment = xlLeft
            .BorderAround LineStyle:=xlContinuous, Weight:=xlThin
            .Columns.AutoFit
        End With
    Next i

Footer:
    
End Sub

Sub RecordsSheetText()
'Called when a program is chosen
    
    Dim RecordsSheet As Worksheet
    Dim c As Range
    Dim i As Long
    Dim TextArray As Variant
    
    Set RecordsSheet = Worksheets("Records Page")
    
    ReDim TextArray(1 To 2, 1 To 1)
    i = 1
    
    ReDim Preserve TextArray(1 To 2, 1 To i)
    TextArray(1, i) = "A6"
    TextArray(2, i) = "First"
    i = i + 1
    
    ReDim Preserve TextArray(1 To 2, 1 To i)
    TextArray(1, i) = "B6"
    TextArray(2, i) = "Last"
    i = i + 1
    
    ReDim Preserve TextArray(1 To 2, 1 To i)
    TextArray(1, i) = "A7"
    TextArray(2, i) = "H BREAK"
    i = i + 1
    
    ReDim Preserve TextArray(1 To 2, 1 To i)
    TextArray(1, i) = "C1"
    TextArray(2, i) = "Label"
    i = i + 1
    
    ReDim Preserve TextArray(1 To 2, 1 To i)
    TextArray(1, i) = "C2"
    TextArray(2, i) = "Practice"
    i = i + 1
    
    ReDim Preserve TextArray(1 To 2, 1 To i)
    TextArray(1, i) = "C3"
    TextArray(2, i) = "Category"
    i = i + 1
    
    ReDim Preserve TextArray(1 To 2, 1 To i)
    TextArray(1, i) = "C4"
    TextArray(2, i) = "Date"
    i = i + 1
    
    ReDim Preserve TextArray(1 To 2, 1 To i)
    TextArray(1, i) = "C5"
    TextArray(2, i) = "Description"
    i = i + 1

    ReDim Preserve TextArray(1 To 2, 1 To i)
    TextArray(1, i) = "D1"
    TextArray(2, i) = "V BREAK"
    i = i + 1

    For i = 1 To UBound(TextArray, 2)
        Set c = RecordsSheet.Range(TextArray(1, i))
        c.Value = TextArray(2, i)
    Next i


End Sub

Sub RosterSheetButtons()
'Called when the program is chosen

    Dim RosterSheet As Worksheet
    Dim i As Long
    Dim ButtonArray As Variant
    Dim TempArray As Variant

    Set RosterSheet = Worksheets("Roster Page")
    
    ReDim ButtonArray(1 To 1)
    i = 1
    
    'Select All
    ReDim TempArray(1 To 4)
        TempArray(1) = "A5:B5"
        TempArray(2) = "SelectAllButton"
        TempArray(3) = "Select All"
        TempArray(4) = "ButtonRosterSelectAll"
        
    ReDim Preserve ButtonArray(1 To i)
    ButtonArray(i) = TempArray
    i = i + 1
    
    'Delete Row
    ReDim TempArray(1 To 4)
        TempArray(1) = "D5:E5"
        TempArray(2) = "RemoveSelectedButton"
        TempArray(3) = "Delete Row"
        TempArray(4) = "ButtonRosterRemoveSelected"
        
    ReDim Preserve ButtonArray(1 To i)
    ButtonArray(i) = TempArray
    i = i + 1

    'Select Activity
    ReDim TempArray(1 To 4)
        TempArray(1) = "G4:H5"
        TempArray(2) = "RosterNewActivityFormButton"
        TempArray(3) = "New Activity"
        TempArray(4) = "ButtonRosterNewActivity"
        
    ReDim Preserve ButtonArray(1 To i)
    ButtonArray(i) = TempArray
    i = i + 1

    'Load Activity
    ReDim TempArray(1 To 4)
        TempArray(1) = "G2:H2"
        TempArray(2) = "RosterLoadActivityFormButton"
        TempArray(3) = "Load Activity"
        TempArray(4) = "ButtonRosterLoadActivity"
        
    ReDim Preserve ButtonArray(1 To i)
    ButtonArray(i) = TempArray
    i = i + 1

    'Add Students
    ReDim TempArray(1 To 4)
        TempArray(1) = "G1:H1"
        TempArray(2) = "RosterAddStudentsFormButton"
        TempArray(3) = "Add to Activity"
        TempArray(4) = "ButtonRosterAddSelected"
        
    ReDim Preserve ButtonArray(1 To i)
    ButtonArray(i) = TempArray
    i = i + 1

    'Parse Roster
    ReDim TempArray(1 To 4)
        TempArray(1) = "A1:B2"
        TempArray(2) = "RosterParseButton"
        TempArray(3) = "Parse Roster"
        TempArray(4) = "ButtonRosterParse"
        
    ReDim Preserve ButtonArray(1 To i)
    ButtonArray(i) = TempArray
    i = i + 1

    'Clear Roster
    ReDim TempArray(1 To 4)
        TempArray(1) = "D1:E1"
        TempArray(2) = "RosterClearButton"
        TempArray(3) = "Clear Roster"
        TempArray(4) = "ButtonRosterClear"
        
    ReDim Preserve ButtonArray(1 To i)
    ButtonArray(i) = TempArray
    i = i + 1

    For i = 1 To UBound(ButtonArray)
        TempArray = ButtonArray(i)
        Call MakeButton(RosterSheet, TempArray)
    Next i

End Sub

Sub ReportSheetButtons()
'Called when the program is chosen

    Dim ReportSheet As Worksheet
    Dim i As Long
    Dim TempArray As Variant
    Dim ButtonArray As Variant
    
    Set ReportSheet = Worksheets("Report Page")
    
    ReDim ButtonArray(1 To 1)
    i = 1
    
    'Select All
    ReDim TempArray(1 To 4)
        TempArray(1) = "A5:B5"
        TempArray(2) = "SelectAllButton"
        TempArray(3) = "Select All"
        TempArray(4) = "ButtonReportSelectAll"
        
    ReDim Preserve ButtonArray(1 To i)
    ButtonArray(i) = TempArray
    i = i + 1
        
    'Pull Totals
    ReDim TempArray(1 To 4)
        TempArray(1) = "A1:B2"
        TempArray(2) = "ReportTabulateTotalsButton"
        TempArray(3) = "Tabulate Totals"
        TempArray(4) = "ButtonReportTabTotals"
        
    ReDim Preserve ButtonArray(1 To i)
    ButtonArray(i) = TempArray
    i = i + 1
    
    'Clear the Report
    ReDim TempArray(1 To 4)
        TempArray(1) = "D1:E2"
        TempArray(2) = "ReportClearButton"
        TempArray(3) = "Clear Report"
        TempArray(4) = "ButtonReportClear"
        
    ReDim Preserve ButtonArray(1 To i)
    ButtonArray(i) = TempArray
    i = i + 1
    
    'Tabulate activities
    ReDim TempArray(1 To 4)
        TempArray(1) = "C1:C2"
        TempArray(2) = "ReportTabulateFormButton"
        TempArray(3) = "Tabulate Activities"
        TempArray(4) = "ButtonReportTabActivities"
        
    ReDim Preserve ButtonArray(1 To i)
    ButtonArray(i) = TempArray
    i = i + 1
    
    'Remove row
    ReDim TempArray(1 To 4)
        TempArray(1) = "D4:E5"
        TempArray(2) = "RemoveSelectedButton"
        TempArray(3) = "Delete Row"
        TempArray(4) = "ButtonReportRemoveSelected"
        
    ReDim Preserve ButtonArray(1 To i)
    ButtonArray(i) = TempArray
    i = i + 1

    For i = 1 To UBound(ButtonArray)
        TempArray = ButtonArray(i)
        Call MakeButton(ReportSheet, TempArray)
    Next i

End Sub


