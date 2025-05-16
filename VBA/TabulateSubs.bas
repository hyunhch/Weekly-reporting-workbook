Attribute VB_Name = "TabulateSubs"
Option Explicit
Option Compare Text

Function DemoTabulate(SearchRange As Range, SearchType As String) As Variant
'Returns an array with the values in the passed range tabulated
'Uses a dictionary so no reference to the columns needs to be done
'SearchType is for renaming "Other" values. This is why this is done piecemeal instead of all at once

    Dim c As Range
    Dim i As Long
    Dim j As Long
    Dim RenameString As String
    Dim TypeArray As Variant
    Dim DemoElement As Variant
    Dim SearchArray As Variant
    Dim CountArray As Variant
    Dim DemoDict As Scripting.Dictionary

    Set DemoDict = New Scripting.Dictionary
    
    'Ignore case
    DemoDict.CompareMode = TextCompare
    
    'Read values into an array. This should always be 1 dimensional since only one column is passed
    'Need to loop for any non-contiguous range
    ReDim SearchArray(1 To SearchRange.Cells.Count)
    
    i = 1
    For Each c In SearchRange
        SearchArray(i) = c.Value
        i = i + 1
    Next c
    
    'Different procedure for Credits
    If SearchType = "Credits" Then
        GoTo TabulateCredits
    ElseIf SearchType = "Ethnicity" Then
        SearchType = "Race"
    End If
    
    'Read into the dictionary
    For i = 1 To UBound(SearchArray)
        DemoElement = SearchArray(i)
        If Not DemoDict.Exists(DemoElement) Then
            DemoDict.Add DemoElement, 1
        Else
            DemoDict(DemoElement) = DemoDict(DemoElement) + 1
        End If
    Next i
    
    'First Generation and Low Income don't have an other category
    If SearchType = "First Generation" Or SearchType = "Low Income" Then
        If DemoDict.Exists("Yes") Then
            DemoDict.Key("Yes") = SearchType
        End If
        
        GoTo SkipOther
    End If
    
    'Rename the "Other" key. Insert one if it doesn't exist
    If Not DemoDict.Exists("Other") Then
        DemoDict.Add "Other", 0
    End If
    
    'If DemoDict.Exists("Other") Then
        RenameString = "Other " & SearchType
        DemoDict.Key("Other") = RenameString
    'End If
    
SkipOther:
    'Read into an array for counting
    ReDim CountArray(1 To DemoDict.Count, 1 To 2)
    
    i = 0
    For Each DemoElement In DemoDict.Keys
        i = i + 1
        CountArray(i, 1) = DemoElement
        CountArray(i, 2) = DemoDict(DemoElement)
    Next DemoElement
    
    GoTo ReturnArray
    
TabulateCredits:
    'Credits go into buckets
    ReDim CountArray(1 To 4, 1 To 2) 'Should probably make this programmatic
        CountArray(1, 1) = "<45"
        CountArray(2, 1) = "45-90"
        CountArray(3, 1) = ">90"
        CountArray(4, 1) = "Other Credits"
        
    For i = 1 To UBound(SearchArray)
        j = SearchArray(i)
        
        If IsEmpty(j) Or j = 0 Then 'VBA will return true for <45 on empty cells
            CountArray(4, 2) = CountArray(4, 2) + 1
        ElseIf Not (IsNumeric(j)) Then
            CountArray(4, 2) = CountArray(4, 2) + 1
        ElseIf j < 45 Then
            CountArray(1, 2) = CountArray(1, 2) + 1
        ElseIf j <= 90 Then
            CountArray(2, 2) = CountArray(2, 2) + 1
        ElseIf j > 90 Then
            CountArray(3, 2) = CountArray(3, 2) + 1
        Else
            CountArray(4, 2) = CountArray(4, 2) + 1 'To catch anything else
        End If
    Next i
    
    GoTo ReturnArray
  
ReturnArray:
    DemoTabulate = CountArray

End Function

Sub RetabulateReport()
'Retabulates the activities on the ReportSheet

    Dim ReportSheet As Worksheet
    Dim ReportLabelRange As Range
    Dim c As Range
    Dim ReportTable As ListObject

    Set ReportSheet = Worksheets("Report Page")
    
    'Make sure there's a table with at least two ListRows
    If CheckTable(ReportSheet) > 2 Then
        GoTo Footer
    End If

    Call UnprotectSheet(ReportSheet)

    'Define area to search and loop through
    Set ReportTable = ReportSheet.ListObjects(1)
    Set ReportLabelRange = FindReportLabel(ReportSheet)
        
    For Each c In ReportLabelRange
        Call TabulateActivity(c)
    Next c

Footer:

End Sub

Sub TabulateReportTotals()
'Called from a button and when the roster is parsed
'Not entirely programmatic yet. Adding a tabulation table on the RefSheet might work

    Dim RosterSheet As Worksheet
    Dim ReportSheet As Worksheet
    Dim RecordsSheet As Worksheet
    Dim ReportHeaderRange As Range
    Dim c As Range
    Dim i As Long
    Dim RosterTable As ListObject
    Dim ReportTable As ListObject
    Dim PasteCell As Range
    Dim TempRange As Range
    Dim TempArray As Variant
    
    Set RosterSheet = Worksheets("Roster Page")
    Set ReportSheet = Worksheets("Report Page")
    Set RecordsSheet = Worksheets("Records Page")
    
    'Make sure we have students in records
    If CheckRecords(RecordsSheet) = 2 Or CheckRecords(RecordsSheet) = 4 Then
        GoTo Footer
    'Make sure the RosterSheet has a table with students
    ElseIf CheckTable(RosterSheet) > 2 Then
        GoTo Footer
    'Make sure there's a table on the ReportSheet
    ElseIf CheckTable(ReportSheet) > 3 Then
        Call CreateReportTable
    End If
    
    Set RosterTable = RosterSheet.ListObjects(1)
    Set ReportTable = ReportSheet.ListObjects(1)
    Set ReportHeaderRange = ReportTable.HeaderRowRange
    
    Call UnprotectSheet(ReportSheet)
    
    'Clear the contents
    Call ClearReportTotals
    
    'Find the cell under "Label"
    Set PasteCell = ReportHeaderRange.Find("Label", , xlValues, xlWhole).Offset(1, 0) 'The columns were just reset, so this should always be here
    
    'Add information from the coversheet
    TempArray = GetSubmissionInfo(ReportSheet, "Yes")
    Call CopyToReport(ReportSheet, PasteCell, TempArray)
    
    'Add the other non-numeric values using a named range. The header values are one cell to the left
    Set TempRange = Range("ReportTotalsRowList")
    ReDim TempArray(1 To TempRange.Cells.Count, 1 To 2)
    
    i = 1
    For Each c In TempRange
        TempArray(i, 1) = c.Offset(0, -1).Value
        TempArray(i, 2) = c.Value

        i = i + 1
    Next c

    Call CopyToReport(ReportSheet, PasteCell, TempArray)

    'Add race, gender. Grade for College Prep, Credits and Major for Transfer Prep and MESA U
    Call TabulateHelper(ReportSheet, RosterSheet, PasteCell)
    
    'Fill in the totals cell
    Set c = ReportHeaderRange.Find("Total", , xlValues, xlWhole).Offset(1, 0)
    c.Value = RosterTable.ListRows.Count
    
    'Apply bold font
    ReportHeaderRange.Offset(1, 0).Font.Bold = True
    
Footer:

End Sub

Sub TabulateHelper(ReportSheet As Worksheet, RosterSheet As Worksheet, PasteCell As Range, Optional NameRange As Range)
'Tabulates every category since this needs to be done both for the totals row and tabulating and activity
'Passing NameRange limits tabulation to only those names. This should be a range on the RosterSheet

    Dim CoverSheet As Worksheet
    Dim RosterTable As ListObject
    Dim TempRange As Range
    Dim c As Range
    Dim i As Long
    Dim SearchTerm As String
    Dim TempArray() As Variant
    Dim SearchTermArray As Variant
    Dim NamesPassed As Boolean
    
    Set CoverSheet = Worksheets("Cover Page")
    Set RosterSheet = Worksheets("Roster Page")
    Set RosterTable = RosterSheet.ListObjects(1)
    
    If NameRange Is Nothing Then
        NamesPassed = False
    Else
        NamesPassed = True
    End If
    
    'Go through each category to tabulate and push them to the ReportSheet
    If IsCollege() = True Then
        ReDim SearchTermArray(1 To 3, 1 To 2)
        SearchTermArray(3, 1) = "Grade"
        SearchTermArray(3, 2) = Range("GradeList").Cells.Count
    Else
        ReDim SearchTermArray(1 To 6, 1 To 2)
        SearchTermArray(3, 1) = "Credits"
        SearchTermArray(4, 1) = "Major"
        SearchTermArray(5, 1) = "First Generation"
        SearchTermArray(6, 1) = "Low Income"
        SearchTermArray(3, 2) = Range("CreditsList").Cells.Count
        SearchTermArray(4, 2) = Range("MajorList").Cells.Count
        SearchTermArray(5, 2) = Range("FirstGenerationList").Cells.Count
        SearchTermArray(6, 2) = Range("LowIncomeList").Cells.Count
    End If
  
    SearchTermArray(1, 1) = "Ethnicity"
    SearchTermArray(2, 1) = "Gender"
    SearchTermArray(1, 2) = Range("EthnicityList").Cells.Count
    SearchTermArray(2, 2) = Range("GenderList").Cells.Count
        
    For i = 1 To UBound(SearchTermArray)
        SearchTerm = SearchTermArray(i, 1)
        Set TempRange = RosterTable.ListColumns(SearchTerm).DataBodyRange
        
        If NamesPassed = True Then
            Set c = NameRange.Offset(0, TempRange.Column - NameRange.Column)
            Set TempRange = c
        End If
        
        Erase TempArray
        TempArray = DemoTabulate(TempRange, SearchTerm)
        Call CopyToReport(ReportSheet, PasteCell, TempArray)
    Next i

End Sub

Sub TabulateActivity(LabelCell As Range)
'Pushes tabulation to the report page for a single activity
'Called automatically when an activity is saved or a student removed from the Roster Page

    Dim RecordsSheet As Worksheet
    Dim ReportSheet As Worksheet
    Dim RosterSheet As Worksheet
    Dim RecordsLabelRange As Range
    Dim ReportLabelRange As Range
    Dim TabulateRange As Range
    Dim TotalRange As Range
    Dim PasteCell As Range
    Dim c As Range
    Dim LabelString As String
    Dim TempArray() As Variant
    Dim RosterTable As ListObject
    Dim ReportTable As ListObject
    
    Set RecordsSheet = Worksheets("Records Page")
    Set ReportSheet = Worksheets("Report Page")
    Set RosterSheet = Worksheets("Roster Page")
    
    'Make sure we have students and activities in records
    If CheckRecords(RecordsSheet) > 2 Then
        GoTo Footer
    'Make sure the RosterSheet has a table with students
    ElseIf CheckTable(RosterSheet) > 2 Then
        GoTo Footer
    'Make sure there's a table on the ReportSheet
    ElseIf CheckTable(ReportSheet) > 3 Then
        Call CreateReportTable
    End If
    
    Set RosterTable = RosterSheet.ListObjects(1)
    Set ReportTable = ReportSheet.ListObjects(1)
    
    'Make sure the activity is in the RecordsSheet
    Set RecordsLabelRange = FindRecordsLabel(RecordsSheet, LabelCell)
    
    If RecordsLabelRange Is Nothing Then 'This shouldn't happen
        GoTo Footer
    End If
    
    'See if the activity is already on the ReportSheet
    Set ReportLabelRange = FindReportLabel(ReportSheet, LabelCell)
    
    If ReportLabelRange Is Nothing Then 'Activities present but label not found
        Set ReportLabelRange = FindReportLabel(ReportSheet)
        Set c = ReportLabelRange.Resize(1, 1)
        Set PasteCell = c.Offset(ReportLabelRange.Rows.Count, 0)
    ElseIf ReportLabelRange.Value = "Total" Then 'No activities
        Set PasteCell = ReportLabelRange.Offset(1, 0)
    Else
        Set PasteCell = ReportLabelRange
    End If
    
    'Clear everything currently in the row
    LabelString = PasteCell.Value
    PasteCell.EntireRow.ClearContents
    PasteCell.Value = LabelString
    
    'Put on the activity information
    Call UnprotectSheet(ReportSheet)
    
    TempArray = GetSubmissionInfo(ReportSheet)
    Call CopyToReport(ReportSheet, PasteCell, TempArray)
    Erase TempArray
    
    TempArray = GetActivityInfo(RecordsSheet, ReportSheet, LabelCell)
    Call CopyToReport(ReportSheet, PasteCell, TempArray)
    
    'Find students to tabulate on the Roster sheet
    Set TabulateRange = FindTabulateRange(RosterSheet, RecordsSheet, RecordsLabelRange)
    If TabulateRange Is Nothing Then
        GoTo RemakeTable
    End If
    
    'Pass the demographics for tabulation and add total
    Set TotalRange = ReportTable.HeaderRowRange.Find("Total", , xlValues, xlWhole)
    
    Call TabulateHelper(ReportSheet, RosterSheet, PasteCell, TabulateRange)
    ReportSheet.Cells(PasteCell.Row, TotalRange.Column) = TabulateRange.Cells.Count
    
RemakeTable:
    'Expand the table, add Marlett boxes
    Set ReportTable = CreateReportTable

    Call AddMarlettBox(ReportTable.ListColumns("Select").DataBodyRange)

Footer:

End Sub

Sub TabulateAllActivities()
'Loop through all activities on the Records Page and tabulate
'Only tabulate the totals if there are none

    Dim RecordsSheet As Worksheet
    Dim ReportSheet As Worksheet
    Dim RecordsLabelRange As Range
    Dim ReportLabelRange As Range
    Dim c As Range
    
    Set RecordsSheet = Worksheets("Records Page")
    
    'Tabulate the totals row
    Call TabulateReportTotals
    
    'Define the range of labels on the Report and Records sheets
    Set RecordsLabelRange = FindRecordsLabel(RecordsSheet)
    
    'If there are no saved activities
    If RecordsLabelRange.Cells.Count = 1 Then
        If RecordsLabelRange.Value = "V BREAK" Then
            GoTo Footer
        End If
    End If
    
    'Loop through labels
    For Each c In RecordsLabelRange
        'These get turned back on somewhere in the TabulateActivity sub
        Application.EnableEvents = False
        Application.ScreenUpdating = False
        Application.DisplayAlerts = False
        
        Call TabulateActivity(c)
    Next c

Footer:

End Sub


