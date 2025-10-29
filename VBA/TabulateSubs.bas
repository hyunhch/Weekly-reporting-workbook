Attribute VB_Name = "TabulateSubs"
Option Explicit
Option Compare Text

Function DemoTabulate(SearchRange As Range, SearchType As String) As Variant
'Redirect

    Err.Raise vbObjectError + 513, , "Wrong function. TabulateDemo"
    
End Function

Sub RetabulateReport()

Err.Raise vbObjectError + 513, , "Wrong function. TabulateReportActivities"

End Sub

Sub TabulateActivity(LabelCell As Range)
'Pushes tabulation to the report page for a single activity
'Called automatically when an activity is saved or a student removed from the Roster Page

    Dim RecordsSheet As Worksheet
    Dim ReportSheet As Worksheet
    Dim RosterSheet As Worksheet
    Dim RecordsLabelRange As Range
    Dim ReportLabelRange As Range
    Dim AttendanceRange As Range
    Dim TabulateRange As Range
    Dim TotalRange As Range
    Dim PasteCell As Range
    Dim c As Range
    Dim i As Long
    Dim LabelString As String
    Dim TempArray() As Variant
    Dim ValueArray() As Variant
    Dim RosterTable As ListObject
    Dim ReportTable As ListObject
    
    Set RecordsSheet = Worksheets("Records Page")
    Set ReportSheet = Worksheets("Report Page")
    Set RosterSheet = Worksheets("Roster Page")
    
    Call UnprotectSheet(ReportSheet)
    
    'Make sure we have students and activities in records
    If CheckRecords(RecordsSheet) > 2 Then
        GoTo Footer
    'Make sure the RosterSheet has a table with students
    ElseIf CheckTable(RosterSheet) > 2 Then
        GoTo Footer
    'Make sure there's a table on the ReportSheet
    ElseIf CheckReport(ReportSheet) > 3 Then
        Call MakeReportTable
    End If
    
    Set RosterTable = RosterSheet.ListObjects(1)
    Set ReportTable = ReportSheet.ListObjects(1)
    
    'Make sure the activity is in the RecordsSheet
    Set RecordsLabelRange = FindRecordsLabel(RecordsSheet, LabelCell)
        If RecordsLabelRange Is Nothing Then 'This shouldn't happen
            GoTo Footer
        End If
    
    'If the activty has no attendance information, don't tabulate
    Set AttendanceRange = FindRecordsAttendance(RecordsSheet, , LabelCell)
        If Not IsChecked(AttendanceRange) Then
            GoTo Footer
        End If
    
    'See if the activity is already on the ReportSheet
    Set ReportLabelRange = FindReportLabel(ReportSheet, LabelCell.Value)
    
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
    LabelString = LabelCell.Value
    PasteCell.EntireRow.ClearContents
    PasteCell.Value = LabelString
    
    'Remake the table if adding to the bottom of it
    If Intersect(PasteCell, ReportTable.ListColumns("Label").DataBodyRange) Is Nothing Then
        Set ReportTable = MakeReportTable
        
        Call TableFormatReport(ReportSheet, ReportTable)
    End If
    
    'Find students to tabulate on the Roster sheet
    Set TabulateRange = FindTabulateRange(RosterSheet, RecordsSheet, RecordsLabelRange)
        If TabulateRange Is Nothing Then
            GoTo RemakeTable
        End If
    
    'Pass the demographics for tabulation and add total
    Set TotalRange = ReportTable.HeaderRowRange.Find("Total", , xlValues, xlWhole)
    
    ValueArray = TabulateHelper(ReportSheet, RosterSheet, PasteCell.Value, TabulateRange)
    Call CopyToReport(ReportSheet, LabelString, ValueArray)
    ReportSheet.Cells(PasteCell.Row, TotalRange.Column) = TabulateRange.Cells.Count
    
RemakeTable:
    'Expand the table, format dates, add Marlett boxes
    Set ReportTable = MakeReportTable
    
    Call AddMarlettBox(ReportTable.ListColumns("Select").DataBodyRange)
    ReportTable.ListColumns("Date").DataBodyRange.NumberFormat = "mm/dd/yyyy"

Footer:

End Sub

Sub TabulateAllActivities()
'Loop through all activities on the Records Page and tabulate
'Only tabulate the totals if there are none

    Dim RecordsSheet As Worksheet
    Dim RecordsLabelRange As Range
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

Function TabulateDemo(SearchRange As Range, SearchType As String) As Variant
'This is an intermediate function that will call either a function for Windows or for MacOS
'MacOS doesn't support dictionaries, but that method is faster
'Returns an array with the count of each category being tabulated

    If Application.OperatingSystem Like "*Mac*" Then
        TabulateDemo = TabulateDemoMac(SearchRange, SearchType)
    Else
        TabulateDemo = TabulateDemoWin(SearchRange, SearchType)
    End If

Footer:

End Function

Function TabulateDemoMac(SearchRange As Range, SearchType As String) As Variant
'Returns an array with the values in the passed range tabulated
'MacOS cannot use dictionaries
'SearchType is for renaming "Other" values. This is why this is done piecemeal instead of all at once
    '(1, i) - header
    '(2, i) - value
    
    Dim HeaderRange As Range
    Dim c As Range
    Dim i As Long
    Dim j As Long
    Dim ListName As String
    Dim ValueString As String
    Dim SearchArray As Variant
    Dim CountArray As Variant
    
    'Read values into an array. This should always be 1 dimensional since only one column is passed
    'Need to loop for any non-contiguous range
    ReDim SearchArray(1 To SearchRange.Cells.Count)
    
    i = 1
    For Each c In SearchRange
        'Change the values for low income and first generation from "yes" to the search term
        If c.Value = "Yes" Then
            ValueString = SearchTerm
        Else
            ValueString = c.Value
        End If
    
        SearchArray(i) = ValueString
        i = i + 1
    Next c

    'Define the list of values to pull from. Get rid of spaces in the SearchType
    ListName = Replace(SearchType, " ", "") & "List"
    Set HeaderRange = Range(ListName)
    
    'Create the array for counting the values
    ReDim CountArray(1 To 2, 1 To HeaderRange.Cells.Count)
    
    i = 1
    For Each c In HeaderRange
        CountArray(1, i) = c.Value
        
        'Rename the "Other" category, if present
        If c.Value = "Other" Then
            CountArray(1, i) = "Other " & SearchType
        End If
        
        i = i + 1
    Next c
    
    'Count elements
    Select Case SearchType
        'Tabulating credits is done differently since we're putting integers into buckets
        Case "Credits"
            For i = LBound(SearchArray) To UBound(SearchArray)
                j = SearchArray(i)
                
                If IsEmpty(j) Or j = 0 Then 'VBA will return true for <45 on empty cells
                    CountArray(2, 4) = CountArray(2, 4) + 1
                ElseIf Not (IsNumeric(j)) Then
                    CountArray(2, 4) = CountArray(2, 4) + 1
                ElseIf j < 45 Then
                    CountArray(2, 1) = CountArray(2, 1) + 1
                ElseIf j <= 90 Then
                    CountArray(2, 2) = CountArray(2, 2) + 1
                ElseIf j > 90 Then
                    CountArray(2, 3) = CountArray(2, 3) + 1
                Else
                    CountArray(2, 4) = CountArray(2, 4) + 1 'To catch anything else
                End If
            Next i
        
        'Tabulating for everything else
        Case Else
            For i = 1 To UBound(SearchArray)
                ValueString = SearchArray(i)
                
                For j = 1 To UBound(CountArray, 2)
                    If ValueString = CountArray(1, j) Then
                        CountArray(2, j) = CountArray(2, j) + 1
                    End If
                Next j
                
                'If not found, put in the "Other" category
                If InStr(1, CountArray(1, j - 1), "Other") > 0 Then
                    CountArray(2, j - 1) = CountArray(2, j - 1) + 1
                End If
            Next i
NextValue:
    End Select

    TabulateDemoMac = CountArray

Footer:

End Function

Function TabulateDemoWin(SearchRange As Range, SearchType As String) As Variant
'Returns an array with the values in the passed range tabulated
'Uses a dictionary so no reference to the columns needs to be done
'SearchType is for renaming "Other" values. This is why this is done piecemeal instead of all at once
    '(1, i) - header
    '(2, i) - value

    Dim c As Range
    Dim OtherCount As Long
    Dim i As Long
    Dim j As Long
    Dim RenameString As String
    Dim TypeArray As Variant
    Dim DemoElement As Variant
    Dim SearchArray As Variant
    Dim CountArray As Variant
    Dim DemoDict As Object
    'Dim DemoDict As Scripting.Dictionary

    Set DemoDict = CreateObject("Scripting.Dictionary")
    'Set DemoDict = New Scripting.Dictionary
    
    'Ignore case
    DemoDict.CompareMode = vbTextCompare
    
    'Read values into an array. This should always be 1 dimensional since only one column is passed
    'Need to loop for any non-contiguous range
    ReDim SearchArray(1 To SearchRange.Cells.Count)
    
    i = 1
    For Each c In SearchRange
        SearchArray(i) = Trim(c.Value)
        i = i + 1
    Next c
    
    'Tabulating credits is done differently since we're putting integers into buckets
    If SearchType = "Credits" Then
        ReDim CountArray(1 To 2, 1 To 4) 'Should probably make this programmatic
            CountArray(1, 1) = "<45"
            CountArray(1, 2) = "45-90"
            CountArray(1, 3) = ">90"
            CountArray(1, 4) = "Other Credits"
            
        For i = 1 To UBound(SearchArray)
            If SearchArray(i) = "" Then
                CountArray(2, 4) = CountArray(2, 4) + 1
            Else
                j = SearchArray(i)
            End If
            
            If IsEmpty(j) Or j = 0 Then 'VBA will return true for <45 on empty cells
                CountArray(2, 4) = CountArray(2, 4) + 1
            ElseIf Not (IsNumeric(j)) Then
                CountArray(2, 4) = CountArray(2, 4) + 1
            ElseIf j < 45 Then
                CountArray(2, 1) = CountArray(2, 1) + 1
            ElseIf j <= 90 Then
                CountArray(2, 2) = CountArray(2, 2) + 1
            ElseIf j > 90 Then
                CountArray(2, 3) = CountArray(2, 3) + 1
            Else
                CountArray(2, 4) = CountArray(2, 4) + 1 'To catch anything else
            End If
        Next i

        GoTo ReturnArray
    End If

    'Read into the dictionary
    OtherCount = 0
    
    For i = 1 To UBound(SearchArray)
        DemoElement = SearchArray(i)
        
        'Rename the "yes" response for Low Income and First Generation
        If SearchType = "First Generation" Or SearchType = "Low Income" Then
            If DemoElement = "Yes" Then
                DemoElement = SearchType
            Else
                GoTo NextElement
            End If
        End If
        
        'Blanks go to "Other"
        If Not Len(DemoElement) > 0 Then
            OtherCount = OtherCount + 1
            
            GoTo NextElement
        End If
        
        'Keys and count
        If Not DemoDict.Exists(DemoElement) Then
            DemoDict.Add DemoElement, 1
        Else
            DemoDict(DemoElement) = DemoDict(DemoElement) + 1
        End If
        
NextElement:
    Next i

    'First Generation and Low Income don't have an other category
    If SearchType = "First Generation" Or SearchType = "Low Income" Then
        GoTo SkipOther
    End If

    'Rename the "Other" key. Insert one if it doesn't exist
    If Not DemoDict.Exists("Other") Then
        DemoDict.Add "Other", 0
    End If
    
    'Add the count of blanks
    DemoDict("Other") = DemoDict("Other") + OtherCount
    
    RenameString = "Other " & SearchType
    DemoDict.Key("Other") = RenameString
    
SkipOther:
    'First Generation and Low Income don't have an other category
    If SearchType = "First Generation" Or SearchType = "Low Income" Then
        ReDim CountArray(1 To 2, 1)
        
        CountArray(2, 1) = 0 'Initialize at zero
    Else
        ReDim CountArray(1 To 2, 1 To DemoDict.Count)
    End If
    
    i = 0
    For Each DemoElement In DemoDict.Keys
        i = i + 1
        
        CountArray(1, i) = DemoElement
        CountArray(2, i) = DemoDict(DemoElement)
SkipElement:
    Next DemoElement

ReturnArray:
    TabulateDemoWin = CountArray

Footer:

End Function

Function TabulateHelper(ReportSheet As Worksheet, RosterSheet As Worksheet, LabelString As String, Optional NameRange As Range) As Variant
'Returns an array with the value for each header in the Report table
'Passing NameRange limits tabulation to only those names. This should be a range on the RosterSheet
    '(1, i) - header
    '(2, i) - value

    Dim RefSheet As Worksheet
    Dim CoverSheet As Worksheet
    Dim RecordsSheet As Worksheet
    Dim RosterHeadersRange As Range
    Dim ReportHeadersRange As Range
    Dim TabulateHeadersRange As Range
    Dim SearchRange As Range
    Dim PasteCell As Range
    Dim c As Range
    Dim d As Range
    Dim i As Long
    Dim HeaderString As String
    Dim ReportTable As ListObject
    Dim RosterTable As ListObject
    Dim TabulateArray() As Variant
    Dim RecordsArray() As Variant
    Dim CoverArray() As Variant
    Dim ReturnArray() As Variant
    
    Set RefSheet = Worksheets("Ref Tables")
    Set CoverSheet = Worksheets("Cover Page")
    Set RecordsSheet = Worksheets("Records Page")
    Set RosterTable = RosterSheet.ListObjects(1)
    Set ReportTable = ReportSheet.ListObjects(1)
    
    'Search for the label in the ReportTable. If it's not found, grab one cell below the table
    Set PasteCell = FindReportLabel(ReportSheet, LabelString)
        If PasteCell Is Nothing Then
            Set PasteCell = FindLastRow(ReportSheet)
        End If
    
    'Create an array and read all of the headers into it
    Set ReportHeadersRange = ReportTable.HeaderRowRange
    
    ReDim ReturnArray(1 To 2, 1 To ReportHeadersRange.Cells.Count)
    
    i = 1
    For Each c In ReportHeadersRange
        ReturnArray(1, i) = c.Value
    
        i = i + 1
    Next c
    
    'For totals, grab info from the CoverSheet
    If PasteCell.Value = "Total" Then
        Set c = Range("ReportTotalsRowList")
    
        CoverArray = GetSubmissionInfo(ReportSheet, "Yes")

        ReDim RecordsArray(1 To 2, 1 To c.Cells.Count)
        i = 1
        For Each d In c
            RecordsArray(1, i) = d.Offset(0, -1).Value
            RecordsArray(2, i) = d.Value
        
            i = i + 1
        Next d
    Else
        'Grab activity headers from the RecordsSheet
        Set c = FindRecordsLabel(RecordsSheet, PasteCell)
        
        CoverArray = GetSubmissionInfo(ReportSheet)
        RecordsArray = GetActivityInfo(RecordsSheet, ReportSheet, c)
            If Not IsArray(CoverArray) Or IsEmpty(CoverArray) Then
                GoTo Footer
            ElseIf Not IsArray(RecordsArray) Or IsEmpty(CoverArray) Then
                GoTo Footer
            End If
    End If

    'Put all three together
    ReturnArray = ArrayAppend(RecordsArray, CoverArray)
        If Not IsArray(ReturnArray) Or IsEmpty(ReturnArray) Then
            GoTo Footer
        End If

    'Tabulate each tabulatable category. It's everything after "Last" and before "Notes" for Transfer Prep and MESA U, "Last" and "School" For College Prep
    '****Make this programmatic
    If IsCollege = True Then
        HeaderString = "School"
    Else
        HeaderString = "Notes"
    End If
    
    Set RosterHeadersRange = RosterTable.HeaderRowRange
    Set c = RosterHeadersRange.Find("Last", , xlValues, xlWhole).Offset(0, 1)
    Set d = RosterHeadersRange.Find(HeaderString, , xlValues, xlWhole).Offset(0, -1)
        If c Is Nothing Or d Is Nothing Then
            GoTo Footer
        End If
        
    Set TabulateHeadersRange = RosterSheet.Range(c, d)
    
    'Pass each term to TabulateHeaderRange
    For Each c In TabulateHeadersRange
        HeaderString = c.Value
            If Not Len(HeaderString) > 0 Then
                GoTo NextHeader
            End If
        
        'If a range of names was passed
        If NameRange Is Nothing Then
            Set SearchRange = RosterTable.ListColumns(HeaderString).DataBodyRange
        Else
            Set d = RosterTable.ListColumns(HeaderString).DataBodyRange
            Set SearchRange = Intersect(NameRange.EntireRow, d)
        End If
            
        If SearchRange Is Nothing Then
            GoTo NextHeader
        End If
        
        'Pass to tabulate
        Erase TabulateArray
        TabulateArray = TabulateDemo(SearchRange, HeaderString)
            If Not IsArray(TabulateArray) Or IsEmpty(TabulateArray) Then
                GoTo NextHeader
            End If
          
        'Put values in the RetunArray
        ReturnArray = ArrayAppend(ReturnArray, TabulateArray)

NextHeader:
    Next c

    'Return
    TabulateHelper = ReturnArray
    
Footer:
 
End Function

Sub TabulateListedActivities()
'Loop through all activities visible on the Report Page and tabulate
'Only tabulate the totals if there are none

    Dim ReportSheet As Worksheet
    Dim ReportLabelRange As Range
    Dim c As Range
    
    Set ReportSheet = Worksheets("Report Page")
    
    'Tabulate the totals row
    Call TabulateReportTotals
    
    'Define the range of labels on the Report and Records sheets
    Set ReportLabelRange = FindReportLabel(ReportSheet)
    
    'If there are no saved activities
    If ReportLabelRange.Cells.Count = 1 Then
        If ReportLabelRange.Value = "Total" Then
            GoTo Footer
        End If
    End If
    
    'Loop through labels
    For Each c In ReportLabelRange
        'These get turned back on somewhere in the TabulateActivity sub
        Application.EnableEvents = False
        Application.ScreenUpdating = False
        Application.DisplayAlerts = False

        Call TabulateActivity(c)
    Next c

Footer:

End Sub

Sub TabulateReportActivities()
'Retabulates the activities on the ReportSheet
'Does not tabulate anything that isn't already present
'Removes activities no longer on the RecordsSheet

    Dim ReportSheet As Worksheet
    Dim RecordsSheet As Worksheet
    Dim ReportLabelRange As Range
    Dim RecordsLabelRange As Range
    Dim DelRange As Range
    Dim c As Range
    Dim ReportTable As ListObject

    Set ReportSheet = Worksheets("Report Page")
    Set RecordsSheet = Worksheets("Records Page")
    
    'Make sure there's a table with at least one activity
    If CheckReport(ReportSheet) > 2 Then
        GoTo Footer
    End If
    
    'Define area to search
    Set ReportTable = ReportSheet.ListObjects(1)
    Set ReportLabelRange = FindReportLabel(ReportSheet)
    Set RecordsLabelRange = FindRecordsLabel(RecordsSheet)
    
    'Any activities on the ReportSheet but not the RecordsSheet are removed
    Call UnprotectSheet(ReportSheet)
    
    For Each c In ReportLabelRange
        If RecordsLabelRange.Find(c.Value, , xlValues, xlWhole) Is Nothing Then
            Set DelRange = BuildRange(c, DelRange)
        Else
            Call TabulateActivity(c)
        End If
    Next c
    
    Set ReportTable = ReportSheet.ListObjects(1)
    
    If Not DelRange Is Nothing Then
        Call RemoveRows(ReportSheet, ReportTable.DataBodyRange, ReportLabelRange, DelRange)
    End If

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
    Dim d As Range
    Dim i As Long
    Dim RosterTable As ListObject
    Dim ReportTable As ListObject
    Dim PasteCell As Range
    Dim TempRange As Range
    Dim PassArray As Variant
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
    ElseIf CheckReport(ReportSheet) > 3 Then
        Call MakeReportTable
    End If
    
    Set RosterTable = RosterSheet.ListObjects(1)
    Set ReportTable = ReportSheet.ListObjects(1)
    Set ReportHeaderRange = ReportTable.HeaderRowRange
    
    Call UnprotectSheet(ReportSheet)
    
    'Clear the contents
    Call ReportClearTotals

    'Add race, gender. Grade for College Prep, Credits and Major for Transfer Prep and MESA U
    PassArray = TabulateHelper(ReportSheet, RosterSheet, "Total")
        If IsEmpty(PassArray) Or Not IsArray(PassArray) Then
            GoTo Footer
        End If
    
    Call CopyToReport(ReportSheet, "Total", PassArray)
    
    'Fill in the totals cell
    Set c = ReportHeaderRange.Find("Total", , xlValues, xlWhole).Offset(1, 0)
    c.Value = RosterTable.ListRows.Count
    
    'Apply bold font
    ReportHeaderRange.Offset(1, 0).Font.Bold = True
    
    'Make sure we have Marlett boxes
    If ReportTable.Range.Rows.Count > 2 Then
        Set c = ReportTable.ListColumns("Select").DataBodyRange
        Set d = c.Resize(c.Rows.Count - 1, 1).Offset(1, 0)
        Call AddMarlettBox(d)
    End If
    
Footer:

End Sub





