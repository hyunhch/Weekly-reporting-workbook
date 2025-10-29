Attribute VB_Name = "MakeSubs"
Option Explicit

Function MakeActivityTable(ActivitySheet As Worksheet) As ListObject
'Puts in headers, creates a table

    Dim RosterSheet As Worksheet
    Dim ActivityTableStart As Range
    Dim HeaderRange As Range
    Dim c As Range
    Dim i As Long
    Dim HeaderArray() As Variant
    Dim RosterTable As ListObject
    Dim ActivityTable As ListObject
    
    Set RosterSheet = Worksheets("Roster Page")
    Set ActivityTableStart = ActivitySheet.Range("A6")
    Set RosterTable = RosterSheet.ListObjects(1)
    
    'If there's already a table, unlist
    Call RemoveTable(ActivitySheet)
    
    'Read in the headers of the Roster Table and pass to copy headers, then make a table
    'Not using the reference sheet so custom columns are pulled in as well
    ReDim HeaderArray(1 To RosterTable.ListColumns.Count)
    i = 1

    For Each c In RosterTable.HeaderRowRange
        HeaderArray(i) = c.Value
        i = i + 1
    Next c
    
    Call TableResetHeaders(ActivitySheet, ActivityTableStart, HeaderArray)
    
    'Create the table, format, add Marlett Boxes
    Set ActivityTable = MakeTable(ActivitySheet)
    
    If ActivityTable.ListRows.Count > 0 Then
        Call AddMarlettBox(ActivityTable.ListColumns("Select").DataBodyRange)
    End If
    
    Set MakeActivityTable = ActivityTable
    
Footer:

End Function

Function MakeButton(TargetSheet As Worksheet, TargetArray As Variant) As Long
'Makes a button on the passed sheet, an array contains the arguments
'Returns 1 on sucess
    '(1) - Range
    '(2) - OnAction
    '(3) - Caption
    '(4) - Name
    
    Dim TargetRange As Range
    Dim RangeString As String
    Dim OnActionString As String
    Dim CaptionString As String
    Dim NameString As String
    Dim TargetButton As Button
    
    RangeString = TargetArray(1)
    OnActionString = TargetArray(2)
    CaptionString = TargetArray(3)
    NameString = TargetArray(4)
    
    Set TargetRange = TargetSheet.Range(RangeString)
    Set TargetButton = TargetSheet.Buttons.Add(TargetRange.Left, TargetRange.Top, _
        TargetRange.Width, TargetRange.Height)

    With TargetButton
        .OnAction = OnActionString
        .Caption = CaptionString
        .Name = NameString
    End With

    MakeButton = 1

Footer:

End Function

Function MakeReportTable() As ListObject

    Dim ReportSheet As Worksheet
    Dim ReportTableRange As Range
    Dim ReportTableStart As Range
    Dim TempRange As Range
    Dim c As Range
    Dim i As Long
    Dim ReportTable As ListObject
    Dim TempArray() As Variant
    Dim HeadersArray() As Variant
    Dim TotalsArray() As Variant

    Set ReportSheet = Worksheets("Report Page")
    Set ReportTableStart = ReportSheet.Range("A:A").Find("Select", , xlValues, xlWhole)
        If ReportTableStart Is Nothing Then 'If the table headers got messed up
            Set ReportTableStart = ReportSheet.Range("A6")
        End If
       
    Call UnprotectSheet(ReportSheet)

    'Remove any existing filters, unlist the table and remove formatting
    If ReportSheet.AutoFilterMode = True Then
        ReportSheet.AutoFilterMode = False
    End If

    Call RemoveTable(ReportSheet)

    'Reset headers. This creates two rows
    HeadersArray = Application.Transpose(ActiveWorkbook.Names("ReportHeadersList").RefersToRange.Value)
    Call TableResetHeaders(ReportSheet, ReportTableStart, HeadersArray)
    Call TableResetReportTotalHeaders(ReportSheet, ReportTableStart)
    
    'Define table range and clear formats
    Set ReportTableRange = FindTableRange(ReportSheet)
        ReportSheet.Cells.ClearFormats
        
    'Make a new table and format
    Set ReportTable = ReportSheet.ListObjects.Add(SourceType:=xlSrcRange, Source:=ReportTableRange, _
        xlListObjectHasHeaders:=xlYes)
        ReportTable.Name = "ReportTable"
        ReportTable.ShowTableStyleRowStripes = False
        
    Call TableFormatReport(ReportSheet, ReportTable)
    
    'Marlett boxes except in the Totals (2nd) row
    If ReportTable.Range.Rows.Count > 2 Then
        Call AddMarlettBox(ReportTable.ListColumns("Select").DataBodyRange)
    End If
    
    ReportTable.ListColumns("Select").DataBodyRange(1, 1).Font.Name = "Aptos Narrow" 'This can be anything except Marlett to prevent the cell from being
    
    Set MakeReportTable = ReportTable
    
Footer:

End Function

Function MakeRosterTable(RosterSheet As Worksheet) As ListObject
'Called when parsing the roster
'Returns the RosterTable if successful
'Returns nothing on error

    Dim RosterTableRange As Range
    Dim RosterNameRange As Range
    Dim DelRange As Range
    Dim c As Range
    Dim i As Long
    Dim RosterTable As ListObject
    Dim HeaderArray As Variant
    
    Set c = RosterSheet.Range("A6")
    
    'Remove the table, if there is one
    If CheckTable(RosterSheet) < 4 Then
        RosterSheet.AutoFilterMode = False
        Call RemoveTable(RosterSheet)
    End If
    
    'Reset the headers
    HeaderArray = Application.Transpose(ActiveWorkbook.Names("RosterHeadersList").RefersToRange.Value)
    Call TableResetHeaders(RosterSheet, c, HeaderArray) 'This will not remove additional columns added to the right of the default ones
    RosterSheet.Cells.ClearFormats
    
    'Find the range for the new table, break if there is nothing but the header
    Set RosterTableRange = FindTableRange(RosterSheet)
        If Not RosterTableRange.Rows.Count > 1 Then
            GoTo Footer
        End If
    
    'Make new table
    Set RosterTable = MakeTable(RosterSheet, "RosterTable", RosterTableRange)
        If RosterTable.DataBodyRange Is Nothing Then
            GoTo Footer
        End If
    
    Call TableFormat(RosterSheet, RosterTable)
    Call AddMarlettBox(RosterTable.ListColumns("Select").DataBodyRange)
    
    Set MakeRosterTable = RosterTable
    
Footer:
    
End Function

Function MakeTable(TargetSheet As Worksheet, Optional NewTableName As String, Optional TargetRange As Range) As ListObject
'Generic function to make a table. Looks for "Select" in the first column

    Dim NewTableRange As Range
    Dim TableName As String
    Dim NewTable As ListObject

    Call UnprotectSheet(TargetSheet)
    
    'Find the range to use
    If TargetRange Is Nothing Then
        Set NewTableRange = FindTableRange(TargetSheet)
    Else
        Set NewTableRange = TargetRange
    End If
    
    If NewTableRange Is Nothing Then
        MsgBox ("There was a problem creating a table on sheet " & TargetSheet.Name)
        GoTo Footer
    End If
        
    'Unlist any existing table, remove formats
    Call RemoveTable(TargetSheet)
    NewTableRange.ClearFormats

    'Create new table
    Set NewTable = TargetSheet.ListObjects.Add(SourceType:=xlSrcRange, Source:=NewTableRange, _
        xlListObjectHasHeaders:=xlYes)
    
    NewTable.ShowTableStyleRowStripes = False

    'Assign a name if passed
    If Len(NewTableName) > 0 Then
        NewTable.Name = NewTableName
    End If

    'Removed getting rid of blank rows
    'Better to use a RemoveBlanks function that can be pointed to any column, rather than just the first name column of a table
    
    'Put in Marlett boxes. I had taken this out but can't remember why
    If NewTable.HeaderRowRange.Find("Select") Is Nothing Then
        GoTo ReturnTable
    ElseIf Not NewTable.DataBodyRange Is Nothing Then
        Call AddMarlettBox(NewTable.ListColumns("Select").DataBodyRange)
    End If
    
ReturnTable:
    Set MakeTable = NewTable

Footer:

End Function



