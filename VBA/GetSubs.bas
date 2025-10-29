Attribute VB_Name = "GetSubs"
Option Explicit

Function GetActivityInfo(RecordsSheet As Worksheet, ReportSheet As Worksheet, LabelCell As Range) As Variant
'Creates an array to put in the Report Page for the passed label
'Returns empty if the search range can't be found
'Values that can't be found are returned as ""
    '(1, i) - header
    '(2, i) - value

    Dim RecordsLabelRange As Range
    Dim RecordsHeaderRange As Range
    Dim c As Range
    Dim d As Range
    Dim i As Long
    Dim InfoArray() As Variant
        
    'Define where to search
    Set RecordsLabelRange = FindRecordsLabel(RecordsSheet, LabelCell)
    Set RecordsHeaderRange = FindRecordsActivityHeaders(RecordsSheet)
        If RecordsLabelRange Is Nothing Then
            GoTo Footer
        ElseIf RecordsHeaderRange Is Nothing Then
            GoTo Footer
        End If
    
    'Read into an array
    ReDim InfoArray(1 To 2, 1 To RecordsHeaderRange.Cells.Count)
    
    i = 1
    For Each c In RecordsHeaderRange
        Set d = RecordsSheet.Cells(c.Row, RecordsLabelRange.Column)
        
        If Not d Is Nothing Then
            InfoArray(1, i) = c.Value
            InfoArray(2, i) = d.Value

            i = i + 1
        End If
    Next c
        
    GetActivityInfo = InfoArray
    
Footer:
   
End Function

Function GetEdition() As String

    Dim ChangeSheet As Worksheet
    
    'Dim FileName As String
    'Dim TempName As String

    'FileName = ThisWorkbook.Name
    'TempName = Trim(RemoveNonNumeric(FileName))
    'GetEdition = TempName
    
    Set ChangeSheet = Worksheets("Change Log")
    
    GetEdition = ChangeSheet.Range("A1").Value
    
End Function

Function GetCoverInfo() As Variant
'Grabs the name, date, center, program, and version from the CoverSheet
'Returns a 2D array with each value
'Returns nothing on error

    Dim CoverSheet As Worksheet
    Dim RefRange As Range
    Dim ValueRange As Range
    Dim c As Range
    Dim d As Range
    Dim ValueString As String
    Dim i As Long
    Dim ReturnArray() As Variant
    
    Set CoverSheet = Worksheets("Cover Page")
    Set RefRange = Range("CoverInfoList")
        If RefRange Is Nothing Then
            GoTo Footer
        End If
    
    'Grab the items we want
    ReDim ReturnArray(1 To 2, 1 To RefRange.Cells.Count)
    
    'Define the search area for the values we want
    Set c = CoverSheet.Range("A:A").Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
        If c Is Nothing Then
            GoTo Footer
        End If
        
    Set ValueRange = CoverSheet.Range("A1")
        
    i = 1
    For Each c In RefRange
        ReturnArray(1, i) = c.Value
        Set d = ValueRange.Offset(i - 1, 0) 'The reference list is in the same order, so we just go down the cells one by one
        
        Select Case c.Value
            
            Case "Report"
                ValueString = Replace(d.Value, "Weekly Report", " ") 'Make this programmatic at some point
                
            Case "Version"
                ValueString = GetEdition()
            
            Case Else
                ValueString = d.Offset(0, 1).Value
                
        End Select
        
        ReturnArray(2, i) = ValueString
NextItem:
        i = i + 1
    Next c
   
    If IsArray(ReturnArray) And Not IsEmpty(ReturnArray) Then
        GetCoverInfo = ReturnArray
    End If
   
Footer:
   
End Function

Function GetFormInfo(UForm As Object, Optional RowIndex As Long) As Variant
'Grabs the information for the selected activity on a form
'Used for creating, loading, or adding students to an activity, and for tabulating
'Creating a new activity has values in separate boxes. The other user forms store the information in a listbox
'Called when adding, loading, or adding students to an activity
    '(1, i) - Header
    '(2, i) - Value
    '(3, i) - Address (might be omitted)

    Dim RefSheet As Worksheet
    Dim ColIndexRange As Range
    Dim ControlRefRange As Range
    Dim HeaderRefRange As Range
    Dim AddressRefRange As Range
    Dim c As Range
    Dim d As Range
    Dim i As Long
    Dim j As Long
    Dim ColIndex As Long
    Dim CategoryString As String
    Dim FormName As String
    Dim CBoxName As String
    Dim TempValue As String
    Dim ControlNameTable As ListObject
    Dim CBox As Object
    Dim RefArray As Variant
    Dim ReturnArray As Variant
    
    'Define the table where the information of control names is kept
    Set RefSheet = Worksheets("Ref Tables")
    Set ControlNameTable = RefSheet.ListObjects("ControlNameTable")
    
    FormName = UForm.Name
    
    'Grab the headers and select the column based on the UserForm's name
    Set HeaderRefRange = ControlNameTable.ListColumns("Form Header").DataBodyRange
        If HeaderRefRange Is Nothing Then
            GoTo Footer
        End If
  
    Set ControlRefRange = Range(FormName & "List")
        If ControlRefRange Is Nothing Then
            GoTo Footer
        End If
        
    Set AddressRefRange = Range("ActivitySheetAddressList")
        If AddressRefRange Is Nothing Then
            GoTo Footer
        End If

    'Grab the colum index if it's not the NewActivityForm
    If UForm.Name <> "NewActivityForm" Then
        Set ColIndexRange = Range("ControlColIndexList")
    End If
    
    'Make two arrays, one containing the reference information to grab values from the form, one to return
    ReDim RefArray(1 To 3, 1 To HeaderRefRange.Cells.Count)
    ReDim ReturnArray(1 To 3, 1 To HeaderRefRange.Cells.Count)
    
    i = 1
    For Each c In HeaderRefRange
        'Headers
        RefArray(1, i) = c.Value
        ReturnArray(1, i) = c.Value
        
        'Address
        Set d = RefSheet.Cells(c.Row, AddressRefRange.Column)
            If d Is Nothing Then
                GoTo NextHeader
            End If
            
        ReturnArray(3, i) = d.Value
    
        'ControlName
        Set d = RefSheet.Cells(c.Row, ControlRefRange.Column)
            If d Is Nothing Then
                GoTo NextHeader
            End If
        
        RefArray(2, i) = d.Value
        
        'Column index, if applicable
        If ColIndexRange Is Nothing Then
            GoTo NextHeader
        End If
        
        Set d = RefSheet.Cells(c.Row, ColIndexRange.Column)
            If d Is Nothing Then
                GoTo NextHeader
            End If
            
        RefArray(3, i) = d.Value

NextHeader:
        i = i + 1
    Next c
    
    'Loop through and grab the values off the user form
    For i = 1 To UBound(RefArray, 2)
        CBoxName = RefArray(2, i)
            If Not Len(CBoxName) > 0 Then
                GoTo NextControl
            End If

        Set CBox = UForm.Controls(CBoxName)
            If CBox Is Nothing Then
                GoTo NextControl
            End If
        
        'We need the column index and passed row index for every sheet except the NewActivityForm
        If UForm.Name <> "NewActivityForm" Then
            ColIndex = RefArray(3, i)
                If Not Len(str(ColIndex)) > 0 Then
                    GoTo NextControl
                End If
            
            TempValue = CBox.List(RowIndex, ColIndex)
        Else
            TempValue = CBox.Value
        End If
        
        If Not Len(TempValue) > 0 Then
            GoTo NextControl
        End If
        
        ReturnArray(2, i) = TempValue
        
        'Grab the value of the Practice to find the category
        If RefArray(1, i) = "Practice" Then
            CategoryString = GetPracticeCategory(TempValue)
            
            If Not Len(CategoryString) > 0 Then
                GoTo NextControl
            End If

            'Loop through to find the Category header. It should always be one after
            For j = 1 To UBound(RefArray, 2)
                If RefArray(1, j) = "Category" Then
                    ReturnArray(2, j) = CategoryString
                End If
            Next j
        End If
        
NextControl:
    TempValue = ""
    Next i

    'Return
    GetFormInfo = ReturnArray
    
Footer:

End Function

Function GetPracticeCategory(PracticeString As String) As String
'Finds the category that a practice belongs to and returns it
'Returns nothing if not found

    
    Dim PracticeRefRange As Range
    Dim c As Range
    Dim CategoryString As String
    Dim ReturnArray As Variant
    
    If Not Len(PracticeString) > 0 Then
        GoTo Footer
    End If
    
    'Search for the practice
    Set PracticeRefRange = Range("ActivitiesList")
        If PracticeRefRange Is Nothing Then
            GoTo Footer
        End If
        
    Set c = PracticeRefRange.Find(PracticeString, , xlValues, xlWhole)
        If c Is Nothing Then
            GoTo Footer
        End If
        
    'Categories are one column to the left
    CategoryString = c.Offset(0, -1).Value
        If Not Len(CategoryString) > 0 Then
            GoTo Footer
        End If
    
    GetPracticeCategory = CategoryString
    
Footer:

End Function

Function GetReadyToExport(SheetArray As Variant) As Variant
'Cheks the passed sheets and returns a 2D array of which are ready
'Can check the Cover, Records, Report, and Roster sheets
    '1      Ready
    '0      No ready
    'Blank  Not checked
    
    Dim RosterSheet As Worksheet
    Dim RecordsSheet As Worksheet
    Dim ReportSheet As Worksheet
    Dim i As Long
    Dim SheetName As String
    Dim ReturnArray As Variant
    
    Set RosterSheet = Worksheets("Roster Page")
    Set RecordsSheet = Worksheets("Records Page")
    Set ReportSheet = Worksheets("Report Page")

    'Loop through the check the passed sheets
    ReDim ReturnArray(1 To 2, 1 To UBound(SheetArray))
    
    For i = 1 To UBound(SheetArray)
        SheetName = SheetArray(i)
        ReturnArray(1, i) = SheetName
        
        Select Case SheetName
        
            Case "Cover Page"
                If CheckCover() = 1 Then
                    ReturnArray(2, i) = 1
                Else
                    ReturnArray(2, i) = 0
                End If
            
            Case "Roster Page"
                If Not CheckTable(RosterSheet) > 2 Then
                    ReturnArray(2, i) = 1
                Else
                    ReturnArray(2, i) = 0
                End If
                
            Case "Records Page"
                If Not CheckRecords(RecordsSheet) > 2 Then
                    ReturnArray(2, i) = 1
                Else
                    ReturnArray(2, i) = 0
                End If
                
            Case "Report Page"
                If Not CheckReport(ReportSheet) > 3 Then 'Doesn't require activities
                    ReturnArray(2, i) = 1
                Else
                    ReturnArray(2, i) = 0
                End If
                
        End Select
    Next i

    If Not IsArray(ReturnArray) Or Not IsEmpty(ReturnArray) Then
        GetReadyToExport = ReturnArray
    End If

Footer:

End Function

Function GetReadyToSave(CoverSheet As Worksheet, ReportSheet As Worksheet, Optional RecordsSheet As Worksheet, Optional RosterSheet As Worksheet) As Variant
'Checks that everything is ready to save or export
'For each sheet, returns 1 if it's filled out, 0 if it's not

    Dim SearchRange As Range
    Dim c As Range
    Dim d As Range
    Dim SearchString As String
    Dim ResultString As String
    Dim i As Long
    Dim j As Long
    Dim TempArray() As Variant
    Dim SearchArray() As Variant
    
    'Put in the two sheets that will always be use
    ReDim TempArray(1 To 2, 1 To 2)
        TempArray(1, 1) = "Cover Page"
        TempArray(1, 2) = "Report Page"
        TempArray(2, 1) = 0
        TempArray(2, 2) = 0
        
    'Redim the array depending on the number of arguments passed
    i = 2
    
    'Roster Sheet
    If RosterSheet Is Nothing Then
        GoTo CheckRecords
    End If
    
    i = i + 1
    ReDim Preserve TempArray(1 To 2, 1 To i)
        TempArray(1, i) = "Roster Page"
        TempArray(2, i) = 0
    
    j = CheckTable(RosterSheet)
    If Not j > 2 Then 'At least one student on the roster table
        TempArray(2, i) = 1
    End If
    
CheckRecords:
    'Records Sheet
    If RecordsSheet Is Nothing Then
        GoTo CheckCover
    End If

    i = i + 1
    ReDim Preserve TempArray(1 To 2, 1 To i)
        TempArray(1, i) = "Records Page"
        TempArray(2, i) = 0
        
    j = CheckRecords(RecordsSheet)
    If Not j > 1 Then 'Both students and activities
        TempArray(2, i) = 1
    End If

CheckCover:
    'Cover sheet
    ReDim SearchArray(1 To 3)
        SearchArray(1) = "Name"
        SearchArray(2) = "Date"
        SearchArray(3) = "Center"

    For j = 1 To UBound(SearchArray)
        SearchString = SearchArray(j)
        Set c = CoverSheet.Range("A:A").Find(SearchString, , xlValues, xlWhole).Offset(0, 1)
        
        If Len(c.Value) < 1 Then
            GoTo CheckReportTable
        End If
    Next j
    
    TempArray(2, 1) = 1
    
CheckReportTable:
    'Report Sheet
    j = CheckReport(ReportSheet)
    If Not j > 2 Then
        TempArray(2, 2) = 1
    End If
    
    'Return
    GetReadyToSave = TempArray

Footer:

End Function

Function GetRecordsActivityHeaders(RecordsSheet As Worksheet, Optional LabelString As String, Optional OperationString As String)
'Returns a 1D array with the headers on the RecordsSheet
'If a label is passed, returns a 2D array
    '(1, i) - header
    '(2, i) - value
'Passing "All" creates an i x j array, where i is the number of activities and j is the number of headers
    '(1, i) - header
    '(1 + j, i) - value
        
    '(1, i) - Headers
    '(2, i) - Values for first activity
    '(3, i) - Values for second activity, etc.

    Dim LabelRange As Range
    Dim HeaderRange As Range
    Dim ValueRange As Range
    Dim c As Range
    Dim i As Long
    Dim j As Long
    Dim HeaderArray As Variant
    Dim ValueArray As Variant
    Dim ReturnArray As Variant
    
    Set HeaderRange = FindRecordsActivityHeaders(RecordsSheet)
        If HeaderRange Is Nothing Then
            GoTo Footer
        End If
    
    j = HeaderRange.Cells.Count
    
    'Grab the headers
    ReDim HeaderArray(1 To j)
    
    i = 1
    For Each c In HeaderRange
        HeaderArray(i) = c.Value
        
        i = i + 1
    Next c

    'If no arguments were passed
    If Not Len(LabelString) > 0 And Not Len(OperationString) > 0 Then
        ReturnArray = HeaderArray
        
        GoTo ReturnArray
    End If
    
    'If a label or operation was passed
    Set c = FindRecordsLabel(RecordsSheet)
        If c Is Nothing Then
            GoTo Footer
        End If

    'Looking for a single label
    If Len(LabelString) > 0 Then
        Set LabelRange = c.Find(LabelString, , xlValues, xlWhole)
            If LabelRange Is Nothing Then
                GoTo Footer
            End If
    'Looking for all labels
    ElseIf OperationString = "All" Then
        Set LabelRange = c
    Else
        GoTo Footer
    End If
    
    'Make the array of values
    i = LabelRange.Cells.Count
    
    ReDim ValueArray(1 To i, 1 To j)
    
    i = 1
    For Each c In LabelRange
        For j = 1 To HeaderRange.Cells.Count
            ValueArray(i, j) = c.Offset(j - 1, 0).Value
        Next j

        i = i + 1
    Next c

    'Put the two together
    ReDim ReturnArray(1 To LabelRange.Cells.Count + 1, 1 To HeaderRange.Cells.Count)
    
    For j = 1 To HeaderRange.Cells.Count
        ReturnArray(1, j) = HeaderArray(j)
    Next j
    
    For i = 1 To LabelRange.Cells.Count
        For j = 1 To HeaderRange.Cells.Count
            ReturnArray(1 + i, j) = ValueArray(i, j)
        Next j
    Next i


ReturnArray:
    GetRecordsActivityHeaders = ReturnArray

Footer:

End Function

Function GetSubmissionInfo(ReportSheet As Worksheet, Optional PullDate As String) As Variant
'Grabs the name, center, and date from the Cover Page
'Will grab the date from the CoverSheet if "Yes" is passed for PullDate

    Dim CoverSheet As Worksheet
    Dim ReportSearchRange As Range
    Dim c As Range
    Dim d As Range
    Dim i As Long
    Dim SearchString As String
    Dim FoundString As String
    Dim TempArray() As Variant
    
    Set CoverSheet = Worksheets("Cover Page")
    
    'Define what part of the headers we want. It'll be between the "Select" and "Date" headers unless "Yes" was passed
    If PullDate = "Yes" Then
        Set c = FindTableHeader(ReportSheet, "Select", "Label")
    Else
        Set c = FindTableHeader(ReportSheet, "Select", "Date")
    End If
    
    If c Is Nothing Then
        GoTo Footer
    End If

    Set ReportSearchRange = c.Resize(1, c.Columns.Count - 2).Offset(0, 1)
    If ReportSearchRange Is Nothing Then
        GoTo Footer
    End If

    'Search on the CoverSheet and find corresponding values
    ReDim TempArray(1 To 2, 1 To ReportSearchRange.Cells.Count)
    
    i = 1
    For Each c In ReportSearchRange
        SearchString = c.Value
        TempArray(1, i) = SearchString
        
        Set d = CoverSheet.Range("A:A").Find(SearchString, , xlValues, xlWhole)
        
        If Not d Is Nothing Then
            FoundString = d.Offset(0, 1).Value
            TempArray(2, i) = FoundString
        End If
        
        i = i + 1
    Next c

    'Return array
    GetSubmissionInfo = TempArray
    
Footer:

End Function

