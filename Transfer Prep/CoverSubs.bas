Attribute VB_Name = "CoverSubs"
Option Explicit

Function NewSaveBook(PasteBook As Workbook, CopyCover As Worksheet, CopyRoster As Worksheet, CopyReport As Worksheet, CopyRecords As Worksheet, SheetNameArray As Variant, Optional SaveType As String) As Boolean
'Create a new workbook to save or upload, depending on what is passed to it

    Dim TableStart As Range
    Dim CoverRange As Range
    Dim ReportRange As Range
    Dim RecordsRange As Range
    Dim c As Range
    Dim i As Long
    
    NewSaveBook = False
    
    'Create the sheets we need
    If SaveType = "SharePoint" Then
        PasteBook.Sheets.Add().Name = "Report"
    Else
        For i = LBound(SheetNameArray) To UBound(SheetNameArray)
            PasteBook.Sheets.Add().Name = SheetNameArray(i)
        Next i
    End If
    
    PasteBook.Worksheets("Sheet1").Delete
    
    'Define the areas to copy
    Set TableStart = CopyReport.Range("A:A").Find("Select", , xlValues, xlWhole)
    Set ReportRange = FindTableRange(CopyReport, TableStart)

    If SaveType = "SharePoint" Then
        GoTo Copy
    Else
        Set CoverRange = CopyCover.Range("A1:B5")
        Set TableStart = CopyRecords.Range("A1")
        Set RecordsRange = FindTableRange(CopyRecords, TableStart)
    End If
    
Copy:
    'Copy information from the Cover sheet, Report sheet, and records sheet
    ReportRange.Copy
    PasteBook.Worksheets("Report").Range("A1").PasteSpecial
    PasteBook.Worksheets("Report").Range("A1").EntireColumn.Delete 'We don't need the "Select" column
    
    If SaveType = "SharePoint" Then
        GoTo FormatReport
    Else
        CoverRange.Copy
        PasteBook.Worksheets("Cover").Range("A1").PasteSpecial
        RecordsRange.Copy
        PasteBook.Worksheets("Attendance").Range("A1").PasteSpecial
    End If
    
    'Use the Roster to make a detailed attendance report
    Dim DetailedSheet As Worksheet
    Dim AttendanceSheet As Worksheet
    Dim RosterRange As Range
    Dim NewRecordsRange As Range
    Dim ActivityLabel As String
    
    Set RosterRange = CopyRoster.ListObjects("RosterTable").ListColumns("First").DataBodyRange
    Set NewRecordsRange = FindTableRange(PasteBook.Worksheets("Attendance"), PasteBook.Worksheets("Attendance").Range("A1"))
    Set DetailedSheet = PasteBook.Worksheets("Detailed Attendance")
    Set AttendanceSheet = PasteBook.Worksheets("Attendance")
    
    If RosterRange Is Nothing Then
        MsgBox ("Something has gone wrong with the Roster Page. Please parse the roster and try again.")
        GoTo Footer
    End If
    
    'Create headers, copying from the roster sheet directly
    Dim HeaderArray() As String
    
    CopyRoster.ListObjects("RosterTable").HeaderRowRange.Copy
    DetailedSheet.Range("D1").PasteSpecial xlPasteValues
    'Write over the first three columns of the roster header
    HeaderArray = Split("First;Last;Label;Practice;Date;Decription", ";")
    DetailedSheet.Range("A1").Resize(1, UBound(HeaderArray) + 1) = HeaderArray
    
    'Loop through and make a new row for each time a student was marked present
    Dim NameCell As Range
    Dim CopyArray() As Variant
    Dim FRow As Long
    Dim FCol As Long
    Dim LRow As Long
    Dim LCol As Long
    Dim j As Long
    Dim k As Long
    
    With AttendanceSheet
        FRow = .Range("A:A").Find("H BREAK", , xlValues, xlWhole).Row + 1
        FCol = .Range("1:1").Find("V BREAK", , xlValues, xlWhole).Column + 1
        LRow = .Range("A:A").Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
        LCol = .Range("1:1").Find("*", SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
        Set TableStart = .Range(Cells(FRow + 1, FCol + 1).Address, Cells(LRow, LCol).Address) 'Reusing this variable
        
        k = 2
        For i = FRow To LRow
            For j = FCol To LCol
                If .Cells(i, j).Value = "a" Then
                    'Copy over names
                    DetailedSheet.Range(Cells(k, 1).Address, Cells(k, 2).Address).Value = .Range(Cells(i, 1).Address, Cells(i, 2).Address).Value
                    'Copy over the activity information
                    CopyArray = WorksheetFunction.Transpose(.Range(Cells(1, j).Address, Cells(4, j).Address).Value)
                    DetailedSheet.Range(Cells(k, 3).Address, Cells(k, 3 + UBound(CopyArray)).Address).Value = CopyArray
                    
                    k = k + 1
                End If
            Next j
        Next i
      End With
      
    'Copy over the demographic information from the roster sheet
    For i = 2 To k - 1
        Set NameCell = NameMatch(DetailedSheet.Cells(i, 1), RosterRange)
        'Make sure there is a match. If there isn't, skip it
        If NameCell Is Nothing Then
            GoTo SkipName
        End If
        'Start from the Ethnicity column and go to the end of the table. Copy the student information
        Erase CopyArray
        CopyArray = WorksheetFunction.Transpose(CopyRoster.Range(Cells(NameCell.Row, 4).Address, Cells(NameCell.Row, CopyRoster.ListObjects("RosterTable").ListColumns.Count).Address).Value)
        DetailedSheet.Range(Cells(i, 7).Address, Cells(i, 6 + UBound(CopyArray)).Address).Value = Application.Transpose(CopyArray) 'Array is starting at 1 here
SkipName:
    Next i
  
    'Reformat the Records page to remove spacers and have "1" instead of "a"
    For Each c In NewRecordsRange
        If c.Value = "a" Then
            c.Value = "1"
        ElseIf c.Value = "V BREAK" Then
            c.Value = ""
        ElseIf c.Value = "H BREAK" Then
            c.EntireRow.Delete
        End If
    Next c
    
    'A little formatting to make things more readable
    AttendanceSheet.Range("A1:A4").Font.Bold = True
    AttendanceSheet.Range("A1:B1").EntireColumn.AutoFit
    Intersect(NewRecordsRange, AttendanceSheet.Range("1:1")).Font.Bold = True
    Intersect(NewRecordsRange, AttendanceSheet.Range("3:3")).NumberFormat = "mm-dd-yyyy"
    
    'Format the Detailed Attendance sheet
    DetailedSheet.Range("A1", Cells(1, 6 + CopyRoster.ListObjects("RosterTable").ListColumns.Count).Address).Font.Bold = True
    DetailedSheet.Range("E2", Cells(k - 1, 5).Address).NumberFormat = "mm-dd-yyyy"
    DetailedSheet.Range("A1:B1").EntireColumn.AutoFit
    
FormatReport:
    'Format the Report sheet
    PasteBook.Worksheets("Report").Range("A1:A2, F1, O1, R1, Z1").EntireColumn.AutoFit
    
    'Everything worked
    NewSaveBook = True
    
Footer:

End Function

Function ReadyToSave(CoverSheet As Worksheet, ReportSheet As Worksheet, RecordsSheet As Worksheet) As Boolean
'Validates all of the needed information before saving

    Dim TableStart As Range
    Dim LRow As Long
    Dim LCol As Long
    Dim CoverName As String
    Dim CoverDate As String
    Dim CoverCenter As String
    
    'Make sure information has been entered on the Cover page
    ReadyToSave = False
    
    CoverName = CoverSheet.Range("A3").Value
    CoverDate = CoverSheet.Range("A4").Value
    CoverCenter = CoverSheet.Range("A5").Value

    If Not Len(CoverName) > 0 Then
        MsgBox ("Please enter your name on the Cover Page")
        GoTo Footer
    ElseIf Not Len(CoverDate) > 0 Then
        MsgBox ("Please enter the date on the Cover Page")
        GoTo Footer
    ElseIf Not Len(CoverCenter) > 0 Then
        MsgBox ("Please select your center from the dropdown menu on the Cover Page.")
        GoTo Footer
    End If
    
    'Make sure there are activities tabulated on the Report sheet
    Set TableStart = ReportSheet.Range("A:A").Find("Select", , xlValues, xlWhole)
    
    If TableStart Is Nothing Then
        MsgBox ("Something has gone wrong on the Report Page. Please clear it and tabulate your activities again.")
        GoTo Footer
    End If
    
    LRow = TableStart.Offset(0, 1).EntireColumn.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row 'We want to search the 2nd column
    If Not LRow > TableStart.Row + 1 Then 'There is a totals row under the header
        MsgBox ("You have no activities tabulated on the Report Page.")
        GoTo Footer
    End If
    
    'Now check the Records sheet. We want the table to go past two spacers, "H BREAK" and "V BREAK"
    LRow = RecordsSheet.Range("A:A").Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    LCol = RecordsSheet.Range("1:1").Find("*", SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
    
    If RecordsSheet.Cells(LRow, 1).Value = "H BREAK" Then
        MsgBox ("You have no student attendance saved. Please parse the roster and tabulate your activities.")
        GoTo Footer
    ElseIf RecordsSheet.Cells(1, LCol).Value = "V BREAK" Then
        MsgBox ("You have no activities saved. Please tabulate your activities.")
        GoTo Footer
    End If
    
    'Set to True
    ReadyToSave = True
    
Footer:

End Function
