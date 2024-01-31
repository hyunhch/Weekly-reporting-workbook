Attribute VB_Name = "SaveSubs"
Option Explicit

Sub CopyReport(OldBook As Workbook, NewBook As Workbook)
'Populate values into the report sheet in a new workbook

    Dim CoverSheet As Worksheet
    Dim ReportSheet As Worksheet
    Dim CenterString As String
    Dim NameString As String
    Dim DateString As String
    Dim LRow As Long
    Dim CopyRange As Range
    
    Set CoverSheet = OldBook.Worksheets("Cover Page")
    Set ReportSheet = OldBook.Worksheets("Report Page")
    
    With CoverSheet
        NameString = .Range("B1").Value
        CenterString = .Range("B2").Value
        DateString = .Range("B3").Value
    End With
    
    ReportSheet.Cells.Copy
    
    'We need to move all of the columns to the top of the new sheet and align them
    
    With NewBook.Worksheets("Sheet1")
        .Cells.PasteSpecial Paste:=xlPasteValues
        
        'Drop empty rows and checkbox column
        .Range("A1:A3").EntireRow.Delete
        .Range("A1").EntireColumn.Delete
        
        'Move up header to first row and populate. Overwrites checkbox column
        .Range("A2").Value = CenterString
        .Range("B2").Value = NameString
        .Range("C2").Value = DateString
        .Range("D2").Value = "All Students"
        .Range("E1").Value = "Description"
        .Range("E2").Value = "Every student in the roster."
        
        'Format the date column
        LRow = .Cells(Rows.Count, 1).End(xlUp).Row
        .Range(Cells(2, 3), Cells(LRow, 3)).NumberFormat = "yyyy-mm-dd"
        
        'Rename
        .Name = "Aggregate Report"
    End With
    
End Sub

Sub CompiledAttendance(NewBook As Workbook, OldBook As Workbook, CopySheet As Worksheet, PasteSheet As Worksheet)

    Dim RosterSheet As Worksheet
    Dim CopyRange As Range
    Dim PasteRange As Range
    Dim TableStart As Range
    Dim NameString As String
    Dim CenterString As String
    Dim DateString As String
    Dim PracticeString As String
    Dim DescriptionString As String
    Dim StringArray As Variant
    Dim LRow As Long
    Dim LCol As Long
    Dim i As Long
    Dim j As Long
    
    Set RosterSheet = OldBook.Worksheets("Roster Page")
    
    With PasteSheet
        .Range("A1").Value = "Center"
        .Range("B1").Value = "Name"
        .Range("C1").Value = "Date"
        .Range("D1").Value = "Practice"
        .Range("E1").Value = "Description"
    End With
    
    'Create the rest of the column names. Programmatic in case additional columns were added
    LCol = RosterSheet.Cells(1, Columns.Count).End(xlToLeft).Column
    Set CopyRange = RosterSheet.Range(Cells(1, 2).Address, Cells(1, LCol).Address)
    Set PasteRange = PasteSheet.Range("F1")
    
    CopyRange.Copy
    PasteRange.PasteSpecial Paste:=xlPasteValues
    
    'Grab students from the activity sheet
    Set TableStart = CopySheet.Range("A6")
    
    With CopySheet
        LRow = .Cells(Rows.Count, 2).End(xlUp).Row
        LCol = .Cells(TableStart.Row, Columns.Count).End(xlToLeft).Column
        
        NameString = .Range("B1").Value
        CenterString = .Range("B2").Value
        DateString = .Range("B3").Value
        PracticeString = .Range("F1").Value
        DescriptionString = .Range("F3").Value
    End With
    
    StringArray = Array(CenterString, NameString, DateString, PracticeString, DescriptionString)
    Set PasteRange = PasteSheet.Range(Cells(Rows.Count, 1).End(xlUp).Address).Offset(1, 0)
    j = 0
    
    For i = TableStart.Row To LRow
        If CopySheet.Cells(i + 1, 1) = "a" Then
            Set CopyRange = CopySheet.Range(Cells(i + 1, 2).Address, Cells(i + 1, LCol).Address)
            CopyRange.Copy
            PasteRange.Offset(j, 5).PasteSpecial Paste:=xlPasteValues
            PasteRange.Offset(j, 0).Resize(1, 5) = StringArray
            
            j = j + 1
        End If
    Next i
    
End Sub
