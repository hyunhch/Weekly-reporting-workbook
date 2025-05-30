Attribute VB_Name = "RosterSubs"
Option Explicit

Sub NamesToRecords()
'Whenever the roster is parted, compare the names listed on the records and roster sheets
'Add new students, delete missing ones with a prompt to export Attendance

    Dim OldBook As Workbook
    Dim NewBook As Workbook
    Dim ActivitySheet As Worksheet
    Dim RosterSheet As Worksheet
    Dim RecordsSheet As Worksheet
    Dim RosterNameRange As Range
    Dim RecordsNameRange As Range
    Dim StudentAddRange As Range
    Dim StudentRemoveRange As Range
    Dim c As Range
    Dim d As Range
    Dim ExportConfirm As Long
    Dim i As Long
    Dim RosterTable As ListObject
     
    Set RosterSheet = Worksheets("Roster Page")
    Set RecordsSheet = Worksheets("Records Page")
    
    'Verify that there is a table with at least one student
    If CheckTable(RosterSheet) > 2 Then
        GoTo Footer
    End If
    
    Set RosterTable = RosterSheet.ListObjects(1)
    Set RosterNameRange = RosterTable.ListColumns("First").DataBodyRange '
    Set RecordsNameRange = FindRecordsName(RecordsSheet)
    
    Call UnprotectSheet(RecordsSheet)
    
    'Check if there are students on the RecordsSheet
    i = CheckRecords(RecordsSheet)
    If i = 1 Or i = 3 Then
        GoTo CompareNames
    End If
    
    'Copy over all students
    Set RosterNameRange = RosterNameRange.Resize(RosterNameRange.Rows.Count, 2)
    Set c = RecordsNameRange.Resize(RosterNameRange.Rows.Count, 2) 'Should just be the "H BREAK" padding cell
    Set RecordsNameRange = c.Offset(1, 0)
    
    RecordsNameRange.Value = RosterNameRange.Value
    GoTo CleanUp
    
CompareNames:
    'Compare the names on the RosterSheet and RecordsSheet
    Set StudentAddRange = FindName(RecordsSheet, RecordsNameRange, RosterNameRange, "Unique") 'On the Roster and not Records
    Set StudentRemoveRange = FindName(RosterSheet, RosterNameRange, RecordsNameRange, "Unique") 'On the Records and not Roster
    
    'Remove students
    If Not StudentRemoveRange Is Nothing Then
        'Prompt for export
        i = StudentRemoveRange.Cells.Count
        ExportConfirm = MsgBox(i & " students are no longer on your roster." _
            & vbCr & "Do you wish to export their attendance before removing them?", vbQuestion + vbYesNo + vbDefaultButton2)
        
        If ExportConfirm = vbYes Then
            Set OldBook = ActiveWorkbook
            Set NewBook = MakeNewBook(OldBook)
        
            Call ExportSimpleAttendance(RecordsSheet, NewBook, RecordsNameRange)
            Call ExportDetailedAttendance(RecordsSheet, RosterSheet, NewBook, RecordsNameRange)
            Call SaveNewBook(OldBook, NewBook)
            OldBook.Activate
        End If
        
        'Remove from any open Activity sheet
        For Each ActivitySheet In ThisWorkbook.Sheets
            If ActivitySheet.Range("A1").Value = "Practice" Then
                Call RemoveFromActivity(ActivitySheet, StudentRemoveRange)
            End If
        Next ActivitySheet
        
        'Delete students no longer on Roster
        Set c = RecordsNameRange.Resize(RecordsNameRange.Rows.Count, 2) 'To span the entire range we are searching
        Call RemoveRows(RecordsSheet, RecordsNameRange, StudentRemoveRange)
    End If

    'Add students
    If Not StudentAddRange Is Nothing Then
        'Define where to start pasting
        Set d = RecordsSheet.Range("A:A").Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Offset(1, 0)
    
        'Copy over
        i = 0
        For Each c In StudentAddRange
            d.Offset(i, 0).Value = c.Value
            d.Offset(i, 1).Value = c.Offset(0, 1).Value
            i = i + 1
        Next c
    End If
    
    Dim ws As Worksheet
    
    
    
    
CleanUp:
    'Make sure there are no duplicates or blank rows
    Set RecordsNameRange = FindRecordsName(RecordsSheet)
    Set StudentRemoveRange = FindName(RecordsSheet, RecordsNameRange, RecordsNameRange, "Duplicate")
    
    Call RemoveBlanks(RecordsSheet, RecordsNameRange, RecordsNameRange)
    
    If Not StudentRemoveRange Is Nothing Then
        Call RemoveRows(RecordsSheet, RecordsNameRange, StudentRemoveRange)
    End If
    
    'Retabulate
    Call RetabulateReport
    
Footer:
    
End Sub


