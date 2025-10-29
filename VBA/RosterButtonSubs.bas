Attribute VB_Name = "RosterButtonSubs"
Option Explicit

Sub RosterAddStudentsFormButton()
'Adds selected students to a saved activity

    Dim RosterSheet As Worksheet
    Dim RecordsSheet As Worksheet
    Dim RosterNameRange As Range
    Dim RosterCheckRane As Range
    Dim AddCheck As Long
    Dim RosterTable As ListObject
    
    Set RosterSheet = Worksheets("Roster Page")
    Set RecordsSheet = Worksheets("Records Page")
    
    'Make sure there are any saved activities
    AddCheck = CheckRecords(RecordsSheet)
    If AddCheck > 2 Then
        MsgBox ("You have no saved activities.")
        GoTo Footer
    End If

    'Make sure there is a roster table, that there's at least one student, and that there's at least one checked
    If RosterSheet.ListObjects.Count = 0 Then
        GoTo Footer
    End If
    
    Set RosterTable = RosterSheet.ListObjects(1)
    If Not RosterTable.ListRows.Count > 0 Then
        GoTo Footer
    End If

    Set RosterNameRange = RosterTable.ListColumns("First").DataBodyRange
    If FindChecks(RosterNameRange.Offset(0, -1)) Is Nothing Then
        GoTo Footer
    End If

    'Show form
    AddStudentsForm.Show

Footer:

End Sub

Sub RosterLoadActivityFormButton()
'Checks to see if there are any saved activities and opens the load activity form

    Dim RecordsSheet As Worksheet
    Dim LoadCheck As Long
    
    Set RecordsSheet = Worksheets("Records Page")
    
    'Check to make sure there are any saved activities. Activities with no students are fine.
    LoadCheck = CheckRecords(RecordsSheet)
    If LoadCheck > 2 Then
        MsgBox ("You have no saved activities")
        GoTo Footer
    End If
    
    'Show form
    LoadActivityForm.Show
    
Footer:

End Sub

Sub RosterNewActivityFormButton()
'Opens form to create a new activity. Does not require any selected students

    Dim RosterSheet As Worksheet
    Dim RosterNameRange As Range
    Dim RosterCheckRange As Range
    Dim RosterTable As ListObject
    
    Set RosterSheet = Worksheets("Roster Page")
    
    'Make sure there's a parsed table with at least one student
    If CheckTable(RosterSheet) > 1 Then
        GoTo Footer
    End If
    
    'Make sure the CoverSheet is filled out
    If CheckCover <> 1 Then
        MsgBox ("Please fill out your name, the date, and your center on the Cover Page")
        
        GoTo Footer
    End If
    
    'Show form
    NewActivityForm.Show
    
Footer:

End Sub

Sub RosterClearButton()
'Delete everything, reset columns, clear records
    
    Dim RosterSheet As Worksheet
    
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    Set RosterSheet = Worksheets("Roster Page")

    Call UnprotectSheet(RosterSheet)
    
    'Show pompts for deleting and exporting
    Call RosterClear(RosterSheet, "Prompt")
    
    'Remake the table
    Call MakeRosterTable(RosterSheet)
    
Footer:
    Call ResetProtection

    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

End Sub

Sub RosterParseButton()
'Read in the roster, table with conditional formatting, Marlett boxes, push to the ReportSheet

    Dim NewBook As Workbook
    Dim OldBook As Workbook
    Dim RosterSheet As Worksheet
    Dim RecordsSheet As Worksheet
    Dim RosterNameRange As Range
    Dim RecordsDelRange As Range
    Dim RecordsNameRange As Range
    Dim RecordsRange As Range
    Dim DelRange As Range
    Dim CopyRange As Range
    Dim PasteRange As Range
    Dim c As Range
    Dim d As Range
    Dim i As Long
    Dim j As Long
    Dim DelConfirm As Long
    Dim ExportConfirm As Long
    Dim NumDuplicate As Long
    Dim NumAdded As Long
    Dim ConfirmMessage As String
    Dim RosterTable As ListObject

    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    Set OldBook = ThisWorkbook
    Set RosterSheet = Worksheets("Roster Page")
    Set RecordsSheet = Worksheets("Records Page")
    
    'Remake the Roster table to include new students, if any
    Set RosterTable = MakeRosterTable(RosterSheet)
    
    'See if we have any students. Break if we don't
    i = CheckTable(RosterSheet)
        If i > 2 Then
            GoTo Footer
        End If
    
    'Remove duplicates empty spaces
    NumDuplicate = RemoveBadRows(RosterSheet, RosterTable.DataBodyRange, RosterTable.ListColumns("First").DataBodyRange)
    
    Set RosterTable = RosterSheet.ListObjects(1)
    Set RosterNameRange = RosterTable.ListColumns("First").DataBodyRange
    
    'Check for students on the Records but not the Roster
    j = CheckRecords(RecordsSheet)
        If j = 2 Or j = 4 Then
            GoTo CopyStudents
        End If
    
    Set RecordsNameRange = FindRecordsName(RecordsSheet)
    Set RecordsDelRange = FindUnique(RecordsNameRange, RosterNameRange)
        If RecordsDelRange Is Nothing Then
            GoTo CopyStudents
        End If
    
    If PromptRemoveRecords(RecordsDelRange) <> 1 Then
        GoTo RemoveExtra
    End If
    
    'Generate a new book
    Set NewBook = ExportFromRecords(RecordsDelRange)
        If NewBook Is Nothing Then
            GoTo RemoveExtra
        End If
        
    'Save and close the new book
    If ExportLocalSave(OldBook, NewBook) > 0 Then
        NewBook.Close savechanges:=False
    End If
    
RemoveExtra:
    'Remove the extra students
    Set RecordsRange = FindRecordsRange(RecordsSheet)
    
    Call RemoveRows(RecordsSheet, RecordsRange, RecordsNameRange, RecordsDelRange)
    
CopyStudents:
    'If everyone was deleted, clear the records and report, add everyone
    i = CheckRecords(RecordsSheet)
    
    If i = 2 Or i = 4 Then
        Call RecordsClear
        Call ReportClear
        
        Set CopyRange = RosterNameRange
    Else
        Set RecordsNameRange = FindRecordsName(RecordsSheet)
        Set CopyRange = FindUnique(RosterNameRange, RecordsNameRange)
    End If

    'No students to copy
    If CopyRange Is Nothing Then
        GoTo RetabulateReport
    End If
    
    NumAdded = CopyNames(CopyRange).Cells.Count
    
    'Clean up the Records name list
    Set RecordsNameRange = FindRecordsName(RecordsSheet) 'Redefine since it may be longer or shorter
    Set RecordsRange = FindRecordsRange(RecordsSheet)
    Set RecordsDelRange = Nothing
    
    Set c = FindDuplicate(RecordsNameRange)
        If Not c Is Nothing Then
            Set RecordsDelRange = c
        End If

    Set d = FindBlanks(RecordsNameRange)
        If Not d Is Nothing Then
            Set RecordsDelRange = BuildRange(d, RecordsDelRange)
        End If
    
    If Not DelRange Is Nothing Then
        Call RemoveRows(RecordsSheet, RecordsRange, RecordsNameRange, RecordsDelRange) 'RemoveBadRows requires a table
    End If
    
RetabulateReport:
    'Tabulate the totals and push to the ReportTable, retabulate activities already listed on the Report
    Call TabulateReportTotals
    
    If CheckRecords(RecordsSheet) < 3 Then
        Call TabulateListedActivities
    End If

    'Show students added and duplicates removed
    If NumAdded > 0 Then
        ConfirmMessage = NumAdded & " students added."
    End If
    
    If NumDuplicate > 0 Then
        If Len(ConfirmMessage) > 0 Then
            ConfirmMessage = ConfirmMessage & vbCr
        End If
        
        ConfirmMessage = ConfirmMessage & NumDuplicate & " duplicates removed"
    End If
    
    If Len(ConfirmMessage) > 0 Then
        MsgBox (ConfirmMessage)
    End If

Footer:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

End Sub
