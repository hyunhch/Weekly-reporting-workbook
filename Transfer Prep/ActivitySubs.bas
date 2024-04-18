Attribute VB_Name = "ActivitySubs"
Option Explicit

Function NameMatch(NameCell As Range, TargetRange As Range) As Range
'Find a student by first and last name. SearchRange is just first names
'Returns the cell for the first name

    Dim c As Range
    
    For Each c In TargetRange
        If NameCell.Value = c.Value And NameCell.Offset(0, 1).Value = c.Offset(0, 1).Value Then
            Set NameMatch = c
            GoTo Break
        End If
    Next

Break:

End Function

Function FindLabel(MatchLabel As String, Optional WhichSheet As String) As Range
'Search the attendance or label page to see if the activity already exists
'If it doesn't, returns a cell in the first empty row or column
'WhichSheet determines what sheet is being searched

    Dim SearchRange As Range
    Dim MatchCell As Range
    Dim FCell As Range
    Dim LCell As Range
    
    If WhichSheet = "ReportSheet" Then
        GoTo SearchReport
    End If
    
    'Searching the Records sheet
    Dim RecordsSheet As Worksheet
    
    Set RecordsSheet = Worksheets("Records Page")
    Set LCell = RecordsSheet.Range("1:1").Find("*", SearchOrder:=xlByColumns, SearchDirection:=xlPrevious)
    Set FCell = RecordsSheet.Range("1:1").Find("V BREAK", , xlValues, xlWhole)
    Set SearchRange = RecordsSheet.Range(FCell.Offset(0, 1), LCell)
    
    Set MatchCell = SearchRange.Find(MatchLabel, , xlValues, xlWhole)
    If MatchCell Is Nothing Then
        Set FindLabel = LCell.Offset(0, 1) 'Past the last filled column
    Else
        Set FindLabel = MatchCell
    End If
    
    GoTo Footer
    
SearchReport:
    'Searching the Report sheet
    Dim ReportSheet As Worksheet
    
    Set ReportSheet = Worksheets("Report Page")
    Set FCell = ReportSheet.Range("6:6").Find("Label", , xlValues, xlWhole)
    Set LCell = FCell.EntireColumn.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    Set SearchRange = ReportSheet.Range(FCell.Offset(2, 0), LCell)

    Set MatchCell = SearchRange.Find(MatchLabel, , xlValues, xlWhole)
    If MatchCell Is Nothing Then
        Set FindLabel = LCell.Offset(1, 0) 'Past the last filled row
    Else
        Set FindLabel = MatchCell
    End If
Footer:

End Function

Sub LoadActivity(LabelString As String)
'Called by the Load Activty userform.

    Dim ActivitySheet As Worksheet
    Dim LabelRange As Range
    
    'Make sure there isn't already a sheet open with the selected activity label
    For Each ActivitySheet In ActiveWorkbook.Worksheets
        Set LabelRange = ActivitySheet.Range("H1")
        If LabelRange.Value = LabelString Then
            GoTo Footer
        End If
    Next ActivitySheet
    
    'Find the stored activity on the Records page
    Dim RecordsSheet As Worksheet
    Dim PracticeString As String
    Dim DateValue As Date
    Dim DescriptionString As String
    
    Set RecordsSheet = Worksheets("Records Page")
    Set LabelRange = RecordsSheet.Range("1:1").Find(LabelString, , xlValues, xlWhole)
    
    'If the label can't be found. This shouldn't happen
    If LabelRange Is Nothing Then
        MsgBox ("Something has gone wrong. The activity " & LabelString & "cannot be found.")
        GoTo Footer
    End If
    
    PracticeString = LabelRange.Offset(1, 0).Value
    DateValue = LabelRange.Offset(2, 0).Value
    DescriptionString = LabelRange.Offset(3, 0).Value
    
    'We can call the same function as when creating a new activity sheet
    Call NewActivitySheet(PracticeString, DateValue, LabelString, DescriptionString, "All")
    
    'Find and copy checks
    Dim RecordsNames As Range
    Dim ActivityNames As Range
    Dim MatchCell As Range
    Dim c As Range
    Dim FRow As Long
    Dim LRow As Long
    Dim i As Long
    
    Set ActivitySheet = Worksheets(Sheets.Count)
    FRow = RecordsSheet.Range("A:A").Find("H BREAK", , xlValues, xlWhole).Row
    LRow = RecordsSheet.Range("A:A").Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    Set RecordsNames = RecordsSheet.Range(Cells(FRow + 1, 1).Address, Cells(LRow, 1).Address)
    Set ActivityNames = ActivitySheet.ListObjects(1).ListColumns("First").DataBodyRange
    
    For Each c In RecordsNames
        Set MatchCell = NameMatch(c, ActivityNames)
        If Not MatchCell Is Nothing Then
            MatchCell.Offset(0, -1).Value = RecordsSheet.Cells(c.Row, LabelRange.Column).Value
        End If
    Next c
    
    'Delete empty rows and turn "0" into ""
    Call TranslateAttendance(ActivitySheet)
    
Footer:
    
End Sub

Sub DeleteActivity(LabelString As String)
'Delete the indicated activity from the Records sheet and Report sheet

    Dim ActivitySheet As Worksheet
    Dim LabelMatch As Range
    
    'Find on the Records sheet
    Set LabelMatch = FindLabel(LabelString, "RecordsSheet")
    
    If Not LabelMatch Is Nothing Then 'Nothing needs to be done
        LabelMatch.EntireColumn.Delete
    End If

    'Find on Report sheet
    Set LabelMatch = FindLabel(LabelString, "ReportSheet")
    
    If Not LabelMatch Is Nothing Then
        LabelMatch.EntireRow.Delete
    End If

    'If there's a sheet open with the activity, delete it
    For Each ActivitySheet In ActiveWorkbook.Sheets
        If ActivitySheet.Range("H1").Value = LabelString Then
            ActivitySheet.Delete
        End If
    Next ActivitySheet

Footer:

End Sub

Function SaveActivity() As Boolean
'Stores the names and attendance of students, uses label for activity name

    Dim ActivitySheet As Worksheet
    Dim ActivityPractice As String
    Dim ActivityDescription As String
    Dim ActivityLabel As String
    Dim ActivityDate As Date
    
    Set ActivitySheet = ActiveSheet
    SaveActivity = True
    
    With ActivitySheet
        ActivityPractice = .Range("B1").Value
        ActivityDate = .Range("B3").Value
        ActivityDescription = .Range("B4").Value
        ActivityLabel = .Range("H1").Value
    End With
    
    'Make sure the table isn't empty and at least one student is selected
    Dim SearchRange As Range
    Dim CopyRange As Range
    
    Set SearchRange = ActivitySheet.ListObjects(1).ListColumns("Select").DataBodyRange
    
    If SearchRange Is Nothing Then
        MsgBox ("You have no students on this sheet. Please add at least one.")
        SaveActivity = False
        GoTo Footer
    End If
    
    If FindChecks(SearchRange) Is Nothing Then
        MsgBox ("You have no students selected.")
        SaveActivity = False
        GoTo Footer
    End If
    
    Set CopyRange = FindChecks(SearchRange).Offset(0, 1) 'Range of students who have been checked
    
    'Find if the label is already on the Records sheet. If it isn't, the function returns the first empty column
    Dim LabelMatch As Range
    
    Set LabelMatch = FindLabel(ActivityLabel)
    
    'Define the range of names to search
    Dim RecordsSheet As Worksheet
    Dim RecordsRange As Range
    Dim PasteRange As Range
    Dim c As Range
    Dim FRow As Long
    Dim LRow As Long
    
    'Set CopyRange = FindChecks(SearchRange).Offset(0, 1) 'Range of first names of selected students
    Set RecordsSheet = Worksheets("Records Page")
    FRow = RecordsSheet.Range("A:A").Find("H BREAK", , xlValues, xlWhole).Row
    LRow = RecordsSheet.Range("A:A").Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    
    'If there are no students. This shouldn't happen
    If FRow = LRow Then
        MsgBox ("Something has gone wrong. Please parse the roster and try again.")
        SaveActivity = False
        GoTo Footer
    End If
    
    Set RecordsRange = RecordsSheet.Range(Cells(FRow + 1, 1).Address, Cells(LRow, 1).Address)
    
    'Clear any existing attendance
    RecordsRange.Offset(0, LabelMatch.Column - 1).Value = ""
    
    'Present is "a", absent is "0", unlisted students are left empty
    For Each c In SearchRange.Offset(0, 1)
        Set PasteRange = NameMatch(c, RecordsRange)
        If PasteRange Is Nothing Then 'Student on the activity sheet that is not on the Attendance sheet. This shouldn't happen
            MsgBox ("The student named " & c.Value & " " & c.Offset(0, 1).Value & " cannot be found." & vbCr & "Please parse the roster and try again")
            SaveActivity = False
            GoTo Footer
        ElseIf CopyRange Is Nothing Then 'No students checked
            RecordsSheet.Cells(PasteRange.Row, LabelMatch.Column).Value = "0"
        ElseIf Not Intersect(c, CopyRange) Is Nothing Then
            RecordsSheet.Cells(PasteRange.Row, LabelMatch.Column).Value = "a"
        Else
            RecordsSheet.Cells(PasteRange.Row, LabelMatch.Column).Value = "0"
        End If
    Next c
    
    'Add values from the Activity sheet
    With LabelMatch
        .Value = ActivityLabel
        .Offset(1, 0).Value = ActivityPractice
        .Offset(2, 0).Value = ActivityDate
        .Offset(3, 0).Value = ActivityDescription
    End With
    
    'Tabulate the activity and push to report page
    Call TabulateActivity(ActivityLabel)
            
Footer:

End Function
