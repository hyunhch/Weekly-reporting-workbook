Attribute VB_Name = "RecordsSubs"
Option Explicit

Sub ClearRecords(Optional ToDelete As String)
'Clear the Records sheet of all students and activities

    Dim RecordsSheet As Worksheet
    Dim StartCell As Range
    Dim LRow As Long
    Dim FRow As Long
    Dim LCol As Long
    Dim FCol As Long
    Dim i As Long
    
    Set RecordsSheet = Worksheets("Records Page")
    
    'H BREAK and V BREAK define where the data start
    FRow = RecordsSheet.Range("A:A").Find("H BREAK", , xlValues, xlWhole).Row
    LRow = RecordsSheet.Range("A:A").Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    FCol = RecordsSheet.Range("1:1").Find("V BREAK", , xlValues, xlWhole).Column
    LCol = RecordsSheet.Range("1:1").Find("*", SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
    
    'If just activities or activities + students are being deleted
    If ToDelete = "Labels" Then
        GoTo DelActivities
    End If
    
    'Check if there are any students
    If FRow = LRow Then
        GoTo DelActivities
    End If
    
    'The ClearSheet() sub deletes rows, so we can call it here
    Set StartCell = RecordsSheet.Cells(FRow + 1, 1)
    Call ClearSheet(StartCell, 1, RecordsSheet)
    
DelActivities:
    If FCol = LCol Then
        GoTo Footer
    End If
    
    'Clear the activity information
    For i = LCol To FCol + 1 Step -1
        RecordsSheet.Range(Cells(1, i).Address, Cells(LRow, i).Address).ClearContents
    Next i
    
Footer:

End Sub

Function PushRosterNames() As Boolean
'Whenever the roster is parsed, compare the names on the Roster sheet with the Records sheet
'Add new students, give a prompt to delete missing ones

    Dim RosterSheet As Worksheet
    Dim RecordsSheet As Worksheet
    Dim RosterNames As Range
    Dim FRow As Long
    Dim LRow As Long
    
    PushRosterNames = True
    
    Set RosterSheet = Worksheets("Roster Page")
    Set RecordsSheet = Worksheets("Records Page")
    Set RosterNames = RosterSheet.ListObjects("RosterTable").ListColumns("First").DataBodyRange
    LRow = RecordsSheet.Range("A:A").Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    'Student names start after a row called "H BREAK"
    FRow = RecordsSheet.Range("A:A").Find("H BREAK", , xlValues, xlWhole).Row
    
    'Unprotect
    Call UnprotectCheck(RecordsSheet)
    
    'If there are no students yet, we can just copy and paste
    If LRow = FRow Then
        RosterNames.Copy
        RecordsSheet.Cells(LRow + 1, 1).PasteSpecial xlPasteValues
        
        RosterNames.Offset(0, 1).Copy
        RecordsSheet.Cells(LRow + 1, 2).PasteSpecial xlPasteValues
        
        GoTo SkipMatching
    End If
    
    'If we already have students, flag all of the ones that are on both sheets
    Dim RecordsNames As Range
    Dim MatchCell As Range
    Dim c As Range
    Dim i As Long
    
    Set RecordsNames = RecordsSheet.Range(Cells(FRow + 1, 1).Address, Cells(LRow, 1).Address)
    
    'First, make sure there aren't any blank rows
    For i = LRow To FRow Step -1
        If Not Len(RecordsSheet.Cells(i, 1).Value) > 0 Then
            RecordsSheet.Cells(i, 1).EntireRow.Delete
        End If
    Next i
    
    'We'll use colors for flagging, so first unflag everything
    RecordsNames.Cells.Interior.Pattern = xlNone
    
    'Add any new students and flag them along with the existing students
    i = 1
    For Each c In RosterNames
        Set MatchCell = NameMatch(c, RecordsNames)
        If MatchCell Is Nothing Then
            RecordsSheet.Cells(LRow + i, 1).Value = c.Value
            RecordsSheet.Cells(LRow + i, 2).Value = c.Offset(0, 1).Value
            RecordsSheet.Cells(LRow + i, 1).Interior.Color = vbRed
            i = i + 1
        Else
            MatchCell.Cells.Interior.Color = vbRed
        End If
    Next c

    'Make a list of all students not on the roster
    Dim MissingList As String
    Dim RemoveStudents As Long
    
    For Each c In RecordsNames
        If c.Cells.Interior.Color <> vbRed Then
            MissingList = MissingList + c.Value + " " + c.Offset(0, 1).Value + vbCr
        End If
    Next c
    
    'Only prompt if any are missing
    If Len(MissingList) > 0 Then
        RemoveStudents = MsgBox("The following students are no longer on your roster: " _
        & vbCr & MissingList & vbCr & "They will be removed from all saved activities. Do you wish to continue?", vbQuestion + vbYesNo + vbDefaultButton2)
    End If
    
    'If yes is chosen, remove any row that isn't flagged red
    If RemoveStudents = vbYes Then
        For i = LRow To FRow + 1 Step -1
            If RecordsSheet.Cells(i, 1).Interior.Color <> vbRed Then
                RecordsSheet.Cells(i, 1).EntireRow.Delete
            End If
        Next i
    ElseIf RemoveStudents = vbNo Then
        MsgBox ("If you wish to keep these students' attendance in any saved activities, add them to your roster and parse it again.")
        PushRosterNames = False
    End If
    
    'If yes, bring up a user form to select which ones to remove
    'I've decided to not use the user form. Clicking yes will remove all missing students from the Records page
    'If RemoveStudents = vbYes Then
        'Dim TempArray() As String
        'Dim NameArray() As String
        'Dim NewForm As New RemoveStudentsForm
        'Dim j As Long
        
    'Transform the string to a 2D array
        'TempArray = Split(Replace(MissingList, vbCr, " "), " ")
        'ReDim NameArray(1 To UBound(TempArray) - 1, 1 To 2)
        
        'j = 1
        'For i = 0 To UBound(TempArray) - 1
            'NameArray(j, 1) = TempArray(i)
            'i = i + 1
            'NameArray(j, 2) = TempArray(i)
            'j = j + 1
        'Next i
            
        'NewForm.SetArray NameArray
        'NewForm.Show
    'End If
    
SkipMatching:
    'Unflag everything
    RecordsSheet.Cells.Interior.Pattern = xlNone

Footer:

End Function

Sub RemoveFromRecords(NamesArray() As String, NumRows As Long)
'Passed from the RemoveStudents user form, selected students are removed from the Records sheet

    Dim RecordsSheet As Worksheet
    Dim NewSheet As Worksheet
    Dim SearchRange As Range
    Dim DelRange As Range
    Dim c As Range
    Dim FRow As Long
    Dim LRow As Long
    Dim i As Long

    Set RecordsSheet = Worksheets("Records Page")
    Set NewSheet = ThisWorkbook.Sheets.Add

    'Print out the passed names on the new sheet
    For i = 1 To NumRows
        NewSheet.Cells(i, 1) = NamesArray(1, i)
        NewSheet.Cells(i, 2) = NamesArray(2, i)
    Next i
    
    'Define the range to search on the Records sheet
    FRow = RecordsSheet.Range("A:A").Find("H BREAK", , xlValues, xlWhole).Row 'Start at the row below this
    LRow = RecordsSheet.Range("A:A").Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    Set SearchRange = RecordsSheet.Range(Cells(FRow + 1, 1).Address, Cells(LRow, 1).Address)
    
    'Match with the names on the temp sheet, deleting from the Records sheet
    For i = 1 To NumRows
        Set c = NameMatch(NewSheet.Cells(i, 1), SearchRange)
        If Not c Is Nothing Then
            c.EntireRow.Delete
        End If
    Next i

    'Delete the temporary sheet
    NewSheet.Delete
    
Footer:

End Sub

Sub TranslateAttendance(TargetSheet As Worksheet)
'Deletes empty rows, then converts "0" to ""
    
    Dim FRow As Long
    Dim LRow As Long
    Dim i As Long
    Dim c As Range
    
    FRow = TargetSheet.Range("A:A").Find("Select", , xlValues, xlWhole).Row + 1
    LRow = TargetSheet.Range("B:B").Find("*", , SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row 'Searching the name column
    
    'Delete if there is no name or if they weren't marked present/absent
    For i = LRow To FRow Step -1
        If Not (Len(TargetSheet.Cells(i, 2).Value)) > 0 Or Not (Len(TargetSheet.Cells(i, 1).Value)) > 0 Then
            TargetSheet.Cells(i, 1).EntireRow.Delete
        End If
    Next i
    
        'Change all "0"s to ""
        For i = LRow To FRow Step -1
        If TargetSheet.Cells(i, 1).Value = "0" Then
            TargetSheet.Cells(i, 1).Value = ""
        End If
    Next i
End Sub

Function SaveCheck(ActivityLabel As String) As Boolean
'Checks if an activity has been saved
    
    Dim RecordsPage As Worksheet
    Dim MatchCell As Range
    
    Set RecordsPage = Worksheets("Records Page")
    Set MatchCell = RecordsPage.Range("1:1").Find(ActivityLabel, , xlValues, xlWhole)
    
    If MatchCell Is Nothing Then
        SaveCheck = False
    Else
        SaveCheck = True
    End If

End Function


