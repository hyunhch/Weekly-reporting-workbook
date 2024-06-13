Attribute VB_Name = "SharedButtonSubs"
Option Explicit

Sub SelectAllButton()
'Searches the Select column and sets the value to "a"

    Dim FRow As Long
    Dim LRow As Long
    Dim CheckRange As Range
    Dim i As Long
   
    'Unprotect
    Call UnprotectCheck(ActiveSheet)
   
    FRow = ActiveSheet.Range("A:A").Find("Select", LookIn:=xlValues).Row
    
    'In case the column name was changed or there is some other problem
    If Not FRow > 0 Then
        MsgBox ("There is a problem with the table." & vbCr & _
            "Please make sure the first column is named ""Select""")
        Exit Sub
    End If
    
    LRow = ActiveSheet.Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    
    'Check that there is at least row of students
    If Not LRow > FRow Then
        MsgBox ("Please add at least one student to the table.")
        Exit Sub
    End If
    
    'Don't allow the totals row to be checked on the Report sheet
    If ActiveSheet.Name = "Report Page" Then
        FRow = FRow + 1
    End If
    
    Set CheckRange = ActiveSheet.Range(Cells(FRow + 1, 1).Address, Cells(LRow, 1).Address)
    CheckRange.Font.Name = "Marlett"
    
    'Check all if any are blank, uncheck all if none are blank
    'Only apply to visible cells
    If Application.CountIf(CheckRange, "a") = LRow - (FRow) Then
        CheckRange.SpecialCells(xlCellTypeVisible).Value = ""
    Else
        CheckRange.SpecialCells(xlCellTypeVisible).Value = "a"
    End If
    
    Call ResetProtection
    
End Sub

Sub RemoveSelectedButton()
'Deletes every checked row

    Dim DelSheet As Worksheet
    Dim RecordsSheet As Worksheet
    Dim IsChecked As Range
    Dim TableStart As Range
    Dim TableEnd As Range
    Dim DelConfirm As Long
    Dim SaveConfirm As Long
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    
    Set DelSheet = ActiveSheet
    Set RecordsSheet = Worksheets("Records Page")
    
    'Unprotect
    Call UnprotectCheck(DelSheet)
    
    'Find where the table starts. This should be the same on every sheet
    Set TableStart = DelSheet.Range("A:A").Find("Select", , xlValues, xlWhole)
    Set TableEnd = DelSheet.Range("B:B").Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Offset(0, -1)
    
    If TableStart Is Nothing Then
        MsgBox ("Something has gone wrong. Please try on a fresh sheet")
        GoTo Footer
    End If
    
    'Make sure we have at least one student
    If Not CheckTableLength(DelSheet, TableStart) > 0 Then
        MsgBox ("You don't have any students or activities on this page.")
        GoTo Footer
    End If
    
    'Check if any students have been selected
    If CountChecks(DelSheet.Range(TableStart.Offset(1, 0), TableEnd)) = 0 Then
        MsgBox ("You don't have any rows selected")
        GoTo Footer
    End If
   
   'If on the Roster sheet, warn that this will remove the student from saved activities as well
    If DelSheet.Name = "Roster Page" Then
        Dim CopyBook As Workbook
        Set CopyBook = ActiveWorkbook

        DelConfirm = MsgBox("This will also remove the students from any recorded activities. Do you wish to continue?", vbQuestion + vbYesNo + vbDefaultButton2)
        If DelConfirm <> vbYes Then
            GoTo Footer
        End If

        'Option to export the attendance information for students being removed from the roster
        SaveConfirm = MsgBox("Do you want to save a copy of these students' attendance before removing them?", vbQuestion + vbYesNo + vbDefaultButton2)
        If SaveConfirm = vbYes Then
            Dim PasteBook As Workbook
            Set PasteBook = Workbooks.Add
            
            With PasteBook
                'Copy over the rows containing activity information on the Records sheet
                RecordsSheet.Range("A1:A3").EntireRow.Copy
                Worksheets("Sheet1").Range("A1").PasteSpecial xlPasteValues
                
                'Delete the padding cell and shift these over one column
                Worksheets("Sheet1").Range("B:B").Delete
                Worksheets("Sheet1").Range("A:A").Insert Shift:=xlToRight
                
                'Headers for student names
                Worksheets("Sheet1").Range("A5").Value = "First"
                Worksheets("Sheet1").Range("B5").Value = "Last"
            End With
            
            CopyBook.Activate
        End If
    End If

    'First check if there are any students on the Records sheet
    Dim DelRange As Range
    Dim DelCell As Range
    Dim FCell As Range
    Dim LCell As Range
    Dim CopyRange As Range
    Dim PasteRange As Range
    Dim c As Range
    Dim LRow As Long
    Dim LCol As Long
    Dim i As Long
    Dim j As Long
    
    Set FCell = RecordsSheet.Range("A:A").Find("H BREAK", , xlValues, xlWhole)
    Set LCell = RecordsSheet.Range("A:A").Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    Set DelRange = RecordsSheet.Range(FCell.Offset(1, 0), LCell)
    LCol = RecordsSheet.Range("1:1").Find("*", SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
    LRow = DelSheet.Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    
    If FCell.Row = LCell.Row Then
        GoTo Footer
    End If
    
    'As for confirmation
    'I'm not going to make this an option, students *will* be removed from saved activities if they are removed from the roster
    'DelConfirm = MsgBox("Do you wish to remove these students from all saved activities as well?", vbQuestion + vbYesNo + vbDefaultButton2)

    'Loop backward through the rows
    LRow = DelSheet.Range("B:B").Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    j = 1
    For i = LRow To TableStart.Row + 1 Step -1
        If DelSheet.Cells(i, 1).Value <> "" Then
            Set DelCell = NameMatch(DelSheet.Cells(i, 2), DelRange)
            
            'If they answered yes to exporting attendance
            If SaveConfirm = vbYes Then
                j = j + 1
                Set CopyRange = RecordsSheet.Range(DelCell, DelCell.Offset(0, LCol - 1))
                Set PasteRange = PasteBook.Worksheets("Sheet1").Range(Cells(j + 3, 1).Address, Cells(j + 3, LCol).Address)
                PasteRange.Value = CopyRange.Value
                
                For Each c In PasteRange
                    If c.Value = "a" Then
                        c.Value = 1
                    End If
                Next c
                CopyBook.Activate
            End If
            
            'If on the roster sheer, delete the row. If on an activity sheet, clear attendance information for that activity
            If Not DelCell Is Nothing Then
                If DelSheet.Name = "Roster Page" Then
                    DelCell.EntireRow.Delete
                ElseIf DelSheet.Name <> "Roster Page" And DelSheet.Name <> "Report Page" Then
                    Set CopyRange = RecordsSheet.Range(DelCell, DelCell.Offset(0, LCol - 1))
                    CopyRange.Value = ""
                End If
            End If
            DelSheet.Cells(i, 1).EntireRow.Delete
        End If
    Next i
    
    'Extra steps if it's the roster sheet
    If DelSheet.Name = "Roster Page" Then
        'Look through any open activity sheets and remove the name there as well
        Dim ActivitySheet As Worksheet
        
        For Each ActivitySheet In CopyBook.Sheets
            If ActivitySheet.Range("A1") = "Practice" Then
                Set DelRange = ActivitySheet.ListObjects(1).ListColumns("First").DataBodyRange
                
                'Make sure it's not an empty tale
                If DelRange Is Nothing Then
                    GoTo SkipSheet
                End If
                
                For i = DelRange.Rows.Count + DelRange.Row To DelRange.Row Step -1
                    Set c = ActivitySheet.Cells(i, 2)
                    Set DelCell = NameMatch(c, DelRange)
                    
                    If Not DelCell Is Nothing Then
                        DelCell.EntireRow.Delete
                    End If
                Next i
            End If
SkipSheet:
        Next ActivitySheet
    End If
 
    'Retabulate the totals and all saved activities on the Report sheet. This function parses the roster as well
    Call PullReportTotalsButton
    Call RetabulateActivities
    
Footer:
    Call ResetProtection
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True

End Sub
