Attribute VB_Name = "Buttons"
Option Explicit

Sub ReadRoster()
'Read in the roster, make a sortable/filterable table, add Marlett boxes, conditional formatting

    Dim RosterSheet As Worksheet
    Dim RosterTableStart As Range

    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    Set RosterSheet = Worksheets("Roster Page")
    Set RosterTableStart = RosterSheet.Range("A6")
    
    If RosterSheet.ProtectContents = True Then
        RosterSheet.Unprotect
    End If
    
    'The column headers need to remain unprotected to allow sorting
    'However, the column names and order need to remain constant
    Dim ColNames() As String
    Dim HeaderRange As Range
    
    ColNames = Split("Select;First;Last;Ethnicity;Gender;Credits;Major;Notes", ";")
    Set HeaderRange = RosterSheet.Range(Cells(RosterTableStart.Row, RosterTableStart.Column).Address, _
        Cells(RosterTableStart.Row, UBound(ColNames) + 1).Address)
    
    HeaderRange.Value = ColNames
    
    'Remove any formatting in the header
    If RosterSheet.AutoFilterMode = True Then
        RosterSheet.AutoFilterMode = False
    End If
    
    'Delete any table objects
    Dim OldTable As ListObject
    
    For Each OldTable In RosterSheet.ListObjects
        OldTable.Unlist
    Next OldTable
    
    'Make sure there are some students added
    Dim LRow As Long
    
    LRow = RosterSheet.Cells(Rows.Count, 2).End(xlUp).Row
    
    If LRow = RosterTableStart.Row Then
        MsgBox ("Please add at least one student.")
        GoTo Footer
    End If
    
    'Make a table object and add conditional formatting
    Dim LCol As Long
    Dim RosterTableRange As Range
    
    LRow = RosterSheet.Cells(Rows.Count, 2).End(xlUp).Row
    LCol = RosterSheet.Cells(RosterTableStart.Row, Columns.Count).End(xlToLeft).Column
    
    Set RosterTableRange = RosterSheet.Range(Cells(RosterTableStart.Row, RosterTableStart.Column), Cells(LRow, LCol))
    RosterSheet.ListObjects.Add(xlSrcRange, RosterTableRange, , xlYes).Name = "AllStudentsTable"
    RosterSheet.ListObjects("AllStudentsTable").ShowTableStyleRowStripes = False
    
    Call TableFormat(RosterSheet.ListObjects("AllStudentsTable"), RosterSheet)
    RosterTableRange.Columns.AutoFit

    'Add Marlett boxes
    Dim BoxRange As Range
    Dim SelectAllRange As Range
    
    Set BoxRange = RosterSheet.Range(Cells(RosterTableStart.Row + 1, 1).Address, Cells(LRow, 1).Address)
    Set SelectAllRange = RosterTableStart.Offset(-1, 0)
    
    Call AddMarlettBox(BoxRange, RosterSheet)
    Call AddSelectAll(SelectAllRange, RosterSheet)
    
Footer:
    'Only lock cells above the headers
    Dim i As Long
    
    RosterSheet.Protect , userinterfaceonly:=True, AllowSorting:=True, AllowFiltering:=True
    RosterSheet.EnableSelection = xlUnlockedCells
    
    RosterSheet.Cells.Locked = False
    For i = 1 To RosterTableStart.Row - 1
        RosterSheet.Cells(i, 1).EntireRow.Locked = True
    Next i

    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    
End Sub

Sub RosterSheetClear()
'Clear everything in the Roster Page

    Dim RosterSheet As Worksheet
    Dim RosterTableStart As Range
    
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    Set RosterSheet = Worksheets("Roster Page")
    Set RosterTableStart = RosterSheet.Range("A6")
    
    Call ClearRoster(RosterTableStart.Offset(1, 0), 0, RosterSheet)
    
    'Put the column headers back in
    Dim ColNames() As String
    Dim HeaderRange As Range
    
    ColNames = Split("Select;First;Last;Ethnicity;Gender;Credits;Major;Notes", ";")
    Set HeaderRange = RosterSheet.Range(Cells(RosterTableStart.Row, RosterTableStart.Column).Address, _
        Cells(RosterTableStart.Row, UBound(ColNames) + 1).Address)
    
    HeaderRange.Value = ColNames
    
    'Only lock cells above the headers
    Dim i As Long
    
    RosterSheet.Protect , userinterfaceonly:=True, AllowSorting:=True, AllowFiltering:=True
    RosterSheet.EnableSelection = xlUnlockedCells
    
    RosterSheet.Cells.Locked = False
    For i = 1 To RosterTableStart.Row - 1
        RosterSheet.Cells(i, 1).EntireRow.Locked = True
    Next i
    
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.EnableEvents = True

End Sub

Sub PullRoster()
'Pull students from the roster page to the cover page

    Dim CoverSheet As Worksheet
    Dim TableStart As Range
    Dim TableRange As Range
    Dim LRow As Long
    Dim LCol As Long
    Dim SelectAllRange As Range
    Dim SelectButton As Shape
    
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    'Clear contents first
    Set CoverSheet = Worksheets("Cover Page")
    Set TableStart = CoverSheet.Range("A6")
    
    If CoverSheet.ProtectContents = True Then
        CoverSheet.Unprotect
    End If

    Call ClearRoster(TableStart, 1, CoverSheet)
    
    'Clear the Select All button
    Set SelectAllRange = TableStart.Offset(-1, 0)
    
    For Each SelectButton In CoverSheet.Shapes
        If SelectButton.TopLeftCell.Address = SelectAllRange.Address Then
            SelectButton.Delete
        End If
    Next SelectButton
    
    'Copy over the roster and verify we have students
    Call CopyRoster(TableStart)
    
    LRow = CoverSheet.Cells(Rows.Count, TableStart.Offset(0, 1).Column).End(xlUp).Row
    LCol = CoverSheet.Cells(TableStart.Row, Columns.Count).End(xlToLeft).Column
    
    If LRow < TableStart.Row + 1 Then
        MsgBox ("There aren't any students here." & vbCr & _
        "Please enter your students on the roster page")
        GoTo Footer
    End If
    
    'Make table object. Unlock the cells in the table to allow for sorting
    Set TableRange = CoverSheet.Range(Cells(TableStart.Row, TableStart.Column), Cells(LRow, LCol))
    CoverSheet.ListObjects.Add(xlSrcRange, TableRange, , xlYes).Name = "RosterTable"
    CoverSheet.ListObjects("RosterTable").ShowTableStyleRowStripes = False
    TableRange.Locked = False

    'Add Marlett Boxes and Select boxes
    Dim BoxRange As Range
    Set BoxRange = CoverSheet.Range(Cells(TableStart.Row + 1, 1).Address, Cells(LRow, 1).Address)

    Call AddMarlettBox(BoxRange, CoverSheet)
    Call AddSelectAll(SelectAllRange, CoverSheet)

    'Conditional Formatting
    Call TableFormat(CoverSheet.ListObjects("RosterTable"), CoverSheet)
    TableRange.Columns.AutoFit
    
    'Add in totals to the Report sheet
    Call ClearReportTotals
    Call PullReportTotals
    
Footer:
    CoverSheet.Protect , userinterfaceonly:=True, AllowSorting:=True, AllowFiltering:=True
    CoverSheet.EnableSelection = xlUnlockedCells

    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.EnableEvents = True

End Sub

Sub AddSheet()
'Take the selected students and populate a new sheet

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    
    Dim CoverSheet As Worksheet
    Dim LRow As Long
    Dim LCol As Long
    Dim TableStart As Range
    
    'Make sure we have at least one student displayed and checked
    Set CoverSheet = Worksheets("Cover Page")
    Set TableStart = CoverSheet.Range("A6")
    LRow = CoverSheet.Cells(Rows.Count, 2).End(xlUp).Row
    LCol = CoverSheet.Cells(TableStart.Row, Columns.Count).End(xlToLeft).Column
    
    If LRow = TableStart.Row Then
        MsgBox ("You have no students displayed. " & vbCr & _
            "Please make sure you have at least one student in your roster.")
        GoTo Footer
    End If

    If AnyChecked(TableStart.Row + 1, LRow, CoverSheet) = False Then
        MsgBox ("You have no students selected")
        GoTo Footer
    End If
    
    'Create a new sheet and fill it
    Dim ActivitySheet As Worksheet
    Dim NewTableStart As Range
    Dim CenterName As String
    Dim DirectorName As String

    CenterName = CoverSheet.Range("B2").Value
    DirectorName = CoverSheet.Range("B1").Value
    Set ActivitySheet = Sheets.Add(After:=Sheets(Sheets.Count))
    Set NewTableStart = ActivitySheet.Range("A6")

    Call PopulateSheet(ActivitySheet, CenterName, DirectorName)
    Call CopySelectedStudents(ActivitySheet, NewTableStart, LRow, LCol, TableStart)
    Call RenameSheets
    
Footer:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True

End Sub

Sub RemoveSheet()
'Delete an activity activity sheet and rename sheets, if needed

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    
    Dim Confirm As Long
    
    Confirm = MsgBox("Are you sure you want to delete this sheet? " & vbCr & _
        "This cannot be undone.", vbQuestion + vbYesNo + vbDefaultButton2)
        
    If Confirm = vbYes Then
        ActiveSheet.Delete
        Call RenameSheets
    End If
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    
End Sub

Sub TabulateCaller()
'This only exists so TabulateActivities() can be called with .OnAction
    
    Dim ReportSheet As Worksheet
    Dim ActivitySheet As Worksheet
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    
    Set ReportSheet = Worksheets("Report Page")
    Set ActivitySheet = ActiveSheet
    
    If ReportSheet.ProtectContents = True Then
        ReportSheet.Unprotect
    End If
    
    'Confirm that the needed information is added
    If NameDatePractice(ActivitySheet) = False Then
        GoTo Footer
    End If
    
    Call TabulateActivities(ActiveSheet)
    ReportSheet.Activate

Footer:
    ReportSheet.Protect , userinterfaceonly:=True, AllowSorting:=True, AllowFiltering:=True, AllowFormattingColumns:=True

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    
End Sub

Sub CompileActivities()
'Button to look at each sheet, tabulate the values, and put them into a new on

    Dim ActivitySheet As Worksheet
    Dim ReportSheet As Worksheet
    Dim ActivitySheetCount As Long
    Dim LRow As Long
        
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    
    Set ReportSheet = Worksheets("Report Page")
    If ReportSheet.ProtectContents = True Then
        ReportSheet.Unprotect
    End If
    
    'Search sheet names to make sure we have at least one activity sheet
    ActivitySheetCount = 0
    
    For Each ActivitySheet In ThisWorkbook.Worksheets
        If InStr(ActivitySheet.Name, "Activity") > 0 Then
            ActivitySheetCount = ActivitySheetCount + 1
        End If
    Next ActivitySheet
    
    If ActivitySheetCount = 0 Then
        MsgBox ("You don't have any activities." & vbNewLine & _
            "Please add at least one activity.")
        GoTo Footer
    End If
    
    For Each ActivitySheet In ThisWorkbook.Worksheets
        If InStr(ActivitySheet.Name, "Activity") > 0 Then
            If StudentsSelected(ActivitySheet) = False Then
                Exit Sub
            ElseIf NameDatePractice(ActivitySheet) = False Then
                Exit Sub
            End If
        End If
    Next
    
    'Make sure the totals are pulled into the report sheet
    LRow = ReportSheet.Cells(Rows.Count, 2).End(xlUp).Row
    
    If LRow < 5 Then
        Call PullReportTotals
    End If
    
    TabulateSelectedSheetsForm.Show

Footer:
    ReportSheet.Protect , userinterfaceonly:=True, AllowSorting:=True, AllowFiltering:=True, AllowFormattingColumns:=True

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    
End Sub

Sub AddSelected()
'Add selected students to an activity sheet
'Will need a dropdown or other way to select which sheet

    Dim CoverSheet As Worksheet
    Dim LRow As Long
    Dim SearchRange As Range
    Dim TableStart As Range
    
    Set CoverSheet = Worksheets("Cover Page")
    Set TableStart = CoverSheet.Range("A6")
    LRow = CoverSheet.Cells(Rows.Count, 2).End(xlUp).Row

    'Make sure a student is selected
    Set SearchRange = CoverSheet.Range(Cells(TableStart.Row, 1).Address, Cells(LRow, 1).Address).Find("a", LookIn:=xlValues)
    
    If SearchRange Is Nothing Then
        MsgBox ("You don't have any students selected.")
        Exit Sub
    End If
    
    'Make sure there's an activity sheet students can be added to
    Dim ActivitySheet As Worksheet
    Dim ActivitySheetCount As Long
    
    ActivitySheetCount = 0
    
    For Each ActivitySheet In ThisWorkbook.Worksheets
        If InStr(ActivitySheet.Name, "Activity") > 0 Then
            ActivitySheetCount = ActivitySheetCount + 1
        End If
    Next ActivitySheet

    If ActivitySheetCount = 0 Then
        MsgBox ("You don't have any activities." & vbNewLine & _
            "Please add at least one activity.")
        Exit Sub
    End If

    AddSelectedStudentsForm.Show

End Sub

Sub RemoveSelected()
'Remove selected rows from a sheet

    Dim DelSheet As Worksheet
    Dim IsChecked As Range
    Dim LRow As Long
    Dim TableStart As Range
    Dim ProtectedSheet As Boolean
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    
    Set DelSheet = ActiveSheet
    ProtectedSheet = False
    If DelSheet.ProtectContents = True Then
        ProtectedSheet = True
        DelSheet.Unprotect
    End If
    
    'Find where the table starts. This should be the same on every sheet
    Set TableStart = DelSheet.Range("A:A").Find("Select", , xlValues, xlWhole)
    
    If TableStart Is Nothing Then
        MsgBox ("Something has gone wrong. Please try on a fresh sheet")
        GoTo Reprotect
    End If
    
    'Make sure we have at least one filled row
    LRow = DelSheet.Cells(Rows.Count, 2).End(xlUp).Row
    If LRow = TableStart.Row Then
        MsgBox ("You don't have any students on this page.")
        GoTo Reprotect
    End If

    'Loop backward through the rows
    Dim NumChecked As Long
    Dim i As Long
    
    For i = LRow To TableStart.Row + 1 Step -1
        If DelSheet.Cells(i, 1).Value <> "" Then
            DelSheet.Cells(i, 1).EntireRow.Delete
            NumChecked = NumChecked + 1
        End If
    Next i

    If NumChecked < 1 Then
        MsgBox ("You don't have any rows selected.")
    End If
    
Reprotect:

    If ProtectedSheet = False Then
        GoTo Footer
    End If
    
    DelSheet.Protect , userinterfaceonly:=True, AllowSorting:=True, AllowFiltering:=True
    
Footer:

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True

End Sub

Sub ClearReport()
'Delete everything from the report sheet
    
    Dim ReportSheet As Worksheet
    Dim ReportStart As Range
    Dim DelRange As Range
    Dim ClearAll As Long

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    
    Set ReportSheet = Worksheets("Report Page")
    Set ReportStart = ReportSheet.Range("A6")
    Set DelRange = ReportSheet.Range(Cells(ReportStart.Row + 1, 1), Cells(Rows.Count, Columns.Count))
    
    If ReportSheet.ProtectContents = True Then
        ReportSheet.Unprotect
    End If
    
    ClearAll = MsgBox("Are you sure you want to clear all content?" & vbCrLf & "This cannot be undone.", vbQuestion + vbYesNo + vbDefaultButton2, "")
    
    If ClearAll = vbYes Then
        With DelRange
            .ClearContents
            .FormatConditions.Delete
            .Font.Name = "Calibri"
        End With
        Call ClearReportTotals
    End If
    
    ReportSheet.Protect , userinterfaceonly:=True, AllowSorting:=True, AllowFiltering:=True, AllowFormattingColumns:=True
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    
End Sub

Sub PullReportTotals()
'Also called when roster is pulled

    Dim RosterSheet As Worksheet
    Dim ReportSheet As Worksheet
    Dim CoverSheet As Worksheet
    Dim TotalStart As Range
    Dim RosterLRow As Long
    Dim AllStudents As Range
    Dim RaceRange As Range
    Dim GenderRange As Range
    Dim GradeRange As Range
    Dim TempArray As Variant
    Dim CenterString As String
    Dim NameString As String
    Dim DateString As String
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    
    Set RosterSheet = Worksheets("Roster Page")
    Set ReportSheet = Worksheets("Report Page")
    Set CoverSheet = Worksheets("Cover Page")
    
    If ReportSheet.ProtectContents = True Then
        ReportSheet.Unprotect
        Call ClearReportTotals
    End If
    
    'First pull in name, center, and date
    With CoverSheet
        NameString = .Range("B1").Value
        CenterString = .Range("B2").Value
        DateString = .Range("B3").Value
    End With
    
    With ReportSheet
        .Range("B7").Value = CenterString
        .Range("C7").Value = NameString
        .Range("D7").Value = DateString
        .Range("D7").NumberFormat = "yyyy-mm-dd"
        .Range("E7").Value = "Total Students"
        .Range("F7").Value = "All students in roster"
    End With
    
    Set TotalStart = ReportSheet.Range("G6")
    Set RaceRange = ReportSheet.Range("H1:O1").Offset(TotalStart.Row, 0)
    Set GenderRange = ReportSheet.Range("P1:R1").Offset(TotalStart.Row, 0)
    Set GradeRange = ReportSheet.Range("S1:V1").Offset(TotalStart.Row, 0)
    
    RosterLRow = RosterSheet.Cells(Rows.Count, 2).End(xlUp).Row
    Set AllStudents = RosterSheet.Range(Cells(2, 1).Address, Cells(RosterLRow, 1).Address)
    
    TempArray = DemoTabulate(RosterSheet, RosterLRow - 1, AllStudents, "Race")
    RaceRange = TempArray
    Erase TempArray
    
    TempArray = DemoTabulate(RosterSheet, RosterLRow - 1, AllStudents, "Gender")
    GenderRange = TempArray
    Erase TempArray
    
    TempArray = DemoTabulate(RosterSheet, RosterLRow - 1, AllStudents, "Grade")
    GradeRange = TempArray
    Erase TempArray
    
    'Total
    TotalStart.Offset(1, 0).Value = RosterLRow - 1

    ReportSheet.Protect , userinterfaceonly:=True, AllowSorting:=True, AllowFiltering:=True, AllowFormattingColumns:=True
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True

End Sub

Sub ClearReportTotals()

    Dim ReportSheet As Worksheet
    Dim TotalStart As Range
    Dim LCol As Long
    Dim DelRange As Range
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    
    Set ReportSheet = Worksheets("Report Page")
    
    If ReportSheet.ProtectContents = True Then
        ReportSheet.Unprotect
    End If

    Set TotalStart = ReportSheet.Range("A6")
    LCol = ReportSheet.Cells(TotalStart.Row, Columns.Count).End(xlToLeft).Column
    Set DelRange = ReportSheet.Range(Cells(TotalStart.Row + 1, TotalStart.Column).Address, Cells(TotalStart.Row + 1, LCol).Address)
    
    DelRange.ClearContents
    DelRange.Font.Name = "Calibri"

    ReportSheet.Protect , userinterfaceonly:=True, AllowSorting:=True, AllowFiltering:=True, AllowFormattingColumns:=True
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    
End Sub

Sub SharePointExport()
'reformat the data and export a new spreadsheet to SharePoint. Use a dynamic name with the center name and date
                
    Dim CopyBook As Workbook
    Dim PasteBook As Workbook
    Dim CoverSheet As Worksheet
    Dim ReportSheet As Worksheet
    Dim AttendenceSheet As Worksheet
    Dim ActivitySheet As Worksheet
    Dim SpPath As String
    Dim SubDate As String
    Dim CenterString As String
    Dim FileName As String
    Dim NameString As String
    Dim DateString As String
    Dim SubTime As String
    Dim LRow As Long
        
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    
    Set CoverSheet = Worksheets("Cover Page")
    Set ReportSheet = Worksheets("Report Page")
    
    LRow = ReportSheet.Cells(Rows.Count, 2).End(xlUp).Row
    If LRow = 7 Then
        MsgBox ("You don't have anything tabulated on the Report sheet")
        GoTo Footer
    End If
    
    'Check to make sure there is a name and center entered
    With CoverSheet
        NameString = .Range("B1").Value
        CenterString = .Range("B2").Value
        DateString = .Range("B3").Value
    End With
    
    If Len(CenterString) < 1 Or Len(NameString) < 1 Or Len(DateString) < 1 Then
        MsgBox ("Please enter in your name, center, and the date on the Cover Page.")
        GoTo Footer
    End If
    
    'Grabs the date and time of  submission
    SubDate = Format(Date, "yyyy-mm-dd")
    SubTime = Format(Time, "hh-nn AM/PM")
    
    'Create a file name based on the center and date of submission. The center *must* be filled
    'Path to the folder these will be saved in
    SpPath = "https://uwnetid.sharepoint.com/sites/partner_university_portal/Data%20Portal/Report%20Submissions/"
    FileName = CenterString & " " & SubDate & "." & SubTime & ".xlsm"
    
    'We want the Narrative and Report sheets in a new workbook
    Set CopyBook = ActiveWorkbook
    Set PasteBook = Workbooks.Add
    
    'Copy and reformat report page
    Call CopyReport(CopyBook, PasteBook)
    
    PasteBook.SaveAs FileName:=SpPath & "/" & FileName, FileFormat:=xlOpenXMLWorkbookMacroEnabled
    ActiveWorkbook.Close SaveChanges:=False
    
    MsgBox ("Submitted to SharePoint")
    
Footer:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    
End Sub

Sub LocalSave()
'Includes reformatted report sheet, student roster, and all individual activities merged into a single table
                
    Dim CopyBook As Workbook
    Dim PasteBook As Workbook
    Dim CoverSheet As Worksheet
    Dim ReportSheet As Worksheet
    Dim AttendenceSheet As Worksheet
    Dim ActivitySheet As Worksheet
    Dim LocalPath As String
    Dim SubDate As String
    Dim SubTime As String
    Dim CenterString As String
    Dim FileName As String
    Dim SaveName As Variant 'This will default to the suggested file name and in the same directory
    Dim NameString As String
    Dim DateString As String
    Dim LRow As Long
        
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    
    Set CoverSheet = Worksheets("Cover Page")
    Set ReportSheet = Worksheets("Report Page")
    
    LRow = ReportSheet.Cells(Rows.Count, 2).End(xlUp).Row
    If LRow = 7 Then
        MsgBox ("You don't have anything tabulated on the Report sheet")
        GoTo Footer
    End If
    
    'Check to make sure there is a name and center entered
    With CoverSheet
        NameString = .Range("B1").Value
        CenterString = .Range("B2").Value
        DateString = .Range("B3").Value
    End With
    
    If Len(CenterString) < 1 Or Len(NameString) < 1 Or Len(DateString) < 1 Then
        MsgBox ("Please enter in your name, center, and the date")
        GoTo Footer
    End If

    'Grabs the date and time of  submission
    SubDate = Format(Date, "yyyy-mm-dd")
    SubTime = Format(Time, "hh-nn AM/PM")

    'Create a file name based on the center and date of submission. The center *must* be filled
    'Path to the folder these will be saved in
    LocalPath = GetLocalPath(ThisWorkbook.path)
    'Debug.Print LocalPath
    FileName = CenterString & " " & SubDate & "." & SubTime & ".xlsm"
    
    'We want the Narrative and Report sheets in a new workbook
    Set CopyBook = ActiveWorkbook
    Set PasteBook = Workbooks.Add

    'Copy and reformat report page
    Call CopyReport(CopyBook, PasteBook)
    
    'Copy over attendence
    PasteBook.Sheets.Add(After:=PasteBook.Sheets(Sheets.Count)).Name = "Attendence Report"
    Set AttendenceSheet = PasteBook.Worksheets("Attendence Report")
    
    For Each ActivitySheet In CopyBook.Worksheets
        If InStr(ActivitySheet.Name, "Activity") > 0 Then
            If NameDatePractice(ActivitySheet) = False Then
                PasteBook.Close SaveChanges:=False
                Exit Sub
            Else
                Call CompiledAttendance(PasteBook, CopyBook, ActivitySheet, AttendenceSheet)
            End If
        End If
    Next ActivitySheet
    
    If Application.OperatingSystem Like "*Mac*" Then
        SaveName = Application.GetSaveAsFilename(LocalPath & "\" & FileName, "Excel Files (*.xlsm), *.xlsm")
        If SaveName = "False" Then
            ActiveWorkbook.Close SaveChanges:=False
            GoTo Footer
        End If
        ActiveWorkbook.SaveAs FileName:=LocalPath & "/" & FileName, FileFormat:=xlOpenXMLWorkbookMacroEnabled
    Else
        SaveName = Application.GetSaveAsFilename(LocalPath & "\" & FileName, "Excel Files (*.xlsm), *.xlsm")
        If SaveName = "False" Then
            ActiveWorkbook.Close SaveChanges:=False
            GoTo Footer
        End If
        ActiveWorkbook.SaveAs FileName:=SaveName, FileFormat:=xlOpenXMLWorkbookMacroEnabled
    End If

    ActiveWorkbook.Close SaveChanges:=False
    MsgBox ("Save complete")
    
Footer:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    
End Sub



