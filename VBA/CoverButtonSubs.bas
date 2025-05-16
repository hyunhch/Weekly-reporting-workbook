Attribute VB_Name = "CoverButtonSubs"
Option Explicit

Sub CoverSaveCopyButton()
'Saves a local copy of attendance records, the roster, report, and cover sheet
'More detailed than what goes to SharePoint
    
    Dim OldBook As Workbook
    Dim NewBook As Workbook
    Dim CoverSheet As Worksheet
    Dim RecordsSheet As Worksheet
    Dim ReportSheet As Worksheet
    Dim RosterSheet As Worksheet
    Dim i As Long
    Dim CenterString As String
    Dim ErrorString As String
    Dim FullErrorString As String
    Dim FileName As String
    Dim LocalPath As String
    Dim SaveName As String
    Dim SubDate As String
    Dim SubTime As String
    Dim TempArray() As Variant
    Dim IsSaved As Boolean

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    
    IsSaved = False
    
    Set CoverSheet = Worksheets("Cover Page")
    Set ReportSheet = Worksheets("Report Page")
    Set RosterSheet = Worksheets("Roster Page")
    Set RecordsSheet = Worksheets("Records Page")
    
    'Check that everything is filled out
    TempArray = GetReadyToSave(CoverSheet, ReportSheet, RecordsSheet, RosterSheet)
    
    For i = 1 To UBound(TempArray, 2)
        If TempArray(2, i) = 0 Then
            ErrorString = TempArray(1, i)
            
            GoTo IncompleteMessage
        End If
    Next i
    
    GoTo NewWorkbook
    
IncompleteMessage:
    'Display what hasn't been filled out
    Select Case ErrorString
        Case "Cover Page"
            FullErrorString = "Please completely fill out the Cover Page and retabulate your activities."
        Case "Report Page"
            FullErrorString = "There are no activities tabulated on the Report Page."
        Case "Roster Page"
            FullErrorString = "There are no students parsed on the Roster Page"
        Case "Records Page"
            FullErrorString = "There are no saved activities with students marked as present."
    End Select
    
    MsgBox FullErrorString
    GoTo Footer
    
NewWorkbook:
    'Make a new workbook
    Set OldBook = ThisWorkbook
    Set NewBook = MakeNewBook(RecordsSheet, ReportSheet, RosterSheet)
    
    'Pull in center from the CoverSheet, date, and time
    CenterString = CoverSheet.Range("A:A").Find("Center", , xlValues, xlWhole).Offset(0, 1).Value
    SubDate = Format(Date, "yyyy-mm-dd")
    SubTime = Format(Time, "hh-nn AM/PM")

    'Make a file name using the center, date, and time of submission
    FileName = CenterString & " " & SubDate & "." & SubTime & ".xlsm"
    
    'Where the OldBook is stored
    OldBook.Activate
    LocalPath = GetLocalPath(ThisWorkbook.path)
    
    'For Win and Mac
    If Application.OperatingSystem Like "*Mac*" Then
        SaveName = Application.GetSaveAsFilename(LocalPath & "\" & FileName, "Excel Files (*.xlsm), *.xlsm")
        If SaveName = "False" Then
            NewBook.Close SaveChanges:=False
            GoTo Footer
        End If
        NewBook.SaveAs FileName:=LocalPath & "/" & FileName, FileFormat:=xlOpenXMLWorkbookMacroEnabled
        'ActiveWorkbook.Close SaveChanges:=False
    Else
        SaveName = Application.GetSaveAsFilename(LocalPath & "\" & FileName, "Excel Files (*.xlsm), *.xlsm")
        If SaveName = "False" Then
            NewBook.Close SaveChanges:=False
            GoTo Footer
        End If
        NewBook.SaveAs FileName:=SaveName, FileFormat:=xlOpenXMLWorkbookMacroEnabled
        'ActiveWorkbook.Close SaveChanges:=False
    End If
    
    'Everything worked
    IsSaved = True
    
Footer:
    Call ResetProtection
    
    If IsSaved = True Then
        NewBook.Activate
    Else
        CoverSheet.Activate
    End If

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True

End Sub

Sub CoverSharePointExportButton()
'Sends the cover sheet and report to SharePoint

    Dim OldBook As Workbook
    Dim NewBook As Workbook
    Dim CoverSheet As Worksheet
    Dim RecordsSheet As Worksheet
    Dim ReportSheet As Worksheet
    Dim RosterSheet As Worksheet
    Dim i As Long
    Dim NameString As String
    Dim CenterString As String
    Dim ErrorString As String
    Dim FullErrorString As String
    Dim FileName As String
    Dim SpPath As String
    Dim SubDate As String
    Dim SubTime As String
    Dim TempArray() As Variant

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    
    Set CoverSheet = Worksheets("Cover Page")
    Set ReportSheet = Worksheets("Report Page")
    Set RosterSheet = Worksheets("Roster Page")
    Set RecordsSheet = Worksheets("Records Page")
    
    'Check that everything is filled out
    TempArray = GetReadyToSave(CoverSheet, ReportSheet, RecordsSheet, RosterSheet)
    
    For i = 1 To UBound(TempArray, 2)
        If TempArray(2, i) = 0 Then
            ErrorString = TempArray(1, i)
            
            GoTo IncompleteMessage
        End If
    Next i
    
    GoTo NewWorkbook
    
IncompleteMessage:
    'Display what hasn't been filled out
    Select Case ErrorString
        Case "Cover Page"
            FullErrorString = "Please completely fill out the Cover Page."
        Case "Report Page"
            FullErrorString = "There are no activities tabulated on the Report Page."
        Case "Roster Page"
            FullErrorString = "There are no students parsed on the Roster Page"
        Case "Records Page"
            FullErrorString = "There are no saved activities with students marked as present."
    End Select
    
    MsgBox FullErrorString
    GoTo Footer
    
NewWorkbook:
    'Make a new workbook
    Set OldBook = ThisWorkbook
    Set NewBook = MakeNewBook(RecordsSheet, ReportSheet, RosterSheet, , "SharePoint")
    
    'Pull in center from the CoverSheet, date, and time
    CenterString = CoverSheet.Range("A:A").Find("Center", , xlValues, xlWhole).Offset(0, 1).Value
    SubDate = Format(Date, "yyyy-mm-dd")
    SubTime = Format(Time, "hh-nn AM/PM")

    'Make a file name using the center, date, and time of submission
    FileName = CenterString & " " & SubDate & "." & SubTime & ".xlsm"
    
    'The address where the new book will be save in SharePoint
    SpPath = "https://uwnetid.sharepoint.com/sites/partner_university_portal/Data%20Portal/Report%20Submissions/"
    
    'Upload
    NewBook.SaveAs FileName:=SpPath & "/" & FileName, FileFormat:=xlOpenXMLWorkbookMacroEnabled
    NewBook.Close SaveChanges:=False
    
    MsgBox ("Submitted to SharePoint")
    
Footer:
    Call ResetProtection

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True

End Sub

Sub CoverImportButton()
'Pull in roster and attendence from a previous version of the workbook

    Dim OldBook As Workbook
    Dim NewBook As Workbook
    Dim OldRosterSheet As Worksheet
    Dim NewRosterSheet As Worksheet
    Dim OldRecordsSheet As Worksheet
    Dim NewRecordsSheet As Worksheet
    Dim CopyRange As Range
    Dim PasteRange As Range
    Dim c As Range
    Dim d As Range
    Dim OldVersion As String
    Dim NewVersion As String
    Dim OldBookFilePath As Variant
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    
    Set NewBook = ActiveWorkbook
    
    'Promp to select the workbook to import
    OldBookFilePath = Application.GetOpenFilename("Excel Files (*.xlsm*), *xlsm*", , "Select the file to import") 'Returns FALSE if nothing is selected
    
    If OldBookFilePath = False Then
        GoTo Footer
    End If
    
    Workbooks.Open (OldBookFilePath)
    On Error Resume Next

    'Make sure it's a valid workbook
    Set OldBook = ActiveWorkbook

    If IsError(CheckRecords(OldBook.Worksheets("Records Page"))) Then
        MsgBox ("It looks like you have selected an incompatible workbook. Please try again.")
        
        GoTo CloseBook
    End If
    
    'This needs a bit of work so it can be generalized to the Term workbook
    If CheckRecords(OldBook.Worksheets("Records Page")) <> 1 Or CheckTable(OldBook.Worksheets("Roster Page")) > 2 Then
        MsgBox ("The file you choose must have both students and activities saved.")
        
        GoTo CloseBook
    End If
    
    'Check that it's the term or weekly version of the workbook
    If InStr(NewBook.Worksheets("Cover Page").Range("A1").Value, "Weekly") > 0 Then
        NewVersion = "Weekly"
    Else
        NewVersion = "Term"
    End If
    
    If InStr(OldBook.Worksheets("Cover Page").Range("A1").Value, "Weekly") > 0 Then
        OldVersion = "Weekly"
    Else
        OldVersion = "Term"
    End If
    
    If Not OldVersion = NewVersion Then
        MsgBox ("It looks like you have selected the " & OldVersion & " reporting workbook." & _
            vbCr & "Please select the " & NewVersion & " reporting workbook")
        OldBook.Close
        
        GoTo CloseBook
    End If

    'Get rid of anything currently on the Roster Sheet
    Set OldRosterSheet = OldBook.Worksheets("Roster Page")
    Set NewRosterSheet = NewBook.Worksheets("Roster Page")
    
    If CheckTable(NewRosterSheet) > 2 Then
        NewRosterSheet.ListObjects(1).DataBodyRange.ClearContents
    End If

    'Copy over the roster and parse it
    Set CopyRange = FindTableRange(OldRosterSheet)
    
    If CopyRange Is Nothing Then 'If there is nothing on the roster sheet. This shouldn't happen
        MsgBox ("There are no students on the selected file's Roster Page")
        GoTo CloseBook
    End If

    'Reset error handling
    On Error GoTo 0

    Set PasteRange = NewRosterSheet.Range(CopyRange.Address)
    
    PasteRange.Value = CopyRange.Value
    NewBook.Activate
    Call RosterParseButton
    
    'Clear the RecordsSheet before copying over information. This is to preserve the order of students
    Set OldRecordsSheet = OldBook.Worksheets("Records Page")
    Set NewRecordsSheet = NewBook.Worksheets("Records Page")
    
    Call ClearSheet(NewRecordsSheet)

    Set c = OldRecordsSheet.Range("A:A").Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    Set d = OldRecordsSheet.Range("1:1").Find("*", SearchOrder:=xlByColumns, SearchDirection:=xlPrevious)
    Set CopyRange = OldRecordsSheet.Range(c, d)
    Set PasteRange = NewRecordsSheet.Range(CopyRange.Address)

    PasteRange.Value = CopyRange.Value

    'Tabulate
    Call ReportTabulateAllButton
    MsgBox ("Import complete")


CloseBook:
    OldBook.Close SaveChanges:=False

Footer:

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True

End Sub
