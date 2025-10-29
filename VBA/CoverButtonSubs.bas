Attribute VB_Name = "CoverButtonSubs"
Option Explicit

Sub CoverChooseProgramButton()
'Brings up the userform to choose a program. Button is deleted afterward

    ChooseProgramForm.Show

Footer:

End Sub

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
    ReDim TempArray(1 To 5)
        TempArray(1) = "Cover Page"
        TempArray(2) = "Report Page"
        TempArray(3) = "Roster Page"
        TempArray(4) = "Simple Attendance"
        TempArray(5) = "Detailed Attendance"

    Set OldBook = ThisWorkbook
    Set NewBook = ExportMakeBook(, TempArray)
        If NewBook Is Nothing Then
            MsgBox ("The submission could not be completed. Close and reopen this workbook and try again.")
            
            GoTo Footer
        End If
    
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
            NewBook.Close savechanges:=False
            GoTo Footer
        End If
        NewBook.SaveAs FileName:=LocalPath & "/" & FileName, FileFormat:=xlOpenXMLWorkbookMacroEnabled
        'ActiveWorkbook.Close SaveChanges:=False
    Else
        SaveName = Application.GetSaveAsFilename(LocalPath & "\" & FileName, "Excel Files (*.xlsm), *.xlsm")
        If SaveName = "False" Then
            NewBook.Close savechanges:=False
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
    ReDim TempArray(1 To 2)
        TempArray(1) = "Cover Page"
        TempArray(2) = "Report Page"
    
    Set OldBook = ThisWorkbook
    Set NewBook = ExportMakeBook(, TempArray)
        If NewBook Is Nothing Then
            MsgBox ("The submission could not be completed. Close and reopen this workbook and try again.")
            
            GoTo Footer
        End If
    
    'Pull in center from the CoverSheet, date, and time
    CenterString = CoverSheet.Range("A:A").Find("Center", , xlValues, xlWhole).Offset(0, 1).Value
    SubDate = Format(Date, "yyyy-mm-dd")
    SubTime = Format(Time, "hh-nn AM/PM")

    'Make a file name using the center, date, and time of submission
    FileName = CenterString & " " & SubDate & "." & SubTime & ".xlsm"
    
    'The address where the new book will be save in SharePoint
    SpPath = "https://uwnetid.sharepoint.com/sites/partner_university_portal/Data%20Portal/Center%20Files%20Data%20and%20Reports/Report%20Submissions/"
    
    'Upload
    NewBook.SaveAs FileName:=SpPath & "/" & FileName, FileFormat:=xlOpenXMLWorkbookMacroEnabled
    NewBook.Close savechanges:=False
    
    MsgBox ("Submitted to SharePoint")
    
Footer:
    Call ResetProtection

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True

End Sub

Sub CoverImportButton()
'Pull in roster and attendence from a previous version of the workbook
'*****Need to add measures for changes in the headers or practice/category spelling*****

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
    Dim OldTable As ListObject
    Dim NewTable As ListObject
    Dim OldBookFilePath As Variant
    Dim IsImported As Boolean
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    
    IsImported = False
    
    Set NewBook = ThisWorkbook
    
    'Promp to select the workbook to import
    OldBookFilePath = Application.GetOpenFilename("Excel Files (*.xlsm*), *xlsm*", , "Select the file to import") 'Returns FALSE if nothing is selected
        If OldBookFilePath = False Then
            GoTo Footer
        End If
    
    'Cear out everything in the new workbook before we open the old one
    Set NewRosterSheet = NewBook.Worksheets("Roster Page")
    Set NewRecordsSheet = NewBook.Worksheets("Records Page")
    
    Call RosterClear(NewRosterSheet) 'avoids prompt, clears the Records and Report as well
    'Clear the RecordsSheet before copying over information. This is to preserve the order of students
    
    'Open the old book
    Set OldBook = Workbooks.Open(OldBookFilePath)
    On Error Resume Next

    'Make sure it's a valid workbook
    'Set OldBook = ActiveWorkbook

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

    Set OldRosterSheet = OldBook.Worksheets("Roster Page")
    Set OldRecordsSheet = OldBook.Worksheets("Records Page")

    'Copy over the roster and parse it
    If Not CheckTable(OldRosterSheet) < 3 Then
        MsgBox ("There are no students on the selected file's Roster Page")
        
        GoTo CloseBook
    End If
    
    Set CopyRange = FindTableRange(OldRosterSheet)
        'If CopyRange Is Nothing Then 'If there is nothing on the roster sheet. This shouldn't happen
            'MsgBox ("There are no students on the selected file's Roster Page")
            'GoTo CloseBook
        'End If

    'Reset error handling
    On Error GoTo 0

    Set PasteRange = NewRosterSheet.Range(CopyRange.Address)
    
    PasteRange.Value = CopyRange.Value
    
    'Copy over the records
    Set CopyRange = FindRecordsRange(OldRecordsSheet)
    Set PasteRange = NewRecordsSheet.Range(CopyRange.Address)

    PasteRange.Value = CopyRange.Value
    
    IsImported = True

CloseBook:
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False

    Set OldRosterSheet = Nothing
    Set OldRecordsSheet = Nothing
    Set OldTable = Nothing
    Set CopyRange = Nothing

    OldBook.Saved = True
    OldBook.Close savechanges:=False
    
    If IsImported <> True Then
        GoTo Footer
    End If

    'Parse the roster and retabulate activities
    Call RosterParseButton 'This might cause problems if the old book is borked
    Call ReportTabulateAllButton
    
    MsgBox ("Import complete")

Footer:

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True

End Sub
