Attribute VB_Name = "CoverButtonSubs"
Option Explicit

Sub SharePointExport()
'reformat the data and export a new spreadsheet to SharePoint. Use a dynamic name with the center name and date
                
    Dim CopyBook As Workbook
    Dim PasteBook As Workbook
    Dim CoverSheet As Worksheet
    Dim ReportSheet As Worksheet
    Dim RosterSheet As Worksheet
    Dim RecordsSheet As Worksheet
    Dim CenterString As String
    Dim SubDate As String
    Dim SubTime As String
    Dim SpPath As String
    Dim FileName As String
    Dim SaveName As String
    Dim NewSheetNames() As String

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    
    'The sheets we need
    Set CoverSheet = Worksheets("Cover Page")
    Set ReportSheet = Worksheets("Report Page")
    Set RosterSheet = Worksheets("Roster Page")
    Set RecordsSheet = Worksheets("Records Page")

    'Submission information
    CenterString = CoverSheet.Range("B5").Value
    SubDate = Format(Date, "yyyy-mm-dd")
    SubTime = Format(Time, "hh-nn AM/PM")
    
    'Make sure information is filled out
    If ReadyToSave(CoverSheet, ReportSheet, RecordsSheet) = False Then
        GoTo Footer
    End If
        
    'Names of the new sheets we'll have
    NewSheetNames = Split("Detailed Attendance;Attendance;Report;Cover", ";")
        
    'Create a new workbook and pass to populate it
    Set CopyBook = ActiveWorkbook
    Set PasteBook = Workbooks.Add
    
    If NewSaveBook(PasteBook, CoverSheet, RosterSheet, ReportSheet, RecordsSheet, NewSheetNames, "SharePoint") = False Then
        'MsgBox ("Something has gone wrong. Please contact the State Office for support.")
        GoTo Footer
    End If

    'Create a file name based on the center and date of submission. The center *must* be filled
    'Path to the folder these will be saved in
    SpPath = "https://uwnetid.sharepoint.com/sites/partner_university_portal/Data%20Portal/Report%20Submissions/"
    FileName = CenterString & " " & SubDate & "." & SubTime & ".xlsm"
    
    PasteBook.SaveAs FileName:=SpPath & "/" & FileName, FileFormat:=xlOpenXMLWorkbookMacroEnabled
    ActiveWorkbook.Close SaveChanges:=False
    
    MsgBox ("Submitted to SharePoint")
    
Footer:
    Call ResetProtection

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    
End Sub

Sub LocalSave()
'Includes additional data for center directors

    Dim CopyBook As Workbook
    Dim PasteBook As Workbook
    Dim CoverSheet As Worksheet
    Dim ReportSheet As Worksheet
    Dim RosterSheet As Worksheet
    Dim RecordsSheet As Worksheet
    Dim CenterString As String
    Dim SubDate As String
    Dim SubTime As String
    Dim LocalPath As String
    Dim FileName As String
    Dim SaveName As String
    Dim NewSheetNames() As String
    Dim SaveSuccessful As Boolean

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    
    SaveSuccessful = False
    
    'The sheets we need
    Set CoverSheet = Worksheets("Cover Page")
    Set ReportSheet = Worksheets("Report Page")
    Set RosterSheet = Worksheets("Roster Page")
    Set RecordsSheet = Worksheets("Records Page")

    'Submission information
    CenterString = CoverSheet.Range("B5").Value
    SubDate = Format(Date, "yyyy-mm-dd")
    SubTime = Format(Time, "hh-nn AM/PM")
    
    'Make sure information is filled out
    If ReadyToSave(CoverSheet, ReportSheet, RecordsSheet) = False Then
        GoTo Footer
    End If
        
    'Names of the new sheets we'll have
    NewSheetNames = Split("Detailed Attendance;Attendance;Report;Cover", ";")
        
    'Create a new workbook and pass to populate it
    Set CopyBook = ActiveWorkbook
    Set PasteBook = Workbooks.Add
    
    If NewSaveBook(PasteBook, CoverSheet, RosterSheet, ReportSheet, RecordsSheet, NewSheetNames, "Local") = False Then
        'MsgBox ("Something has gone wrong. Please contact the State Office for support.")
        GoTo Footer
    End If
    
    'Create a file name based on the center and date of submission. The center *must* be filled
    'Path to the folder these will be saved in
    LocalPath = GetLocalPath(ThisWorkbook.path)
    FileName = CenterString & " " & SubDate & "." & SubTime & ".xlsm"
    
    'For Win and Mac
    If Application.OperatingSystem Like "*Mac*" Then
        SaveName = Application.GetSaveAsFilename(LocalPath & "\" & FileName, "Excel Files (*.xlsm), *.xlsm")
        If SaveName = "False" Then
            ActiveWorkbook.Close SaveChanges:=False
            GoTo Footer
        End If
        ActiveWorkbook.SaveAs FileName:=LocalPath & "/" & FileName, FileFormat:=xlOpenXMLWorkbookMacroEnabled
        ActiveWorkbook.Close SaveChanges:=False
    Else
        SaveName = Application.GetSaveAsFilename(LocalPath & "\" & FileName, "Excel Files (*.xlsm), *.xlsm")
        If SaveName = "False" Then
            ActiveWorkbook.Close SaveChanges:=False
            GoTo Footer
        End If
        ActiveWorkbook.SaveAs FileName:=SaveName, FileFormat:=xlOpenXMLWorkbookMacroEnabled
        'ActiveWorkbook.Close SaveChanges:=False
    End If
    
'Everything worked
    SaveSuccessful = True
    
Footer:
    CoverSheet.Activate
    Call ResetProtection
    
    If SaveSuccessful = True Then
        PasteBook.Activate
    End If

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    
End Sub
