Attribute VB_Name = "ForDebug"
Option Explicit

Sub ScreenUpdating()

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True

End Sub

Sub testing()
End Sub

Sub RosterButtons()
'Add Buttons

    Dim RosterSheet As Worksheet
    Dim NewButton As Button
    Dim NewButtonRange As Range
    
    Set RosterSheet = Worksheets("Roster Page")
    
    'Select All
    Set NewButtonRange = RosterSheet.Range("A5:B5")
    Set NewButton = RosterSheet.Buttons.Add(NewButtonRange.Left, NewButtonRange.Top, _
        NewButtonRange.Width, NewButtonRange.Height)
    
    With NewButton
        .OnAction = "SelectAllButton"
        .Caption = "Select All"
    End With

    'Delete Row
    Set NewButtonRange = RosterSheet.Range("D5:E5")
    Set NewButton = RosterSheet.Buttons.Add(NewButtonRange.Left, NewButtonRange.Top, _
        NewButtonRange.Width, NewButtonRange.Height)
    
    With NewButton
        .OnAction = "RemoveSelectedButton"
        .Caption = "Delete Row"
    End With
    
    'New activity
    Set NewButtonRange = RosterSheet.Range("H5:I5")
    Set NewButton = RosterSheet.Buttons.Add(NewButtonRange.Left, NewButtonRange.Top, _
        NewButtonRange.Width, NewButtonRange.Height)
    
    With NewButton
        .OnAction = "OpenNewActivityButton"
        .Caption = "New Activity"
    End With
    
    'Load activity
    Set NewButtonRange = RosterSheet.Range("H4:I4")
    Set NewButton = RosterSheet.Buttons.Add(NewButtonRange.Left, NewButtonRange.Top, _
        NewButtonRange.Width, NewButtonRange.Height)
    
    With NewButton
        .OnAction = "OpenLoadActivityButton"
        .Caption = "Load Activity"
    End With
    
    'Add students
    Set NewButtonRange = RosterSheet.Range("H1:I1")
    Set NewButton = RosterSheet.Buttons.Add(NewButtonRange.Left, NewButtonRange.Top, _
        NewButtonRange.Width, NewButtonRange.Height)
    
    With NewButton
        .OnAction = "AddSelectedStudentsButton"
        .Caption = "Add to Activity"
    End With
    
    'Read roster
    Set NewButtonRange = RosterSheet.Range("A1:B2")
    Set NewButton = RosterSheet.Buttons.Add(NewButtonRange.Left, NewButtonRange.Top, _
        NewButtonRange.Width, NewButtonRange.Height)
    
    With NewButton
        .OnAction = "ReadRosterButton"
        .Caption = "Parse Roster"
    End With
    
    'Clear roster
    Set NewButtonRange = RosterSheet.Range("D1:E1")
    Set NewButton = RosterSheet.Buttons.Add(NewButtonRange.Left, NewButtonRange.Top, _
        NewButtonRange.Width, NewButtonRange.Height)
    
    With NewButton
        .OnAction = "ClearRosterButton"
        .Caption = "Clear Roster"
    End With
    
End Sub

Sub ReportButtons()
'Add buttons to the Report page

    Dim ReportSheet As Worksheet
    Dim NewButton As Button
    Dim NewButtonRange As Range
    
    Set ReportSheet = Worksheets("Report Page")
    
    'Select All
    Set NewButtonRange = ReportSheet.Range("A5:B5")
    Set NewButton = ReportSheet.Buttons.Add(NewButtonRange.Left, NewButtonRange.Top, _
        NewButtonRange.Width, NewButtonRange.Height)
    
    With NewButton
        .OnAction = "SelectAllButton"
        .Caption = "Select All"
    End With
    
    'Pull Totals
    Set NewButtonRange = ReportSheet.Range("A1:B2")
    Set NewButton = ReportSheet.Buttons.Add(NewButtonRange.Left, NewButtonRange.Top, _
        NewButtonRange.Width, NewButtonRange.Height)
    
    With NewButton
        .OnAction = "PullReportTotalsButton"
        .Caption = "Tabulate Totals"
    End With
    
    'Clear the roster
    Set NewButtonRange = ReportSheet.Range("G5")
    Set NewButton = ReportSheet.Buttons.Add(NewButtonRange.Left, NewButtonRange.Top, _
        NewButtonRange.Width, NewButtonRange.Height)
    
    With NewButton
        .OnAction = "ClearReportButton"
        .Caption = "Clear Report"
    End With
    
    'Tabulate activities
    Set NewButtonRange = ReportSheet.Range("D1:E2")
    Set NewButton = ReportSheet.Buttons.Add(NewButtonRange.Left, NewButtonRange.Top, _
        NewButtonRange.Width, NewButtonRange.Height)
    
    With NewButton
        .OnAction = "TabulateButton"
        .Caption = "Tabulate Activities"
    End With
    
    'Remove row
    Set NewButtonRange = ReportSheet.Range("D5:E5")
    Set NewButton = ReportSheet.Buttons.Add(NewButtonRange.Left, NewButtonRange.Top, _
        NewButtonRange.Width, NewButtonRange.Height)
    
    With NewButton
        .OnAction = "RemoveSelectedButton"
        .Caption = "Delete Row"
    End With
    
End Sub

Sub CoverSheetButtons()
'Formatting and buttons for the cover sheet

    Dim CoverSheet As Worksheet
    Dim DateRange As Range
    Dim CenterRange As Range
    Dim NewButton As Button
    Dim NewButtonRange As Range
    
    Set CoverSheet = Worksheets("Cover Page")
    Set DateRange = CoverSheet.Range("B4")
    Set CenterRange = CoverSheet.Range("B5")
    
    'Date validation and a dropdown menu for the center
    Call DateValidation(CoverSheet, DateRange)
    Call CenterDropdown(CoverSheet, CenterRange)
    
    'Submit button
    Set NewButtonRange = CoverSheet.Range("D1:F2")
    Set NewButton = CoverSheet.Buttons.Add(NewButtonRange.Left, NewButtonRange.Top, _
        NewButtonRange.Width, NewButtonRange.Height)
    
    With NewButton
        .OnAction = "SharePointExport"
        .Caption = "Submit to SharePoint"
    End With
    
    'Save button
    Set NewButtonRange = CoverSheet.Range("D4:F5")
    Set NewButton = CoverSheet.Buttons.Add(NewButtonRange.Left, NewButtonRange.Top, _
        NewButtonRange.Width, NewButtonRange.Height)
    
    With NewButton
        .OnAction = "LocalSave"
        .Caption = "Save a Copy"
    End With
    
End Sub

Sub BreakExternalLinks()
'PURPOSE: Breaks all external links that would show up in Excel's "Edit Links" Dialog Box
'SOURCE: www.TheSpreadsheetGuru.com/The-Code-Vault

Dim ExternalLinksArray As Variant
Dim wb As Workbook
Dim X As Long

Set wb = ActiveWorkbook

'Create an Array of all External Links stored in Workbook
  ExternalLinksArray = wb.LinkSources(Type:=xlLinkTypeExcelLinks)

'if the array is not empty the loop Through each External Link in ActiveWorkbook and Break it
 If IsEmpty(ExternalLinksArray) = False Then
     For X = 1 To UBound(ExternalLinksArray)
        wb.BreakLink Name:=ExternalLinksArray(X), Type:=xlLinkTypeExcelLinks
      Next X
End If

End Sub


Sub SaveTest()
'Make sure the save button will work

    Dim PasteBook As Workbook
    Dim CoverSheet As Worksheet
    Dim RosterSheet As Worksheet
    Dim ReportSheet As Worksheet
    Dim RecordsSheet As Worksheet
    Dim SheetNames() As String
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    
    Set CoverSheet = Worksheets("Cover Page")
    Set RosterSheet = Worksheets("Roster Page")
    Set ReportSheet = Worksheets("Report Page")
    Set RecordsSheet = Worksheets("Records Page")
    SheetNames = Split("Detailed Attendance;Attendance;Report;Cover", ";")
    Set PasteBook = Workbooks.Add
    
    Call NewSaveBook(PasteBook, CoverSheet, RosterSheet, ReportSheet, RecordsSheet, SheetNames)
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True

End Sub
Public NameString As String


