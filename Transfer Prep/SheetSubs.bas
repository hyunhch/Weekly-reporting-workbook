Attribute VB_Name = "SheetSubs"
Option Explicit

Sub DateValidation(NewSheet As Worksheet, DateRange As Range)

    Call UnprotectCheck(NewSheet)

    With DateRange.Validation
        .Delete
        .Add Type:=xlValidateDate, AlertStyle:=xlValidAlertStop, Operator:=xlGreaterEqual, Formula1:="1/1/1990"
        .IgnoreBlank = True
        .InputTitle = ""
        .ErrorTitle = "Error"
        .ErrorMessage = "Please enter in a valid date"
        .ShowInput = True
        .ShowError = True
    End With
    
    Call ResetProtection

End Sub

Sub CenterDropdown(NewSheet As Worksheet, CenterRange As Range)
'Make a dropdown list with center names in the indicated cell

    Call UnprotectCheck(NewSheet)

    With CenterRange.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=CentersList"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = "Error"
        .InputMessage = ""
        .ErrorMessage = "Please choose from the drop-down list"
        .ShowInput = True
        .ShowError = True
    End With
    
    Call ResetProtection
    
End Sub

Sub ClearSheet(DelStart As Range, Repull As Long, TargetSheet As Worksheet)
'Repull = 1 avoids warning message

    Dim DelRange As Range
    Dim ClearAll As Long
    Dim OldTable As ListObject

    Set DelRange = TargetSheet.Range(Cells(DelStart.Row, DelStart.Column).Address, Cells(TargetSheet.Rows.Count, TargetSheet.Columns.Count).Address)
   
    If Repull <> 1 Then
            ClearAll = MsgBox("Are you sure you want to clear all content?" & vbCrLf & "This cannot be undone.", vbQuestion + vbYesNo + vbDefaultButton2, "")
    Else
        ClearAll = vbYes
    End If
 
    If ClearAll = vbYes Then
        For Each OldTable In TargetSheet.ListObjects
            OldTable.Unlist
        Next OldTable
        
        With DelRange
            .ClearContents
            .ClearFormats
            .Validation.Delete
        End With
    End If
    
End Sub

Sub UnprotectCheck(TargetSheet As Worksheet)
'Checks if a sheet is protected and unprotects
'Used to avoid trying to unprotect an already unprotected sheet

    If TargetSheet.ProtectContents = True Then
        TargetSheet.Unprotect
    End If

End Sub

Sub ResetProtection()
'Reset all sheet protections
    
    Dim RosterSheet As Worksheet
    Dim ReportSheet As Worksheet
    Dim CoverSheet As Worksheet
    Dim ChangeSheet As Worksheet
    
    Set RosterSheet = Worksheets("Roster Page")
    Set ReportSheet = Worksheets("Report Page")
    Set CoverSheet = Worksheets("Cover Page")
    Set ChangeSheet = Worksheets("Change Log")

    RosterSheet.Protect , userinterfaceonly:=True, AllowSorting:=True, AllowFiltering:=True, AllowFormattingColumns:=True
    ReportSheet.Protect , userinterfaceonly:=True, AllowSorting:=True, AllowFiltering:=True, AllowFormattingColumns:=True
    CoverSheet.Protect , userinterfaceonly:=True
    ChangeSheet.Protect , userinterfaceonly:=True

    'Lock/Unlock areas
    CoverSheet.Range("B3:B5").Locked = False
    
    RosterSheet.Cells.Locked = False
    RosterSheet.Range("A1:A5").EntireRow.Locked = True
    
    'Lock the entire page
    ReportSheet.Cells.Locked = True
    
End Sub

Sub NewActivitySheet(PracticeString As String, DateValue As Date, LabelString As String, DescriptionString As String, Optional HowMany As String)
'To be called from the userforms

    'Create a sheet at the end
    Dim NewSheet As Worksheet
    
    With ActiveWorkbook
        Set NewSheet = .Sheets.Add(After:=.Sheets(.Sheets.Count))
    End With
    
    'Add information from the userform
    Dim PracticeRange As Range
    Dim DateRange As Range
    Dim DescriptionRange As Range
    Dim LabelRange As Range
    
    With NewSheet
        Set PracticeRange = .Range("B1")
        Set DateRange = .Range("B3")
        Set DescriptionRange = .Range("B4")
        Set LabelRange = .Range("H1")
        
        PracticeRange.Value = PracticeString
        DateRange.Value = DateValue
        LabelRange.Value = LabelString
        DescriptionRange.Value = DescriptionString
    End With
    
    'Populate the new sheet
    Dim NewTableStart As Range
    Dim NewBoxRange As Range
    
    Set NewTableStart = NewSheet.Range("A6")
    Call PopulateSheet(NewSheet)
    Call CopySelected(NewSheet, HowMany)
    Call TableCreate(NewSheet, NewTableStart) 'Not passing a table name
    Call DateValidation(NewSheet, DateRange)

    Set NewBoxRange = NewSheet.ListObjects(1).ListColumns("Select").DataBodyRange
    Call AddMarlettBox(NewBoxRange, NewSheet)
    
    'Fit the first two columns
    NewSheet.Columns("A:A").AutoFit
    NewSheet.Range("B3").Columns.AutoFit
    
    'Apply protection to the first five rows
    NewSheet.Protect , userinterfaceonly:=True, AllowSorting:=True, AllowFiltering:=True, AllowFormattingColumns:=True
    NewSheet.Cells.Locked = False
    NewSheet.Range("A1:A5").EntireRow.Locked = True
    NewSheet.Range("B3:B4").Locked = False 'Allow the date and decription to be editable
    
    NewSheet.Activate
    
End Sub

Sub PopulateSheet(ActivitySheet As Worksheet)
'When a new sheet is created through a userform. Static text and formatting

    'Unprotect. This shouldn't ever be needed
    If ActivitySheet.ProtectContents = True Then
        ActivitySheet.Unprotect
    End If
    
    'Add text
    Dim TextArray() As String
    Dim TextRange As Range
    Dim LabelRange As Range
    
    Set TextRange = ActivitySheet.Range("A1:A4")
    TextArray = Split("Practice;Category;Date;Description", ";")
    TextRange.Value = Application.Transpose(TextArray)
    
    Set LabelRange = ActivitySheet.Range("G1")
    LabelRange.Value = "Label"
    
    'Find and add Practice Category
    Dim RefSheet As Worksheet
    Dim PracticeRange As Range
    
    Set RefSheet = Worksheets("Ref Tables")
    Set PracticeRange = ActivitySheet.Range("B1")
    PracticeRange.Offset(1, 0).Value = RefSheet.Range("B:B").Find(PracticeRange.Value, , xlValues, xlWhole).Offset(0, -1).Value
    
    'Format text and cells
    Dim c As Range

    For Each c In TextRange
        Set c = Union(c, c.Offset(0, 1))
        c.Borders(xlEdgeBottom).LineStyle = xlContinuous
        c.Borders(xlEdgeBottom).Weight = xlMedium
        c.Font.Bold = True
        c.WrapText = False
    Next c
    
    With Range(LabelRange, LabelRange.Offset(0, 1))
        .Cells.Font.Bold = True
        .WrapText = False
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Weight = xlMedium
    End With
    
    'Add Buttons
    Dim NewButton As Button
    Dim NewButtonRange As Range
    
    'Select All
    Set NewButtonRange = ActivitySheet.Range("A5:B5")
    Set NewButton = ActivitySheet.Buttons.Add(NewButtonRange.Left, NewButtonRange.Top, _
        NewButtonRange.Width, NewButtonRange.Height)
    
    With NewButton
        .OnAction = "SelectAllButton"
        .Caption = "Select All"
    End With

    'Delete Row
    Set NewButtonRange = ActivitySheet.Range("H5:I5")
    Set NewButton = ActivitySheet.Buttons.Add(NewButtonRange.Left, NewButtonRange.Top, _
        NewButtonRange.Width, NewButtonRange.Height)
    
    With NewButton
        .OnAction = "RemoveSelectedButton"
        .Caption = "Delete Row"
    End With

    'Delete Sheet button
    Set NewButtonRange = ActivitySheet.Range("H3:I3")
    Set NewButton = ActivitySheet.Buttons.Add(NewButtonRange.Left, NewButtonRange.Top, _
        NewButtonRange.Width, NewButtonRange.Height)
    
    With NewButton
        .OnAction = "DeleteActivityButton"
        .Caption = "Delete Activity"
    End With
    
    'Save Activity button
    Set NewButtonRange = ActivitySheet.Range("C5:D5")
    Set NewButton = ActivitySheet.Buttons.Add(NewButtonRange.Left, NewButtonRange.Top, _
        NewButtonRange.Width, NewButtonRange.Height)
    
    With NewButton
        .OnAction = "SaveActivityButton"
        .Caption = "Save Activity"
    End With
    
    'Close Activity button
    Set NewButtonRange = ActivitySheet.Range("E5:F5")
    Set NewButton = ActivitySheet.Buttons.Add(NewButtonRange.Left, NewButtonRange.Top, _
        NewButtonRange.Width, NewButtonRange.Height)
    
    With NewButton
        .OnAction = "CloseActivityButton"
        .Caption = "Close Activity"
    End With
    
Footer:

End Sub

