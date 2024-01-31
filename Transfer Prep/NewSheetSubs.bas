Attribute VB_Name = "NewSheetSubs"
Option Explicit

Sub CopySelectedStudents(NewSheet As Worksheet, PasteRange As Range, LRow As Long, LCol As Long, OldTableStart As Range)
'Only grab students with a checkmark

    Dim CoverSheet As Worksheet
    Dim i As Long
    Dim j As Long
    Dim CopyRange As Range
    
    Set CoverSheet = Worksheets("Cover Page")
    'Sometimes the 1st column will be renamed and I'm not sure why
    If OldTableStart.Value <> "Select" Then
        OldTableStart.Value = "Select"
    End If
    
    'Grab the header and every row that's been checked
    j = 0
    For i = OldTableStart.Row To LRow
        If CoverSheet.Cells(i, 1) <> "" Then
            Set CopyRange = CoverSheet.Range(Cells(i, 1).Address, Cells(i, LCol).Address)
            CopyRange.Copy
            PasteRange.Offset(j, 0).PasteSpecial xlPasteValues
            j = j + 1
        End If
    Next i
    
    'Create a table
    Dim NewLRow As Long
    Dim NewLCol As Long
    Dim NewTableRange As Range
    Dim NewTable As ListObject

    NewLRow = NewSheet.Cells(Rows.Count, 2).End(xlUp).Row
    NewLCol = NewSheet.Cells(PasteRange.Row, Columns.Count).End(xlToLeft).Column
    Set NewTableRange = NewSheet.Range(Cells(PasteRange.Row, 1).Address, Cells(NewLRow, NewLCol).Address)
    
    NewTableRange.FormatConditions.Delete
    Set NewTable = NewSheet.ListObjects.Add(SourceType:=xlSrcRange, Source:=NewTableRange, _
        xlListObjectHasHeaders:=xlYes)
    NewSheet.ListObjects(1).ShowTableStyleRowStripes = False
    
    'Formatting
    Call TableFormat(NewSheet.ListObjects(1), NewSheet)
    NewTableRange.Columns.AutoFit
    
    'Add Marlett boxes and select buttons
    Dim BoxRange As Range
    Dim SelectAllRange As Range
    
    Set BoxRange = NewSheet.ListObjects(1).ListColumns("Select").DataBodyRange
    Set SelectAllRange = PasteRange.Offset(-1, 0)
    
    Call AddMarlettBox(BoxRange, NewSheet)
    Call AddSelectAll(SelectAllRange, NewSheet)

End Sub

Sub PopulateSheet(NewSheet As Worksheet, NewCenterName As String, NewDirectorName As String)
'Pull the selected students and activity into a new sheet
'Rename sheets, put in buttons for updating the pulled student list and tabulating

    Dim RefSheet As Worksheet
    Dim DateRange As Range
    Dim DropDownRange As Range
    
    Set RefSheet = Worksheets("Ref Tables")
    Set DateRange = Range("B3")
    Set DropDownRange = Range("F1")

    'Copy the center. Name and date can be entered
    With NewSheet
        .Range("A1").Value = "Name"
        .Range("A2").Value = "Center"
        .Range("A3").Value = "Date"
        .Range("B1").Value = NewDirectorName
        .Range("B2").Value = NewCenterName
        .Range("E1").Value = "Practice"
        .Range("E2").Value = "Logic Model Category"
        .Range("E3").Value = "Description"
    End With
    Call DateValidation(NewSheet, DateRange)

    'Put in the drop down menu
    Call ActivityDropDown(NewSheet, DropDownRange)

    'Insert tabulate button
    Dim NewButton As Button
    Dim NewButtonRange As Range
    Dim NewSheetString As String
    
    NewSheetString = "NewSheet"
    Set NewButtonRange = NewSheet.Range("G5:H5")
    Set NewButton = NewSheet.Buttons.Add(NewButtonRange.Left, NewButtonRange.Top, _
        NewButtonRange.Width, NewButtonRange.Height)
    
    With NewButton
        .OnAction = "TabulateCaller"
        .Caption = "Tabulate Activity"
    End With

    'Insert Delete Sheet button
    Set NewButtonRange = NewSheet.Range("D5:E5")
    Set NewButton = NewSheet.Buttons.Add(NewButtonRange.Left, NewButtonRange.Top, _
        NewButtonRange.Width, NewButtonRange.Height)
    
    With NewButton
        .OnAction = "RemoveSheet"
        .Caption = "Delete Sheet"
    End With
    
    'Insert Delete Row button
    Set NewButtonRange = NewSheet.Range("B5:C5")
    Set NewButton = NewSheet.Buttons.Add(NewButtonRange.Left, NewButtonRange.Top, _
        NewButtonRange.Width, NewButtonRange.Height)
    
    With NewButton
        .OnAction = "RemoveSelected"
        .Caption = "Delete Row"
    End With
    
    'Formatting
    Dim c As Range
    
    With NewSheet
        .Range("B3").HorizontalAlignment = xlLeft
        .Range("A1:A3").Font.Bold = True
        .Range("A1:A3").HorizontalAlignment = xlRight
        .Range("E1:E3").Font.Bold = True
        .Range("E1:E3").HorizontalAlignment = xlRight
        .Cells.WrapText = False
    End With
    
    For Each c In NewSheet.Range("B1:B3")
        With c.Borders(xlEdgeBottom)
            .LineStyle = xlSingle
            .Weight = xlThick
        End With
    Next
    
End Sub

Sub RenameSheets()
'To be called when a sheet is created or deleted

    Dim NumSheets As Long
    
    NumSheets = ThisWorkbook.Sheets.Count
    
    'There should always be at least five sheets
    'Ref tables, Change log, roster sheet, report sheet, cover sheet
    If NumSheets < 5 Then
        MsgBox ("Something has gone wrong with this Excel workbook." & vbNewLine & _
            "Please try downloading a fresh copy.")
        Exit Sub
    ElseIf NumSheets = 5 Then
        Exit Sub
    End If
    
    Dim i As Long
    i = 1
    
    Do While Sheets.Count > 5 And i <= Sheets.Count - 5 'So we don't go out of range
        Sheets(Sheets(5).Index + i).Name = "Activity " & i
        i = i + 1
    Loop
    
End Sub

Sub ActivityDropDown(NewSheet As Worksheet, DropDownRange As Range)
'Create a dropdown menu and autopoulate the indicated cell

    With DropDownRange.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=ActivitiesList"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = "Error"
        .InputMessage = ""
        .ErrorMessage = "Please choose from the drop-down list"
        .ShowInput = True
        .ShowError = True
    End With

End Sub

Sub DateValidation(NewSheet As Worksheet, DateRange As Range)

    With DateRange.Validation
        .Delete
        .Add Type:=xlValidateDate, Operator:=xlLessEqual, AlertStyle:=xlValidAlertStop, Formula1:=Date
        .IgnoreBlank = True
        .InputTitle = ""
        .ErrorTitle = "Error"
        .ErrorMessage = "Please enter in a valid date"
        .ShowInput = True
        .ShowError = True
    End With

End Sub

Sub CenterDropdown(NewSheet As Worksheet, CenterRange As Range)
'Make a dropdown list with center names in the indicated cell

    With NewSheet.Range("CenterRange").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=CenterNames"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = "Error"
        .InputMessage = ""
        .ErrorMessage = "Please choose from the drop-down list"
        .ShowInput = True
        .ShowError = True
    End With
    
End Sub

