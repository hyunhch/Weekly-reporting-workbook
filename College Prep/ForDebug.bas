Attribute VB_Name = "ForDebug"
Sub CoverPlaceButtons()
'To place some control buttons inside cells

    Dim CoverSheet As Worksheet
    Dim MyButton As Button
    Dim ButtonRange As Range
    
    Set CoverSheet = Worksheets("Cover Page")
    
    'Pull Roster button
    Set ButtonRange = CoverSheet.Range("D2:E3")
    Set MyButton = CoverSheet.Buttons.Add(ButtonRange.Left, ButtonRange.Top, ButtonRange.Width, ButtonRange.Height)
    With MyButton
        .OnAction = "PullRoster"
        .Caption = "Pull Roster"
    End With
    
    'Add students buttons
    Set ButtonRange = CoverSheet.Range("D5:E5")
    Set MyButton = CoverSheet.Buttons.Add(ButtonRange.Left, ButtonRange.Top, ButtonRange.Width, ButtonRange.Height)
    With MyButton
        .OnAction = "AddSelected"
        .Caption = "Add to Activity"
    End With
    
    'Delete row
    Set ButtonRange = CoverSheet.Range("B5:C5")
    Set MyButton = CoverSheet.Buttons.Add(ButtonRange.Left, ButtonRange.Top, ButtonRange.Width, ButtonRange.Height)
    With MyButton
        .OnAction = "RemoveSelected"
        .Caption = "Delete Row"
    End With

    'Tabulate
    Set ButtonRange = CoverSheet.Range("G2:H3")
    Set MyButton = CoverSheet.Buttons.Add(ButtonRange.Left, ButtonRange.Top, ButtonRange.Width, ButtonRange.Height)
    With MyButton
        .OnAction = "CompileActivities"
        .Caption = "Tabulate Activities"
    End With
    
    'Add sheet
    Set ButtonRange = CoverSheet.Range("G5:H5")
    Set MyButton = CoverSheet.Buttons.Add(ButtonRange.Left, ButtonRange.Top, ButtonRange.Width, ButtonRange.Height)
    With MyButton
        .OnAction = "AddSheet"
        .Caption = "New Activity"
    End With

End Sub

Sub RosterPlaceButtons()

    Dim RosterSheet As Worksheet
    Dim MyButton As Button
    Dim ButtonRange As Range
    
    Set RosterSheet = Worksheets("Roster Page")
    
    'Read roster button
    Set ButtonRange = RosterSheet.Range("A2:B3")
    Set MyButton = RosterSheet.Buttons.Add(ButtonRange.Left, ButtonRange.Top, ButtonRange.Width, ButtonRange.Height)
    With MyButton
        .OnAction = "ReadRoster"
        .Caption = "Read Roster"
    End With
    
    'Clear Roster button
    Set ButtonRange = RosterSheet.Range("C2:D3")
    Set MyButton = RosterSheet.Buttons.Add(ButtonRange.Left, ButtonRange.Top, ButtonRange.Width, ButtonRange.Height)
    With MyButton
        .OnAction = "RosterSheetClear"
        .Caption = "Clear Roster"
    End With

    'Delete Row button
    Set ButtonRange = RosterSheet.Range("B5:C5")
    Set MyButton = RosterSheet.Buttons.Add(ButtonRange.Left, ButtonRange.Top, ButtonRange.Width, ButtonRange.Height)
    With MyButton
        .OnAction = "RemoveSelected"
        .Caption = "Delete Row"
    End With

End Sub

Sub ReportPlaceButtons()

    Dim ReportSheet As Worksheet
    Dim MyButton As Button
    Dim ButtonRange As Range
    
    Set ReportSheet = Worksheets("Report Page")
    
    'Select all button
    Set ButtonRange = ReportSheet.Range("A5")
    Set MyButton = ReportSheet.Buttons.Add(ButtonRange.Left, ButtonRange.Top, ButtonRange.Width, ButtonRange.Height)
    With MyButton
        .OnAction = "SelectAll"
        .Caption = "Select All"
    End With

    'Delete Row button
    Set ButtonRange = ReportSheet.Range("B5:C5")
    Set MyButton = ReportSheet.Buttons.Add(ButtonRange.Left, ButtonRange.Top, ButtonRange.Width, ButtonRange.Height)
    With MyButton
        .OnAction = "RemoveSelected"
        .Caption = "Delete Row"
    End With
    
    'Clear entire report
    Set ButtonRange = ReportSheet.Range("A2:B3")
    Set MyButton = ReportSheet.Buttons.Add(ButtonRange.Left, ButtonRange.Top, ButtonRange.Width, ButtonRange.Height)
    With MyButton
        .OnAction = "ClearReport"
        .Caption = "Clear Report"
    End With
    
    'Pull totals
    Set ButtonRange = ReportSheet.Range("G5:H5")
    Set MyButton = ReportSheet.Buttons.Add(ButtonRange.Left, ButtonRange.Top, ButtonRange.Width, ButtonRange.Height)
    With MyButton
        .OnAction = "PullReportTotals"
        .Caption = "Pull Totals"
    End With
    
    'Clear totals
    Set ButtonRange = ReportSheet.Range("I5:J5")
    Set MyButton = ReportSheet.Buttons.Add(ButtonRange.Left, ButtonRange.Top, ButtonRange.Width, ButtonRange.Height)
    With MyButton
        .OnAction = "ClearReportTotals"
        .Caption = "Clear Totals"
    End With

    'Local Save
    Set ButtonRange = ReportSheet.Range("G2:H3")
    Set MyButton = ReportSheet.Buttons.Add(ButtonRange.Left, ButtonRange.Top, ButtonRange.Width, ButtonRange.Height)
    With MyButton
        .OnAction = "LocalSave"
        .Caption = "Save Report"
    End With

    'Export to SharePoint
    Set ButtonRange = ReportSheet.Range("I2:J3")
    Set MyButton = ReportSheet.Buttons.Add(ButtonRange.Left, ButtonRange.Top, ButtonRange.Width, ButtonRange.Height)
    With MyButton
        .OnAction = "SharePointExport"
        .Caption = "Export to SharePoint"
    End With

End Sub

