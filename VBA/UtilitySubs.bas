Attribute VB_Name = "UtilitySubs"
Option Explicit

Function BuildRange(NewCell As Range, Optional OldRange As Range) As Range
'A function for building ranges cell by cell
'This may be slower

    If OldRange Is Nothing Then
        Set BuildRange = NewCell
    Else
        Set BuildRange = Union(OldRange, NewCell)
    End If

Footer:

End Function

Sub CenterDropdown(TargetSheet As Worksheet, CenterRange As Range)
'Make a dropdown list with center names in the indicated cell

    Call UnprotectSheet(TargetSheet)

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
    
End Sub

Sub ClearSheet(TargetSheet As Worksheet, Optional ShowWarning As String, Optional DelStart As Range)
'Clears everything on a sheet and deletes tables
'Passing "Warn" prompts a confirmation for deletion
'Passing a range deletes everything to the right and below

    Dim DelRange As Range
    Dim DelConfirm As Long
    Dim DelTable As ListObject
    
    'Warning prompt
    If ShowWarning = "Yes" Then
        DelConfirm = MsgBox("Are you sure you want to clear all content?" & vbCrLf & _
            "This cannot be undone.", vbQuestion + vbYesNo + vbDefaultButton2, "")
    Else
        DelConfirm = vbYes
    End If

    'If DelRange was passed, only delete from that point to the right and down
    If Not DelStart Is Nothing Then
        Set DelRange = TargetSheet.Range(DelStart, Cells(TargetSheet.Rows.Count, TargetSheet.Columns.Count).Address)
    Else
        Set DelRange = TargetSheet.Cells
    End If

    'Delete content and formats
    If DelConfirm = vbYes Then
        Call RemoveTable(TargetSheet)
        
        With DelRange
            .ClearContents
            .ClearFormats
            .Validation.Delete
        End With
    End If

End Sub

Sub DateValidation(TargetSheet As Worksheet, DateRange As Range)
'Date greater than 1990

    Call UnprotectSheet(TargetSheet)

    With DateRange.Validation
        .Delete
        .Add Type:=xlValidateDate, AlertStyle:=xlValidAlertStop, Operator:=xlGreaterEqual, Formula1:="1/1/1990"
        .IgnoreBlank = True
        .InputTitle = ""
        .ErrorTitle = "Error"
        .ErrorMessage = "Please enter a date as mm/dd/yyyy"
        .ShowInput = True
        .ShowError = True
    End With

End Sub

Sub ResetProtection()
'Reset all sheet protections
    
    Dim ReportBook As Workbook
    Dim RosterSheet As Worksheet
    Dim ReportSheet As Worksheet
    Dim CoverSheet As Worksheet
    Dim ChangeSheet As Worksheet
    Dim ActivitySheet As Worksheet
    
    Set ReportBook = ActiveWorkbook
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
    
    'Lock the entire page besides the "Select: Column
    ReportSheet.Cells.Locked = True
    ReportSheet.Range("A:A").Locked = False
    ReportSheet.Range("A1:A5").EntireRow.Locked = True
    
    'All activity sheets
    For Each ActivitySheet In ReportBook.Sheets
        If ActivitySheet.Range("A1").Value = "Practice" Then
            ActivitySheet.Protect , userinterfaceonly:=True, AllowSorting:=True, AllowFiltering:=True, AllowFormattingColumns:=True
            ActivitySheet.Cells.Locked = False
            ActivitySheet.Range("A1:A5").EntireRow.Locked = True
            ActivitySheet.Range("B3:B4").Locked = False 'Allow the date and decription to be editable
        End If
    Next ActivitySheet
    
End Sub

Sub UnprotectSheet(TargetSheet As Worksheet)
'Checks if a sheet is protected and unprotects
'Used to avoid trying to unprotect an already unprotected sheet

    If TargetSheet.ProtectContents = True Then
        TargetSheet.Unprotect
    End If

End Sub


