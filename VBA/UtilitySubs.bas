Attribute VB_Name = "UtilitySubs"
Option Explicit

Sub AddMarlettBox(BoxHere As Range)
'Doing this instead of actual checkboxes to deal with sorting issues
'This only changes the font of a range to Marlett
    
    Dim c As Range

    With BoxHere
        .Font.Name = "Marlett"
        .HorizontalAlignment = xlRight
    End With
    
    'Preserve checks, but get rid of anything other than an "a"
    For Each c In BoxHere
        If c.Value <> "a" Then
            c.Value = ""
        End If
    Next c

End Sub

Function ArrayAppend(Array1 As Variant, Array2 As Variant) As Variant
'Returns an array with the length of both passed arrays
'If the passed arrays have different dimensions, the smaller value is used
'Returns nothing on error
    
    Dim i As Long
    Dim j As Long
    Dim L1 As Long
    Dim L2 As Long
    Dim L3 As Long
    Dim D1 As Long
    Dim D2 As Long
    Dim D3 As Long
    Dim TempString As String
    Dim ReturnArray As Variant
    
    'Check the dimensions of the arrays
    L1 = UBound(Array1, 2)
    L2 = UBound(Array2, 2)
    D1 = UBound(Array1, 1)
    D2 = UBound(Array2, 1)

    If L1 * L2 * D1 * D2 = 0 Then
        GoTo Footer
    End If
    
    'Make a new array
    L3 = L1 + L2
    
    If D1 < D2 Then
        D3 = D1
    Else
        D3 = D2
    End If
    
    ReDim ReturnArray(1 To D3, 1 To L3)
    
    For j = 1 To L1
        For i = 1 To D3
            ReturnArray(i, j) = Array1(i, j)
        Next i
    Next j
    
    For j = 1 To L2
        For i = 1 To D3
            ReturnArray(i, j + L1) = Array2(i, j)
        Next i
    Next j

    'Return
    ArrayAppend = ReturnArray

Footer:

End Function

Function ArrayMatch(Array1 As Variant, Array2 As Variant, WhichDimension As Long) As Variant
'Searches the 1st dimension of both arrays and attempts to match them
'Values from the passed dimension of Array2 are added to Array1 on matches
'Does not make the array longer than Array1
'Returns nothing if no matches
'Will overwrite if there are duplicate values in the 1st dimension, except for blanks

    Dim i As Long
    Dim j As Long
    Dim ReturnDim As Long
    Dim NewDim As Long
    Dim Key1 As String
    Dim Key2 As String
    Dim Value2 As String
    Dim ReturnArray As Variant

    'Check that the passed dimension exists
    If WhichDimension > 0 Then
        ReturnDim = WhichDimension
    Else
        ReturnDim = 1
    End If
    
    i = UBound(Array2, 1)
        If i < ReturnDim Then
            GoTo Footer
        End If
    
    'Make a new array to populate
    NewDim = UBound(Array1, 1)
        If LBound(Array1, 1) = 0 Then
            NewDim = NewDim + 1
        End If
        
    j = UBound(Array1, 2)
        If LBound(Array1, 2) = 0 Then
            j = j + 1
        End If
    

    ReDim ReturnArray(1 To NewDim, 1 To j)
    
    'Recreate the original array
    For i = 1 To UBound(Array1, 1)
        For j = 1 To UBound(Array1, 2)
            ReturnArray(i, j) = Array1(i, j)
        Next j
    Next i
    
    'Loop through and look for matches
    i = 1 'Not necessary, but makes me feel better
    j = 1
    
    For i = 1 To UBound(Array1, 2)
        Key1 = Array1(1, i)
        
        For j = 1 To UBound(Array2, 2)
            Key2 = Array2(1, j)
            
            If Key1 = Key2 Then
                Value2 = Array2(ReturnDim, j)
                
                If Len(Value2) > 0 Then
                    ReturnArray(NewDim, i) = Value2
                End If
            End If
        Next j
    Next i

    'Return
    ArrayMatch = ReturnArray

Footer:

End Function



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

Function NudgeToHeader(SourceSheet As Worksheet, SourceRange As Range, HeaderName As String) As Range
'Shifts a range to a particular column in a table
'Returns nudged range
'Returns nothing on error

    Dim TargetHeader As Range
    Dim TargetRange As Range
    Dim SourceTable As ListObject
    
    Set SourceTable = SourceSheet.ListObjects(1)

    Set TargetHeader = FindTableHeader(SourceSheet, HeaderName)
        If TargetHeader Is Nothing Then
            GoTo Footer
        ElseIf TargetHeader.Column = SourceRange.Column Then 'If it's already in the same column
            Set TargetRange = SourceRange
            
            GoTo ReturnRange
        End If
    
    Set TargetRange = Intersect(SourceRange.EntireRow, SourceTable.ListColumns(HeaderName).DataBodyRange) 'Not using offset to avoid swapping the sign of the # columns in .Offset()
        If TargetRange Is Nothing Then
            GoTo Footer
        End If
    
ReturnRange:
    Set NudgeToHeader = TargetRange
        
Footer:

End Function

Function PopulateListBox(UForm As Object, UListBox As Object, LabelArray As Variant) As Long
'Inserts three visible and one insivible columns for each entry in the passed ListBox
'LabelArray passes which activities to be added, since each form has different criteria for when they show up
'Returns 1 on success, 0 on failure

    Dim RecordsSheet As Worksheet
    Dim ActivitySheet As Worksheet
    Dim RecordsLabelRange As Range
    Dim ActivityLabelRange As Range
    Dim RecordsHeaderRange As Range
    Dim LabelCell As Range
    Dim c As Range
    Dim d As Range
    Dim i As Long
    Dim j As Long
    Dim LabelIndex As Long
    Dim LabelString As String
    Dim PracticeString As String
    Dim DateString As String
    Dim DescriptionString As String
    Dim NotEmpty As Boolean
    
    Dim ActivityArray As Variant
    Dim HeaderArray As Variant
    
    PopulateListBox = 0
    NotEmpty = False
    
    Set RecordsSheet = Worksheets("Records Page")
    
    'Grab the headers and values for all activities
    '(1, i) - Headers
    '(2, i) - Values for first activity
    '(3, i) - Values for second activity, etc.
    ActivityArray = GetRecordsActivityHeaders(RecordsSheet, , "All")
    
    'Loop through the ValueArray and compare against the labels in InfoArray. Erase any from ValueArray that aren't in InfoArray
    For i = 2 To UBound(ActivityArray, 1) 'The first element is "Label"
        LabelString = ActivityArray(i, 1)
        
        For j = 1 To UBound(LabelArray)
            If LabelArray(j) = LabelString Then
                GoTo ExtractValues
            End If
        Next j
    
        'If not found
        ActivityArray(i, 1) = ""
        GoTo NextLabel
        
ExtractValues:
        'Loop through headers and match values
        For j = 1 To UBound(ActivityArray, 2)
            Select Case ActivityArray(1, j)
            
                Case "Label"
                    LabelString = ActivityArray(i, j)
                
                Case "Practice"
                     PracticeString = ActivityArray(i, j)
                
                Case "Date"
                    DateString = ActivityArray(i, j)
                
                Case "Description"
                    DescriptionString = ActivityArray(i, j)
            
            End Select
        Next j

        'Insert into the ListBox
        With UListBox
            j = UListBox.ListCount 'Number of inserted items
        
            .AddItem LabelString
            .List(j, 1) = PracticeString
            .List(j, 2) = CDate(DateString)
            .List(j, 3) = DescriptionString
        End With

NextLabel:
    Next i

    PopulateListBox = 1

Footer:

End Function

Function PromptExport(DelRange As Range) As Long
'Gives a prompt to export students that are being deleted from the RecordsSheet
'Returns 1 if yes, 0 if no

    Dim i As Long
    Dim ConfirmMessage As String
    
    If DelRange Is Nothing Then
        GoTo Footer
    End If
    
    PromptExport = 0
    i = DelRange.Cells.Count
    
    'Different message based on which sheet the passed range is on
    If DelRange.Worksheet.Name = "Roster Page" Then
        ConfirmMessage = "Would you like to export these students' records before removing them?"
    
    ElseIf DelRange.Worksheet.Name = "Records Page" Then
        ConfirmMessage = "There are " & i & " students no longer on your roster. Would you like to export these students' records before removing them?"
    
    End If
    
    'Show pompt
    i = MsgBox(ConfirmMessage, vbQuestion + vbYesNo + vbDefaultButton2)
        If i <> vbYes Then
            GoTo Footer
        End If

    'Return
    PromptExport = 1

Footer:

End Function

Function PromptRemoveRecords(DelRange As Range) As Long
'Gives a prompt to confirm deletion of students from the RecordsSheet that are not on the RosterSheet
'Returns 1 if yes, 0 if no

    Dim i As Long
    Dim ConfirmMessage As String
    
    If DelRange Is Nothing Then
        GoTo Footer
    End If
    
    PromptRemoveRecords = 0
    
    i = DelRange.Cells.Count
    ConfirmMessage = i & " students have attendance recorded but are no longer on your roster." & vbCr & "Do you wish export their attendance before removing them?"
    
    'Show pompt
    i = MsgBox(ConfirmMessage, vbQuestion + vbYesNo + vbDefaultButton2)
        If i <> vbYes Then
            GoTo Footer
        End If

    'Return
    PromptRemoveRecords = 1

Footer:

End Function

Function PromptRemoveRoster(DelRange As Range) As Long
'Gives a prompt to confirm deletion of students from the RosterSheet
'Returns 1 if yes, 0 if no

    Dim i As Long
    Dim ConfirmMessage As String
    
    If DelRange Is Nothing Then
        GoTo Footer
    End If
    
    PromptRemoveRoster = 0
    
    i = DelRange.Cells.Count
    ConfirmMessage = "This will remove " & i & " students from your roster and cannot be undone." & vbCr & "Do you wish to continue?"
    
    'Show pompt
    i = MsgBox(ConfirmMessage, vbQuestion + vbYesNo + vbDefaultButton2)
        If i <> vbYes Then
            GoTo Footer
        End If

    'Return
    PromptRemoveRoster = 1

Footer:

End Function

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

