Attribute VB_Name = "TabulateSubs"
Option Explicit
Option Compare Text

Sub TabulateActivities(ActivitySheet As Worksheet)
'Finds all students who are checked and returns the range we care about
'F-name, L-name, ethnicity, gender, grade

    Dim ReportSheet As Worksheet
    Dim BlankCheck As Range
    Dim TableStart As Range
    Dim PracticeName As String
    Dim PracticeNotes As String
    Dim DateString As String
    Dim StaffName As String
    Dim CenterName As String
    Dim LRow As Long
    Dim ActLRow As Long
    
    'Find the activity, center, name, date
    Set ReportSheet = Worksheets("Report Page")
    Set BlankCheck = ReportSheet.Range("B7")
    
    If ReportSheet.ProtectContents = True Then
        ReportSheet.Unprotect
    End If
    
    'Make sure the totals are in there. This protects the sheet again, so we need to unprotect it
    If IsEmpty(BlankCheck) Then
        Call PullReportTotals
        ReportSheet.Unprotect
    End If
    
    With ActivitySheet
        PracticeName = .Range("F1").Value
        PracticeNotes = .Range("F3").Value
        StaffName = .Range("B1").Value
        CenterName = .Range("B2").Value
        DateString = .Range("B3").Value
    End With
    
    LRow = ReportSheet.Cells(Rows.Count, 2).End(xlUp).Row
    
    With ReportSheet
        .Cells(LRow + 1, 2).Value = CenterName
        .Cells(LRow + 1, 3).Value = StaffName
        .Cells(LRow + 1, 3).NumberFormat = "yyyy-mm-dd"
        .Cells(LRow + 1, 4).Value = DateString
        .Cells(LRow + 1, 5).Value = PracticeName
        .Cells(LRow + 1, 6).Value = PracticeNotes
    End With
    
    'Find which students were present. Create range and count
    Dim IsChecked As Range
    Dim NumChecked As Long
    Dim c As Range

    Set TableStart = ActivitySheet.Range("A6")
    NumChecked = 0
    ActLRow = ActivitySheet.Cells(Rows.Count, 2).End(xlUp).Row
    
    For Each c In ActivitySheet.Range(Cells(TableStart.Row + 1, 1).Address, Cells(ActLRow, 1).Address)
        If c.Value <> "" Then
            If Not IsChecked Is Nothing Then
                Set IsChecked = Union(IsChecked, c)
            Else
                Set IsChecked = c
            End If
            NumChecked = NumChecked + 1
        End If
    Next c
    
    'Tabulate for each category
    Dim TempArray As Variant
    Dim RaceRange As Range
    Dim GenderRange As Range
    Dim GradeRange As Range
    Dim TotalRange As Range
    
    Set RaceRange = ReportSheet.Range("H1:O1").Offset(LRow, 0)
    Set GenderRange = ReportSheet.Range("P1:R1").Offset(LRow, 0)
    Set GradeRange = ReportSheet.Range("S1:V1").Offset(LRow, 0)
    
    TempArray = DemoTabulate(ActivitySheet, NumChecked, IsChecked, "Race")
    RaceRange = TempArray
    Erase TempArray
    
    TempArray = DemoTabulate(ActivitySheet, NumChecked, IsChecked, "Gender")
    GenderRange = TempArray
    Erase TempArray
    
    TempArray = DemoTabulate(ActivitySheet, NumChecked, IsChecked, "Grade")
    GradeRange = TempArray
    Erase TempArray
    
    'Total
    Set TotalRange = ReportSheet.Range(Cells(LRow + 1, 7).Address)
    TotalRange.Value = NumChecked
    
    'Conditional Formatting
    Dim AllRange As Range
    Dim TotalStart As Range
    Dim FormulaString As String
    
    Set AllRange = Union(RaceRange, GenderRange, GradeRange, TotalRange)
    Set TotalStart = ReportSheet.Range("G7")
    For Each c In AllRange
        FormulaString = "=" + ReportSheet.Cells(TotalStart.Row, c.Column).Address + "<" + c.Address
        c.FormatConditions.Add Type:=xlExpression, Formula1:=FormulaString
        With c.FormatConditions(1)
            .StopIfTrue = False
            .Interior.Color = vbRed
        End With
    Next c
    
    'Put a checkbox in the first column
    Dim BoxRange As Range
    
    Set BoxRange = ReportSheet.Cells(LRow + 1, 1)
    Call AddMarlettBox(BoxRange, ReportSheet)
    
End Sub

Function StudentsSelected(ActivitySheet As Worksheet) As Boolean
'Checks to make sure that students are selected on each sheet
    
    Dim LRow As Long
    Dim SearchRange As Range
    Dim TableStart As Range
    
    Set TableStart = ActivitySheet.Range("A6")
    LRow = ActivitySheet.Cells(Rows.Count, 2).End(xlUp).Row
    
    'This shouldn't happen, but included in case something unexpected occurs
    If LRow = TableStart.Row Then
        MsgBox ("You don't have any students on " & ActivitySheet.Name & "." & _
            "Please add at least one students to that sheet.")
        StudentsSelected = False
        Exit Function
    End If
    
    Set SearchRange = ActivitySheet.Range(Cells(TableStart.Row, 1).Address, Cells(LRow, 1).Address).Find("a", LookIn:=xlValues)
    
    If SearchRange Is Nothing Then
        MsgBox ("You don't have any students selected on " & ActivitySheet.Name & ". " & vbCr & _
            "Please select at least one student on that sheet.")
        StudentsSelected = False
        Exit Function
    End If
        
    StudentsSelected = True

End Function

Function NameDatePractice(ActivitySheet As Worksheet) As Boolean
'Make sure activity info is filled out

    Dim NameString As String
    Dim DateString As String
    Dim PracticeString As String
    
    NameString = ActivitySheet.Range("B1").Value
    DateString = ActivitySheet.Range("B3").Value
    PracticeString = ActivitySheet.Range("F1").Value
    
    If Len(NameString) < 1 Or Len(DateString) < 1 Or Len(PracticeString) < 1 Then
        MsgBox ("Please fill out your name, date, and practice on page " & ActivitySheet.Name & ".")
        NameDatePractice = False
    Else
        NameDatePractice = True
    End If

End Function

Function DemoTabulate(ActivitySheet As Worksheet, NumChecked As Long, SearchRange As Range, SearchType As String) As Variant
'Count how many fall into each demographic category and return an array
'Can specify if we're looking for race, gender, or grade

    Dim RaceArray As Variant
    Dim GenderArray As Variant
    Dim GradeArray As Variant
    Dim SearchArray As Variant
    Dim CountArray As Variant
    Dim SearchTerm As String
    Dim SearchHere As Range
    Dim i As Long
    Dim c As Range
    
    RaceArray = Array("White", "Asian", "Black", "Latino", "AIAN", "NHPI", "Mixed", "Other")
    GenderArray = Array("Female", "Male", "Other")
    GradeArray = Array("<45", "45-90", ">90", "Other")
    
    'What we are searching
    If SearchType = "Race" Then
        ReDim SearchArray(0 To UBound(RaceArray))
        ReDim CountArray(0 To UBound(RaceArray))
        SearchArray = RaceArray
        Set SearchHere = SearchRange.Offset(0, 3)
        
    ElseIf SearchType = "Gender" Then
        ReDim SearchArray(0 To UBound(GenderArray))
        ReDim CountArray(0 To UBound(GenderArray))
        SearchArray = GenderArray
        Set SearchHere = SearchRange.Offset(0, 4)
        
    ElseIf SearchType = "Grade" Then
        ReDim CountArray(0 To UBound(GradeArray))
        Set SearchHere = SearchRange.Offset(0, 5)
        GoTo GradeCount
    End If
    
    ReDim CountArray(0 To UBound(SearchArray))
   
    'CountIf doesn't work with discontiguous ranges
    For i = 0 To UBound(SearchArray)
        SearchTerm = SearchArray(i)
        For Each c In SearchHere
            If Trim(c.Value) = SearchTerm Then
                CountArray(i) = CountArray(i) + 1
            End If
        Next c
    Next i
    GoTo MissingValues

GradeCount:
    For Each c In SearchHere
        If IsEmpty(c.Value) Then 'VBA will return true for <45 on empty cells
            CountArray(3) = CountArray(3) + 1
        ElseIf Not (IsNumeric(c.Value)) Then
            CountArray(3) = CountArray(3) + 1
        ElseIf c.Value < 45 Then
            CountArray(0) = CountArray(0) + 1
        ElseIf c.Value <= 90 Then
            CountArray(1) = CountArray(1) + 1
        ElseIf c.Value > 90 Then
            CountArray(2) = CountArray(2) + 1
        Else
            CountArray(3) = CountArray(3) + 1 'To catch anything else
        End If
    Next c
    
   GoTo Footer
    
MissingValues:
    'Blank and invalid entries aren't counted above
    Dim Missing As Long
    Dim ArrayIndex As Long
    
    Missing = NumChecked - WorksheetFunction.Sum(CountArray)
    ArrayIndex = UBound(CountArray)
    
    If Missing > 0 Then
        CountArray(ArrayIndex) = CountArray(ArrayIndex) + Missing
    End If
    
    Erase SearchArray

Footer:
    DemoTabulate = CountArray
    Erase CountArray
    
End Function
