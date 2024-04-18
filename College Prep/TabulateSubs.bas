Attribute VB_Name = "TabulateSubs"
Option Explicit

Function DemoTabulate(NumChecked As Long, SearchRange As Range, SearchType As String) As Variant
'Count how many fall into each demographic category and return an array
'Can specify if we're looking for race, gender, or grade

    Dim RaceArray() As Variant
    Dim GenderArray() As Variant
    Dim GradeArray() As Variant
                             
    Dim SearchArray() As Variant
    Dim CountArray() As Variant
    Dim SearchTerm As String
    Dim SearchHere As Range
    Dim i As Long
    Dim c As Range
                                                                                       
     RaceArray = Application.Transpose(Range("EthnicityList"))
     GenderArray = Application.Transpose(Range("GenderList"))
     GradeArray = Application.Transpose(Range("GradeList"))
                                                                                           
    'What we are searching
    If SearchType = "Race" Then
        ReDim SearchArray(1 To UBound(RaceArray))
        ReDim CountArray(1 To UBound(RaceArray))
        SearchArray = RaceArray
        Set SearchHere = SearchRange.Offset(0, 2)
        
    ElseIf SearchType = "Gender" Then
        ReDim SearchArray(1 To UBound(GenderArray))
        ReDim CountArray(1 To UBound(GenderArray))
        SearchArray = GenderArray
        Set SearchHere = SearchRange.Offset(0, 3)
        
    ElseIf SearchType = "Grade" Then
        ReDim SearchArray(1 To UBound(GradeArray))
        ReDim CountArray(1 To UBound(GradeArray))
        SearchArray = GradeArray
        Set SearchHere = SearchRange.Offset(0, 4)
        GoTo GradeCount
    End If
    
    ReDim CountArray(1 To UBound(SearchArray))
   
    'CountIf doesn't work with discontiguous ranges
    For i = 1 To UBound(SearchArray)
        SearchTerm = SearchArray(i)
        For Each c In SearchHere
            If Trim(c.Value) = SearchTerm Then
                CountArray(i) = CountArray(i) + 1
            End If
        Next c
    Next i
    GoTo MissingValues

GradeCount:
    'We can't use Trim() for this
    For i = 1 To UBound(SearchArray)
        SearchTerm = SearchArray(i)
        For Each c In SearchHere
            If c.Value = SearchTerm Then
                CountArray(i) = CountArray(i) + 1
            End If
        Next c
              
    Next i
              
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
    DemoTabulate = CountArray
    Erase CountArray
    
End Function

Sub TabulateActivity(LabelString As String)
'Pushes tabulation to the report page for a single activity
'Called automatically when an activity is saved
'Can be called from anywhere

    Dim ReportSheet As Worksheet
    Dim RosterSheet As Worksheet
    Dim RecordsSheet As Worksheet
    Dim CoverSheet As Worksheet
    Dim MatchCell As Range
    
    Set ReportSheet = Worksheets("Report Page")
    Set RosterSheet = Worksheets("Roster Page")
    Set RecordsSheet = Worksheets("Records Page")
    Set CoverSheet = Worksheets("Cover Page")
    
    'First verify that the activity has been saved
    Set MatchCell = RecordsSheet.Range("1:1").Find(LabelString, , xlValues, xlWhole)
    
    If MatchCell Is Nothing Then
        MsgBox ("There was a problem tabulating this activity." & vbCr & "Please save the activity and try again.")
        GoTo Footer
    End If

    'Grab the activity information
    Dim InfoArray() As Variant
    
    InfoArray = Application.Transpose(Range(MatchCell, MatchCell.Offset(3, 0)).Value)

    'Identify all students who were present
    Dim SearchRange As Range
    Dim PresentRange As Range
    Dim c As Range
    Dim FRow As Long
    Dim LRow As Long
    
    FRow = RecordsSheet.Range("A:A").Find("H BREAK", , xlValues, xlWhole).Row
    LRow = RecordsSheet.Range("A:A").Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    Set SearchRange = RecordsSheet.Range(Cells(FRow + 1, MatchCell.Column).Address, Cells(LRow, MatchCell.Column).Address)
    
    For Each c In SearchRange
        If c.Value = "a" Then
            If Not PresentRange Is Nothing Then
                Set PresentRange = Union(PresentRange, RecordsSheet.Cells(c.Row, 1))
            Else
                Set PresentRange = RecordsSheet.Cells(c.Row, 1)
            End If
        End If
    Next c
    
    'If nothing was found on the Records sheet
    If PresentRange Is Nothing Then
        'MsgBox ("No attendance is recorded for the activity labeled " & LabelString & "." & vbCr & _
        '"Please save the activity and try again.")
        GoTo Footer
    End If
    
    'Use this to find the same students on the Roster page
    Dim TabulateRange As Range
    Dim NumChecked As Long
    
    Set SearchRange = RosterSheet.ListObjects("RosterTable").ListColumns("First").DataBodyRange
    
    For Each c In PresentRange
        Set MatchCell = NameMatch(c, SearchRange)
        If Not MatchCell Is Nothing Then
            If Not TabulateRange Is Nothing Then
                Set TabulateRange = Union(TabulateRange, MatchCell)
            Else
                Set TabulateRange = MatchCell.Offset
            End If
            NumChecked = NumChecked + 1 'For the total later
        End If
    Next c
    
    'Define the ranges we need on the Report page
    Dim CenterRange As Range
    Dim NameRange As Range
    Dim InfoRange As Range
    Dim TotalRange As Range
    Dim RaceRange As Range
    Dim GenderRange As Range
    Dim GradeRange As Range
                           
    Dim PasteRow As Long
    Dim TempArray() As Variant
    
    Set CenterRange = FindReportRange("Center")
    Set NameRange = FindReportRange("Name")
    Set InfoRange = FindReportRange("Label", "Description")
    Set TotalRange = FindReportRange("Total")
    Set RaceRange = FindReportRange("White", "Other Race")
    Set GenderRange = FindReportRange("Female", "Other Gender")
    Set GradeRange = FindReportRange("6", "Other Grade")
                                                              
    
    'See if the activity already exists on the Report page
    Set MatchCell = ReportSheet.Range("D:D").Find(LabelString, , xlValues, xlWhole)
    
    'If there's no match, put it at the bottom. Else, overwrite it
    If MatchCell Is Nothing Then
        PasteRow = ReportSheet.Range("D:D").Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row + 1
    Else
        PasteRow = MatchCell.Row
    End If
    
    'Insert center and submitter
    CenterRange.Offset(PasteRow - 6, 0) = CoverSheet.Range("B5")
    NameRange.Offset(PasteRow - 6, 0) = CoverSheet.Range("B3")
    
    'Insert the activity information and total. The data range of the table starts in row 7
    InfoRange.Offset(PasteRow - 6, 0) = InfoArray
    TotalRange.Offset(PasteRow - 6, 0) = NumChecked
    
    'Pass the demographics for tabulation
    TempArray = DemoTabulate(NumChecked, TabulateRange, "Race")
    RaceRange.Offset(PasteRow - 6, 0) = TempArray
    Erase TempArray
    
    TempArray = DemoTabulate(NumChecked, TabulateRange, "Gender")
    GenderRange.Offset(PasteRow - 6, 0) = TempArray
    Erase TempArray
    
    TempArray = DemoTabulate(NumChecked, TabulateRange, "Grade")
    GradeRange.Offset(PasteRow - 6, 0) = TempArray
    Erase TempArray

    'Insert a Marlett box
    Call AddMarlettBox(ReportSheet.Cells(PasteRow, 1), ReportSheet)
    
Footer:

End Sub

Sub TabulateAll()
'Tabulates all saved activities

    Dim ReportSheet As Worksheet
    Dim RecordsSheet As Worksheet
    
    Set RecordsSheet = Worksheets("Records Page")
    Set ReportSheet = Worksheets("Report Page")
    
    'Make sure we have at least one saved activity
    Dim FCell As Range
    Dim LCell As Range
    
    Set FCell = RecordsSheet.Range("1:1").Find("V BREAK", , xlValues, xlWhole) 'Designates the cell before labels start
    Set LCell = RecordsSheet.Range("1:1").Find("*", SearchOrder:=xlByColumns, SearchDirection:=xlPrevious)
    
    If FCell.Column = LCell.Column Then
        MsgBox ("You have no saved activities.")
        GoTo Footer
    End If
    
    'Unprotect the Report page
    Call UnprotectCheck(ReportSheet)
    
    'Make sure totals are present
    Call PullReportTotals
    
    'Loop through all labels on the Report page and pass for tabulation
    Dim LabelRange As Range
    Dim c As Range
    
    Set LabelRange = RecordsSheet.Range(FCell.Offset(0, 1), LCell)
    
    For Each c In LabelRange
        Call TabulateActivity(c.Value)
    Next c
    
Footer:
    Call ResetProtection
    
End Sub

Sub RetabulateActivities()
'If a student is removed from the roster, retabulate the all activities on the Report page

    Dim ReportSheet As Worksheet
    Dim RecordsSheet As Worksheet

    Set RecordsSheet = Worksheets("Records Page")
    Set ReportSheet = Worksheets("Report Page")
    
    'Check to see if there are any activities saved on the Records sheet
    Dim FCol As Long
    Dim LCol As Long
    
    FCol = RecordsSheet.Range("1:1").Find("V BREAK", , xlValues, xlWhole).Column
    LCol = RecordsSheet.Range("1:1").Find("*", SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
    
    If Not LCol > FCol Then
        GoTo Footer 'We don't need to do anything
    End If
    
    'Check if any activities are on the Report sheet
    Dim FRow As Long
    Dim LRow As Long
    Dim LabelCol As Long
    
    FRow = ReportSheet.Range("A:A").Find("Select", , xlValues, xlWhole).Row
    LabelCol = ReportSheet.Cells(FRow, 1).EntireRow.Find("Label", , xlValues, xlWhole).Column
    LRow = ReportSheet.Cells(FRow, LabelCol).EntireColumn.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    
    If Not LRow > FRow + 1 Then 'The first row is the header, the second is the totals row
        GoTo Footer
    End If
    
    'Create a list of all labels on the Report sheet
    Dim LabelArray() As String
    Dim i As Long
    Dim j As Long
    
    j = 0
    For i = FRow + 2 To LRow
        j = j + 1
        ReDim Preserve LabelArray(1 To j)
        LabelArray(j) = ReportSheet.Cells(i, LabelCol).Value
    Next i
    
    'Tabulate for each element of the array
    For i = 1 To j
        Call TabulateActivity(LabelArray(i))
    Next i
    
Footer:

End Sub


