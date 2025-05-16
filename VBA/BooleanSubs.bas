Attribute VB_Name = "BooleanSubs"
Option Explicit

Function IsChecked(TargetRange As Range, Optional SearchType As String) As Boolean
'Checks a range for any students marked present
'If "Absent" is passed, it looks for absent students
'If "All" is passed, looks for absent and present students

    IsChecked = False
    
    If SearchType = "Absent" Then
        GoTo AbsentSearch
    ElseIf SearchType = "All" Then
        GoTo AllSearch
    End If
    
PresentSearch:
    'Check for "a" and "1" for present
    If Not TargetRange.Find("1", , xlValues, xlWhole) Is Nothing Then
        IsChecked = True
    ElseIf Not TargetRange.Find("a", , xlValues, xlWhole) Is Nothing Then
        IsChecked = True
    End If
    
    GoTo Footer
    
AbsentSearch:
    'Check for "0" for absent
    If Not TargetRange.Find("0", , xlValues, xlWhole) Is Nothing Then
        IsChecked = True
    End If
    
    GoTo Footer
    
AllSearch:
    'Check for anything
    If Not TargetRange.Find("*", , xlValues, xlWhole) Is Nothing Then
        IsChecked = True
    End If
    
    GoTo Footer
    
Footer:

End Function

Function IsCollege() As Boolean
'Simple function to determine if the workbooks is being used for College Prep or not

    Dim CoverSheet As Worksheet
    
    Set CoverSheet = Worksheets("Cover Page")
    
    If InStr(1, CoverSheet.Range("A1").Value, "College") > 0 Then
        IsCollege = True
    Else
        IsCollege = False
    End If

End Function
