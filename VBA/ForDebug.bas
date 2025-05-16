Attribute VB_Name = "ForDebug"
Option Explicit

Sub ScreenUpdating()

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True

End Sub

Sub BreakExternalLinks()
'PURPOSE: Breaks all external links that would show up in Excel's "Edit Links" Dialog Box
'SOURCE: www.TheSpreadsheetGuru.com/The-Code-Vault

Dim ExternalLinksArray As Variant
Dim wb As Workbook
Dim x As Long

Set wb = ActiveWorkbook

'Create an Array of all External Links stored in Workbook
  ExternalLinksArray = wb.LinkSources(Type:=xlLinkTypeExcelLinks)

'if the array is not empty the loop Through each External Link in ActiveWorkbook and Break it
 If IsEmpty(ExternalLinksArray) = False Then
     For x = 1 To UBound(ExternalLinksArray)
        wb.BreakLink Name:=ExternalLinksArray(x), Type:=xlLinkTypeExcelLinks
      Next x
End If

End Sub
