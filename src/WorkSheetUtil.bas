Attribute VB_Name = "WorkSheetUtil"
Option Explicit

' ワークシートをコピーします。
'https://docs.microsoft.com/ja-jp/office/vba/api/excel.worksheet.copy
Public Sub CopySheet(ByVal src As String, ByVal dist As String)

    Call Worksheets(src).Copy(After:=Worksheets(Worksheets.Count))
    ActiveSheet.Name = dist

End Sub
