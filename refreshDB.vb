' Author: Parikshita Tripathi
' Date : 07/29/2016

Sub RefreshData()
   
    ' define variables
    Dim src As Excel.Workbook, dest As Excel.Workbook
    Dim wsSrc As Worksheet, wsDest As Worksheet
    Dim srcRange As Range, destRange As Range
    
    ' define destination sheet
    Set dest = ThisWorkbook
    Application.ScreenUpdating = 0

    
    ' set source information
    Set src = Workbooks.Open(“file.xlsx")
    Set wsSrc = src.Sheets("Data")
    Set srcRange = wsSrc.UsedRange
    
    ' set destination information
    Set wsDest = ActiveSheet
    Set wsDest = dest.Sheets("Data")
    wsDest.UsedRange.Clear
    'wsDest.UsedRange.Clear
    
    Set destRange = wsDest.Cells(wsDest.Rows.Count, 1).End(xlUp).Offset(0, 0)
    
    ' copy sorce sheet
    srcRange.Copy
    
    ' paste the copied data in destiantion sheet
    destRange.PasteSpecial Paste:=xlPasteValuesAndNumberFormats
    
    ' close source file
    Application.CutCopyMode = False
    src.Close True
    
End Sub