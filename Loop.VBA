# Auto open multiple file in loop and copy paste in one sheet

Dim wrkMyWorkBook As Workbook

Dim i As Integer

Dim sh As Worksheet
Set sh = ThisWorkbook.Sheets("MAIN")


Dim last_row As Integer
last_row = Application.CountA(sh.Range("S:S"))

Sheets("MAIN").Select



With ActiveSheet

        

For i = 1 To last_row
        



Set wrkMyWorkBook = Workbooks.Open(Filename:=Sheets("MAIN").Range("I11").Value & Sheets("MAIN").Range("P" & i).Value)
'Set wrkMyWorkBook = Workbooks.Open(sh.Range("R" & i).Value)
Range("A4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Windows("Employee_rating_Send_Mail.xlsm").Activate
    ActiveWindow.ScrollWorkbookTabs Sheets:=11
    Sheets("help").Select
    Range("A4").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Range("K7").Select
    Windows("Employee_rating_Send_Mail.xlsm").Activate
    Sheets("MAIN").Select
    Workbooks(ActiveSheet.Range("P" & i).Value).Activate
    ActiveWorkbook.Close
    Windows("Employee_rating_Send_Mail.xlsm").Activate
    Sheets("MAIN").Select


Next

End With
