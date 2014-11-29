```vb
Option Explicit
Private insertCell, insertRow, lastRow, pasteRow, copyRow As Range
Private firstCol, checkCol, codeCol, i, j As Integer

Sub shiftRowsDown()
Application.ScreenUpdating = False

'firstCol is the first column in each row to be selected for copying
firstCol = 9
'codeCol is the last column in each row to be selected for copying and _
    also the one with the code that determines the insertion point, rows that can move, and rows that can't move
codeCol = 14
'checkCol is the column to be checked for blanks
checkCol = 10

Application.Run "initialRangeDown"
Application.Run "findPasteRowDown"
Application.Run "findCopyRowDown"

If copyRow.Row = insertRow.Row Then
    Application.Run "copyPasteDown"
    Else
        Do While copyRow.Row <> insertRow.Row
            Application.Run "findCopyRowDown"
            Application.Run "copyPasteDown"
            Set pasteRow = copyRow
            If copyRow.Row = insertRow.Row Then Exit Do
            Loop
End If

Application.Run "removeCF"
Application.Run "resetCF"

Application.ScreenUpdating = True
End Sub

Sub initialRangeDown()
Application.ScreenUpdating = False

    Set insertCell = Columns(codeCol).Find("I", lookat:=xlWhole)
    Set insertRow = Range(Cells(insertCell.Row, firstCol), Cells(insertCell.Row, codeCol))
    
End Sub

Sub findPasteRowDown()
'PasteRow could also be called BlankRow
Application.ScreenUpdating = False

lastRow = Cells(Rows.Count, 1).End(xlUp).Row
For i = insertCell.Row To lastRow
    If IsEmpty(Cells(i, checkCol)) Then
        Set pasteRow = Range(Cells(i, firstCol), Cells(i, codeCol))
            Exit For
    End If
Next i

End Sub

Sub findCopyRowDown()
Application.ScreenUpdating = False

i = 0

Do
    i = i + 1
    Set copyRow = Range(Cells(pasteRow.Row - i, firstCol), Cells(pasteRow.Row - i, codeCol))
Loop Until Cells(pasteRow.Row, codeCol).Offset(-i, 0) <> "f"

End Sub

Sub copyPasteDown()
Application.ScreenUpdating = False

        copyRow.Copy
        pasteRow.PasteSpecial xlPasteAllExceptBorders
        copyRow.ClearContents
End Sub

Sub removeCF()
    With Worksheets("14-15").Range("$G$3:$K$159")
        .FormatConditions.Delete
    End With
End Sub

Sub resetCF()
    Dim Rng1, Rng2 As Range
    Set Rng1 = Range("$G$3:$K$159")
    Set Rng2 = Range("$H$3:$K$159")

    Rng1.FormatConditions.Add Type:=xlExpression, Formula1:="=$D3:$D159=""OFF"""
    Rng1.FormatConditions(Rng1.FormatConditions.Count).SetFirstPriority
    Rng1.FormatConditions(1).Interior.Color = RGB(217, 217, 217)
    Rng1.FormatConditions(1).StopIfTrue = False
    
    Rng2.FormatConditions.Add Type:=xlExpression, Formula1:="=$A3:$A159=""l"""
    Rng2.FormatConditions(Rng2.FormatConditions.Count).SetFirstPriority
    Rng2.FormatConditions(1).Interior.Color = RGB(198, 223, 251)
    Rng2.FormatConditions(1).StopIfTrue = False
    
    Rng2.FormatConditions.Add Type:=xlExpression, Formula1:="=$A3:$A159=""t"""
    Rng2.FormatConditions(Rng2.FormatConditions.Count).SetFirstPriority
    Rng2.FormatConditions(1).Interior.Color = RGB(228, 255, 202)
    Rng2.FormatConditions(1).StopIfTrue = False

End Sub
```
