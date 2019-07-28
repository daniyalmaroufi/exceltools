Sub dothestuff()

flpath = Application.ActiveWorkbook.Path

canstart = MsgBox(flpath & "  Do you want to start?", vbYesNo)

If canstart = vbYes Then

n = Sheets("Sheet2").Range("A" & Rows.Count).End(xlUp).Row
For i = 230 To n

    Sheets("Sheet2").Range("B" & i).Copy Sheets("Sheet1").Range("C2")
    Sheets("Sheet1").Range("C2") = Sheets("Sheet1").Range("C2") / 1000000
    Sheets("Sheet2").Range("C" & i).Copy Sheets("Sheet1").Range("B2")
    Sheets("Sheet1").Range("B2") = Sheets("Sheet1").Range("B2") / 1000000
    Sheets("Sheet2").Range("D" & i).Copy Sheets("Sheet1").Range("C3")
    Sheets("Sheet2").Range("E" & i).Copy Sheets("Sheet1").Range("B3")
    Sheets("Sheet2").Range("F" & i).Copy Sheets("Sheet1").Range("C4")
    Sheets("Sheet2").Range("G" & i).Copy Sheets("Sheet1").Range("B4")
    Sheets("Sheet2").Range("H" & i).Copy Sheets("Sheet1").Range("C5")
    Sheets("Sheet2").Range("I" & i).Copy Sheets("Sheet1").Range("B5")
    flname = Sheets("Sheet2").Range("A" & i).Value
    Sheets("Sheet1").ChartObjects("mychart").Chart.Export flpath & "/charts/" & flname & ".jpg"
    
    cancontinue = MsgBox(i & "  Do you want to continue?", vbYesNo)
    If cancontinue = vbNo Then
        Exit For
    End If
    

Next i

Else

MsgBox "Finish!"

End If

End Sub
