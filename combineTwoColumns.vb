Sub dothestuff()


canstart = MsgBox("  Do you want to start?", vbYesNo)

If canstart = vbYes Then

For i = 1 To 50

    name = Range("G" & i).Value
    lname = Range("H" & i).Value
    Range("B" & i).Value = name & " " & lname
    
    cancontinue = MsgBox(i & "  Do you want to continue?", vbYesNo)
    If cancontinue = vbNo Then
        Exit For
    End If
    

Next i

Else

MsgBox "Finish!"

End If

End Sub
