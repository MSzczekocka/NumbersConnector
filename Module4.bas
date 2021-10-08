Attribute VB_Name = "Module4"
Sub addQuantity()

'add quantity of part - logic similar like in module2
For i = 6 To Sheets("parts_station").Range("A" & Rows.Count).End(xlUp).Row
    psPSTemp = Sheets("parts_station").Cells(i, 3)
    psPS = Mid(psPSTemp, 1, 9)
    nrPS = Sheets("parts_station").Cells(i, 6)
    fcPS = Sheets("parts_station").Cells(i, 7)
        For j = 3 To Sheets("04").Range("G" & Rows.Count).End(xlUp).Row
            ps04 = Sheets("04").Cells(j, 8)
            nr04 = Sheets("04").Cells(j, 21)
            fc04 = Sheets("04").Cells(j, 24)
                If fc04 = "" Then
                    fc04Temp = Sheets("04").Cells(j, 25)
                    nr1 = Mid(fc04Temp, 1, 5)
                    nr2 = Mid(fc04Temp, 7, Len(fc04Temp))
                    fc04 = nr1 + nr2
                        If fc04 = "" Then GoTo 50
                End If

                If psPS = ps04 Then
                    If InStr(nrPS, nr04) <> 0 And InStr(fcPS, fc04) <> 0 Then
                        Sheets("parts_station").Cells(i, 12) = Sheets("04").Cells(j, 26)
                        GoTo 100
                    End If
                End If

50      Next j
100 Next i
End Sub
