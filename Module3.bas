Attribute VB_Name = "Module3"
Sub newFcTest()
'Test for connecting fcs
For i = 1 To 11
fc1 = Sheets("test").Cells(i, 1)

    If Len(fc1) > 8 Then
        nr1 = Mid(fc1, 1, 5)
        Sheets("test").Cells(i, 2) = nr1
        nr2 = Mid(fc1, 7, Len(fc1))
        Sheets("test").Cells(i, 3) = nr2
        nr3 = nr1 + nr2
        Sheets("test").Cells(i, 4) = nr3
    Else
        Sheets("test").Cells(i, 4) = fc1
    End If

Next i
End Sub
