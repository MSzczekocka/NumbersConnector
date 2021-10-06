Attribute VB_Name = "Module2"
Sub numbersConnector()

'Get number 1 (extraction 1)
For i = 2 To Sheets("parts_station").Range("A" & Rows.Count).End(xlUp).Row
    nrPsaPSTemp = Sheets("parts_station").Cells(i, 6)
    'Get 8 digits from number (end of number can be differ depending on extraction)
    nrPsaPS = Mid(nrPsaPSTemp, 1, 8)
    fcPS = Sheets("parts_station").Cells(i, 7)
        
        ' Get number 2 (extraction 2)
        For j = 3 To Sheets("100").Range("A" & Rows.Count).End(xlUp).Row
            nrPsa100Temp = Sheets("100").Cells(j, 40)
            'Get 8 digits from number (end of number can be differ depending on extraction)
            nrPsa100 = Mid(nrPsa100Temp, 1, 8)
                'Skip empty cells
                If nrPsa100 = "" Then GoTo 50
                
            fc100Temp = Sheets("100").Cells(j, 3)
                'if connecting fc from differ extractions
                If Len(fc100Temp) > 8 Then
                    nr1 = Mid(fc100Temp, 1, 5)
                    nr2 = Mid(fc100Temp, 7, Len(fc100Temp))
                    fc100 = nr1 + nr2
                Else
                    fc100 = fc100Temp
                End If
                ' if connecting numbers and their fc
                If nrPsaPS = nrPsa100 And InStr(fcPS, fc100) <> 0 Then
                    Sheets("parts_station").Cells(i, 8) = Sheets("100").Cells(j, 41)
                    Sheets("parts_station").Cells(i, 9) = Sheets("100").Cells(j, 50)
                    GoTo 100
                End If

50      Next j

100 Next i

End Sub
