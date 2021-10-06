Attribute VB_Name = "Module1"
Sub copyExtraction1()

'Copy needed cells from Extraction 1
For i = 11 To Sheets("GL5_Master").Range("A" & Rows.Count).End(xlUp).Row
Sheets("parts_station").Cells(i - 9, 1) = Sheets("GL5_Master").Cells(i, 1)
Sheets("parts_station").Cells(i - 9, 2) = Sheets("GL5_Master").Cells(i, 2)
Sheets("parts_station").Cells(i - 9, 3) = Sheets("GL5_Master").Cells(i, 3)
Sheets("parts_station").Cells(i - 9, 4) = Sheets("GL5_Master").Cells(i, 5)
Sheets("parts_station").Cells(i - 9, 5) = Sheets("GL5_Master").Cells(i, 4)
Sheets("parts_station").Cells(i - 9, 6) = Sheets("GL5_Master").Cells(i, 6)
Sheets("parts_station").Cells(i - 9, 7) = Sheets("GL5_Master").Cells(i, 9)
Sheets("parts_station").Cells(i - 9, 10) = Sheets("GL5_Master").Cells(i, 7)
Sheets("parts_station").Cells(i - 9, 11) = Sheets("GL5_Master").Cells(i, 8)
Sheets("parts_station").Cells(i - 9, 13) = Sheets("GL5_Master").Cells(i, 10)
Sheets("parts_station").Cells(i - 9, 14) = Sheets("GL5_Master").Cells(i, 11)
Sheets("parts_station").Cells(i - 9, 15) = Sheets("GL5_Master").Cells(i, 12)
Next i

End Sub
