Attribute VB_Name = "Module1"
Sub copyExtraction1()

'Copy needed cells from Extraction 1
For i = 11 To Sheets("GL5_Master").Range("A" & Rows.Count).End(xlUp).Row
Sheets("parts_station").Cells(i - 5, 1) = Sheets("GL5_Master").Cells(i, 1)
Sheets("parts_station").Cells(i - 5, 2) = Sheets("GL5_Master").Cells(i, 2)
Sheets("parts_station").Cells(i - 5, 3) = Sheets("GL5_Master").Cells(i, 3)
Sheets("parts_station").Cells(i - 5, 4) = Sheets("GL5_Master").Cells(i, 5)
Sheets("parts_station").Cells(i - 5, 5) = Sheets("GL5_Master").Cells(i, 4)
Sheets("parts_station").Cells(i - 5, 6) = Sheets("GL5_Master").Cells(i, 6)
Sheets("parts_station").Cells(i - 5, 7) = Sheets("GL5_Master").Cells(i, 9)
Sheets("parts_station").Cells(i - 5, 10) = Sheets("GL5_Master").Cells(i, 7)
Sheets("parts_station").Cells(i - 5, 11) = Sheets("GL5_Master").Cells(i, 8)
Sheets("parts_station").Cells(i - 5, 13) = Sheets("GL5_Master").Cells(i, 10)
Sheets("parts_station").Cells(i - 5, 14) = Sheets("GL5_Master").Cells(i, 11)
Sheets("parts_station").Cells(i - 5, 15) = Sheets("GL5_Master").Cells(i, 12)
Next i

End Sub
