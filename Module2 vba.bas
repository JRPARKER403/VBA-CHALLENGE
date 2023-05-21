Attribute VB_Name = "Module2"
Sub alphabetical_testing()

Dim total As Double
Dim j As Long
j = 2

total = 0

Dim rowcount As Double

Range("I1").Value = "ticker"
Range("L1").Value = "Stock Volume"
Range("J1").Value = "Yearly Change"
Range("K1").Value = "Percent Change"

rowcount = Cells(Rows.Count, "A").End(xlUp).Row

For i = 2 To rowcount
    
    If Cells(i + 1, 1) <> Cells(i, 1) Then
    
        total = total + Cells(i, 7).Value
        Range("I" & j).Value = Cells(i, 1).Value
        Range("L" & j).Value = total
        j = j + 1
        total = 0
        
    Else
        total = total + Cells(i, 7).Value
    End If
Next i

End Sub


Sub PercentChange()
 Dim oldValue As Double
 Dim newValue As Double
 oldValue = Range("C2").Value
 newValue = Range("F2").Value 'replace with cell containing new value
 Range("A3").Value = PercentageChange(oldValue, newValue)
End Sub
    
'Add a sheet named "Combined Data"
Sheets.Add.Name = "Combined Data"
'move created sheet to be first sheet
Sheets("Combined Data").Move Before:=Sheets(1)
'Specify the location of the combined sheet
Set Combined_sheet = Worksheets("Combiined_Data")

'Loop through all sheets
For Each ws In Worksheets

'Find the last row of the combined sheet
'Add 1 to get first empty row
Lastrow = Combined_sheet.Cells(Rows.Count, "A").End(xlUp).Row + 1

'Find the last row of each worksheet
'Subtract 1 to return the numberof rows without header
lastRowYear = ws.Cells(Rows.Count, "A").End(xlUp).Row - 1

'Copy contents of each year sheet into the combine sheet
Combined_sheet.Range("A" & astRow & ":L" & ((lastRowYear - 1) + Lastrow)).Value = ws.Range("A2:L" & (lastRowYear + 1)).Value

Next ws
'Copy the headers from sheet1
Combined_sheet.Range("A1:L1").Value = Sheets(2).Range("A1:L1").Value = Sheets(2).Range("A1;L1").Value

'Autofit to display data
Combined_sheet.Columns("A:L").AutoFit

End Sub

