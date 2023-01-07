Attribute VB_Name = "Module1"
Sub Tickers()

'Create Variables
    'Ticker Symbol
Dim Ticker_Name As String
Dim Ticker_Number As Double
    'Opening Price
Dim Open_Price As Double
    'Closing Price
Dim Close_Price As Double
    'Yearly Change
Dim Annual_Change As Double
    'Percent Change
Dim Percent_Change As Double
    'Volume
Dim Total_Volume As Double

    'Set to run through all sheets

For Each ws In Worksheets
ws.Activate
    
    'Create a Summary Table

Range("K1").Value = "Ticker"
Range("L1").Value = "Annual Change"
Range("M1").Value = "Percentage Change"
Range("N1").Value = "Total Volume"


        
    'set last row
lastrow = Cells(Rows.Count, 1).End(xlUp).Row

    'Set Variables to zero
Ticker_Number = 0
Ticker_Name = ""
Open_Price = 0
Annual_Change = 0
Percent_Change = 0
Total_Volume = 0

    'Work Tickers
For i = 2 To lastrow
Ticker = Cells(i, 1).Value

If Open_Price = 0 Then
Open_Price = Cells(i, 3).Value
End If

Total_Volume = Total_Volume + Cells(i, 7).Value

If Cells(i + 1, 1).Value <> Ticker Then
Ticker_Number = Ticker_Number + 1

Cells(Ticker_Number + 1, 11) = Ticker
Close_Price = Cells(i, 6)
Annual_Change = Close_Price - Open_Price
Cells(Ticker_Number + 1, 12).Value = Annual_Change

If Open_Price = 0 Then
Percent_Change = 0
Else
Percent_Change = (Annual_Change / Open_Price)
Cells(Ticker_Number + 1, 13).Value = Percent_Change

If Percent_Change < 0 Then
Cells(Ticker_Number + 1, 13).Interior.Color = vbRed
Else
Cells(Ticker_Number + 1, 13).Interior.Color = vbGreen
End If
'Insert formatting
If Annual_Change < 0 Then
Cells(Ticker_Number + 1, 12).Interior.Color = vbRed
Else
Cells(Ticker_Number + 1, 12).Interior.Color = vbGreen
End If

End If
Cells(Ticker_Number + 1, 14).Value = Total_Volume
Open_Price = 0
Total_Volume = 0

End If
Next i


'New table
'create variables
Dim Greatest_Volume As Double
Dim Greatest_Percent_Up As Double
Dim Greatest_Percent_Down As Double
'Dim maxreference As Range
'Dim percentreference As Range
'Dim minreference As Range


Range("Q1").Value = "Ticker"
Range("R1").Value = "Value"
Range("P2").Value = "Greatest Percentage Increase"
Range("P3").Value = "Greatest Percentage Decrease"
Range("P4").Value = "Greatest Total Volume"

'PRINT TO TABLE WITH TICKER IDENTIFIER
Greatest_Volume = Application.WorksheetFunction.Max(ws.Range("N:N"))
Range("R4").Value = Greatest_Volume
'Set maxreference = ws.Range("N:N").Find(Greatest_Volume, lookat:=xlWhole)
'Range("q4") = maxreference.Offset(0, -3)


Greatest_Percent_Up = Application.WorksheetFunction.Max(ws.Range("M:M"))
Range("R2").Value = Greatest_Percent_Up

'Set percentreference = ws.Range("m:m").Find(Greatest_Percent_Up, lookat:=xlWhole)
'Range("q2") = percentreference.Offset(0, -3)
Greatest_Percent_Down = Application.WorksheetFunction.min(ws.Range("M:M"))
Range("R3").Value = Greatest_Percent_Down
'Set minreference = ws.Range("m:m").Find(Greatest_Percent_Down, lookat:=xlWhole)
'Range("q3") = minreference.Offset(0, -3)

For i = 2 To lastrow

If Cells(i, 13) = Greatest_Percent_Up Then
Range("Q2").Value = Cells(i, 13).Offset(0, -2)
End If
If Cells(i, 13) = Greatest_Percent_Down Then
Range("Q3").Value = Cells(i, 13).Offset(0, -2)
End If
If Cells(i, 14) = Greatest_Volume Then
Range("Q4").Value = Cells(i, 14).Offset(0, -3)
End If


Next i

    'Format Summary Tables
Range("K1:N1").Font.Bold = True
Columns("K:N").Columns.AutoFit
Columns("M:M").NumberFormat = "0.00%"
Range("q1:r1").Font.Bold = True
Range("P:P").Font.Bold = True
Columns("p:r").Columns.AutoFit
Range("R2:R3").NumberFormat = "0.00%"

Next ws
Sheets("2018").Select

    
  
'FIND MIN/MAX ACROSS ALL SHEETS
    
'For J = 1 To Sheets.Count
'Greatest_Volume = WorksheetFunction.Max(Sheets(J).Range("N:N"))
'Range("R4").Value = Greatest_Volume
'Greatest_Percent_Up = WorksheetFunction.Max(Sheets(J).Range("M:M"))
'Range("R2").Value = Greatest_Percent_Up
'Greatest_Percent_Down = WorksheetFunction.Min(Sheets(J).Range("M:M"))
'Range("R3").Value = Greatest_Percent_Down

    

'Next J
  







End Sub


