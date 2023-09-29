Attribute VB_Name = "Module1"
Option Explicit

Sub Stock_analysis()

Dim i, j, openIdx As Integer
Dim vol, rowCount As Long

' Naming the headers

Range("I1").Value = "Ticker"
Range("J1").Value = "Yearly Change"
Range("K1").Value = "Percent Change"
Range("L1").Value = "Total Stock Volume"
Range("P1").Value = "Ticker"
Range("Q1").Value = "Value"
Range("O2").Value = "Greatest % Increase"
Range("O3").Value = "Greatest % Decrease"
Range("O4").Value = "Greatest Total Volume"

vol = 0
j = 0

rowCount = Cells(Rows.Count, "A").End(xlUp).Row

For i = 2 To rowCount

    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        vol = vol + Cells(i, 7).Value
        
        Range("I" & 2 + j).Value = Cells(i, 1).Value
        Range("L" & 2 + j).Value = vol
        
        vol = 0
        j = j + 1
        
    Else
        
        vol = vol + Cells(i, 7).Value
        
    End If

Next i


End Sub

'Changing the price

Sub Stock_analyis2()

Dim i, j, openstock As Integer
Dim vol, rowCount As Long
Dim Delta As Double

' Naming the headers
Range("I1").Value = "Ticker"
Range("J1").Value = "Yearly Change"
Range("K1").Value = "Percent Change"
Range("L1").Value = "Total Stock Volume"
Range("P1").Value = "Ticker"
Range("Q1").Value = "Value"
Range("O2").Value = "Greatest % Increase"
Range("O3").Value = "Greatest % Decrease"
Range("O4").Value = "Greatest Total Volume"

vol = 0
j = 0
openstock = 2

rowCount = Cells(Rows.Count, "A").End(xlUp).Row

For i = 2 To rowCount

    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        vol = vol + Cells(i, 7).Value
        Range("I" & 2 + j).Value = Cells(i, 1).Value
        Range("L" & 2 + j).Value = vol
        
        Delta = Cells(i, 6).Value - Cells(openstock, 3).Value
        Range("J" & 2 + j).Value = Delta
       
        
        vol = 0
        Delta = 0
        j = j + 1
        openstock = openstock + 1
        
    Else
        
        vol = vol + Cells(i, 7).Value
        
    End If

Next i


End Sub

' % change

Sub Stock_analysis3()

Dim i, j, openstock As Integer
Dim vol, rowCount As Long
Dim Change, percChange As Double


Range("I1").Value = "Ticker"
Range("J1").Value = "Yearly Change"
Range("K1").Value = "Percent Change"
Range("L1").Value = "Total Stock Volume"
Range("P1").Value = "Ticker"
Range("Q1").Value = "Value"
Range("O2").Value = "Greatest % Increase"
Range("O3").Value = "Greatest % Decrease"
Range("O4").Value = "Greatest Total Volume"

vol = 0
j = 0
openstock = 2

rowCount = Cells(Rows.Count, "A").End(xlUp).Row

For i = 2 To rowCount

    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        vol = vol + Cells(i, 7).Value
        Range("I" & 2 + j).Value = Cells(i, 1).Value
        Range("L" & 2 + j).Value = vol
        
        Change = Cells(i, 6).Value - Cells(openstock, 3).Value
        Range("J" & 2 + j).Value = Change
        
        percChange = Change / Cells(openstock, 3).Value
        Range("K" & 2 + j).Value = percChange
        
        
        vol = 0
        Change = 0
        percChange = 0
        j = j + 1
        openstock = openstock + 1
        
    Else
        
    vol = vol + Cells(i, 7).Value
        
    End If

Next i


End Sub

' Coloring the cells

Sub Stock_analysis4()

Dim i, j, openstock As Integer
Dim vol, rowCount As Long
Dim Change, percChange As Double

' Naming the headers

Range("I1").Value = "Ticker"
Range("J1").Value = "Yearly Change"
Range("K1").Value = "Percent Change"
Range("L1").Value = "Total Stock Volume"
Range("P1").Value = "Ticker"
Range("Q1").Value = "Value"
Range("O2").Value = "Greatest % Increase"
Range("O3").Value = "Greatest % Decrease"
Range("O4").Value = "Greatest Total Volume"

vol = 0
j = 0
openstock = 2

rowCount = Cells(Rows.Count, "A").End(xlUp).Row

For i = 2 To rowCount

    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        vol = vol + Cells(i, 7).Value
        Range("I" & 2 + j).Value = Cells(i, 1).Value
        Range("L" & 2 + j).Value = vol
        
        Change = Cells(i, 6).Value - Cells(openstock, 3).Value
        Range("J" & 2 + j).Value = Change
        Range("J" & 2 + j).NumberFormat = "0.00"
        
        If Change > 0 Then
            Range("J" & 2 + j).Interior.ColorIndex = 4
        ElseIf Change < 0 Then
            Range("J" & 2 + j).Interior.ColorIndex = 3
        Else
            Range("J" & 2 + j).Interior.ColorIndex = 0
        End If
            
        
        percChange = Change / Cells(openstock, 3).Value
        Range("K" & 2 + j).Value = percChange
        
        
        vol = 0
        Change = 0
        percChange = 0
        j = j + 1
        openstock = openstock + 1
        
    Else
        
        vol = vol + Cells(i, 7).Value
        
    End If

Next i

End Sub

' Broad metrics based on values above

Sub Stock_analysis5()

Dim i, j, openstock As Integer
Dim vol, rowCount, rowCountK, rowCountL As Long
Dim Change, percChange As Double
Dim beststock, worststock, moststock As Integer

' Naming the headers

Range("I1").Value = "Ticker"
Range("J1").Value = "Yearly Change"
Range("K1").Value = "Percent Change"
Range("L1").Value = "Total Stock Volume"
Range("P1").Value = "Ticker"
Range("Q1").Value = "Value"
Range("O2").Value = "Greatest % Increase"
Range("O3").Value = "Greatest % Decrease"
Range("O4").Value = "Greatest Total Volume"

vol = 0
j = 0
openstock = 2

rowCount = Cells(Rows.Count, "A").End(xlUp).Row

For i = 2 To rowCount

    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        vol = vol + Cells(i, 7).Value
        Range("I" & 2 + j).Value = Cells(i, 1).Value
        Range("L" & 2 + j).Value = vol
        
        Change = Cells(i, 6).Value - Cells(openstock, 3).Value
        Range("J" & 2 + j).Value = Change
      
        If Change > 0 Then
            Range("J" & 2 + j).Interior.ColorIndex = 4
        ElseIf Change < 0 Then
            Range("J" & 2 + j).Interior.ColorIndex = 3
        Else
            Range("J" & 2 + j).Interior.ColorIndex = 0
        End If
            
        
        percChange = Change / Cells(openstock, 3).Value
        Range("K" & 2 + j).Value = percChange
        
        vol = 0
        Change = 0
        percChange = 0
        j = j + 1
        openstock = openstock + 1
        
    Else
        
        vol = vol + Cells(i, 7).Value
        
    End If
    
Next i

rowCountK = Cells(Rows.Count, "K").End(xlUp).Row
rowCountL = Cells(Rows.Count, "L").End(xlUp).Row

' stocks that had the biggest % changes in price and volume
Range("Q2") = WorksheetFunction.Max(Range("K2:K" & rowCountK).Value)
Range("Q3") = WorksheetFunction.Min(Range("K2:K" & rowCountK).Value)
Range("Q4") = WorksheetFunction.Max(Range("L2:L" & rowCountL).Value)

' Use MATCH to find the index of the best, worst and most
beststock = WorksheetFunction.Match(WorksheetFunction.Max(Range("K2:K" & rowCountK).Value), Range("K2:K" & rowCountK).Value, 0)
worststock = WorksheetFunction.Match(WorksheetFunction.Min(Range("K2:K" & rowCountK).Value), Range("K2:K" & rowCountK).Value, 0)
moststock = WorksheetFunction.Match(WorksheetFunction.Max(Range("L2:L" & rowCountL).Value), Range("L2:L" & rowCountL).Value, 0)

' Plug the values into the required cells
Range("P2").Value = Cells(beststock + 1, 9).Value

Range("P3").Value = Cells(worststock + 1, 9).Value

Range("P4").Value = Cells(moststock + 1, 9).Value

End Sub

' Run across all sheets

Sub Stock_analysis6()

Dim ws As Object

For Each ws In Worksheets

Dim i, j, openstock As Integer
Dim vol, rowCount, rowCountK, rowCountL As Long
Dim Change, percChange As Double
Dim beststock, worststock, moststock As Integer


' Naming the headers
ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"
ws.Range("P1").Value = "Ticker"
ws.Range("Q1").Value = "Value"
ws.Range("O2").Value = "Greatest % Increase"
ws.Range("O3").Value = "Greatest % Decrease"
ws.Range("O4").Value = "Greatest Total Volume"

vol = 0
j = 0
openstock = 2


rowCount = ws.Cells(Rows.Count, "A").End(xlUp).Row

For i = 2 To rowCount

    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        vol = vol + ws.Cells(i, 7).Value
        ws.Range("I" & 2 + j).Value = ws.Cells(i, 1).Value
        ws.Range("L" & 2 + j).Value = vol
        
        Change = ws.Cells(i, 6).Value - ws.Cells(openstock, 3).Value
        ws.Range("J" & 2 + j).Value = Change
        
        If Change > 0 Then
            ws.Range("J" & 2 + j).Interior.ColorIndex = 4
        ElseIf Change < 0 Then
            ws.Range("J" & 2 + j).Interior.ColorIndex = 3
        Else
            ws.Range("J" & 2 + j).Interior.ColorIndex = 0
        End If
            
        
        percChange = Change / ws.Cells(openstock, 3).Value
        ws.Range("K" & 2 + j).Value = percChange
        
        vol = 0
        Change = 0
        percChange = 0
        j = j + 1
        openstock = openstock + 1
        
    Else
        
        vol = vol + ws.Cells(i, 7).Value
        
    End If
    
Next i

rowCountK = ws.Cells(Rows.Count, "K").End(xlUp).Row
rowCountL = ws.Cells(Rows.Count, "L").End(xlUp).Row

' Compute the stocks that had the biggest % changes in price and volume
ws.Range("Q2").Value = WorksheetFunction.Max(ws.Range("K2:K" & rowCountK).Value)
ws.Range("Q2").NumberFormat = "0.00%"

ws.Range("Q3").Value = WorksheetFunction.Min(ws.Range("K2:K" & rowCountK).Value)
ws.Range("Q3").NumberFormat = "0.00%"
ws.Range("Q4").Value = WorksheetFunction.Max(ws.Range("L2:L" & rowCountL).Value)

' Use MATCH to find the index of the best, worst and most
beststock = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("K2:K" & rowCountK).Value), ws.Range("K2:K" & rowCountK).Value, 0)
worststock = WorksheetFunction.Match(WorksheetFunction.Min(ws.Range("K2:K" & rowCountK).Value), ws.Range("K2:K" & rowCountK).Value, 0)
moststock = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("L2:L" & rowCountL).Value), ws.Range("L2:L" & rowCountL).Value, 0)

' Change format of column K to reflect percentage values
ws.Range("K2:K" & rowCountK).NumberFormat = "0.00%"

' Plug the values into the required cells
ws.Range("P2").Value = ws.Cells(beststock + 1, 9).Value

ws.Range("P3").Value = ws.Cells(worststock + 1, 9).Value

ws.Range("P4").Value = ws.Cells(moststock + 1, 9).Value


Next ws

End Sub

