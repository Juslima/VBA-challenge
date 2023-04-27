# VBA-challenge
Sub analyzeStockMarketData()
    
    ' Define variables
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim ticker As String
    Dim openingPrice As Double
    Dim closingPrice As Double
    Dim yearlyChange As Double
    Dim percentChange As Double
    Dim totalVolume As Double
    Dim summaryTableRowIndex As Long
    
    ' Loop through all worksheets in the workbook
    For Each ws In ThisWorkbook.Worksheets
        
        summaryTableRowIndex = 2
        
        ' Column Creation
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        
        ' Get the last row of data in the worksheet
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        ' Loop through all rows of data in the worksheet
        For i = 2 To lastRow
            
            If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
                
                openingPrice = ws.Cells(i, 3).Value
                totalVolume = ws.Cells(i, 7).Value
            Else
               
                totalVolume = totalVolume + ws.Cells(i, 7).Value
            End If
            
           
            If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
                
                closingPrice = ws.Cells(i, 6).Value
                yearlyChange = closingPrice - openingPrice
                If openingPrice <> 0 Then
                    percentChange = yearlyChange / openingPrice
                Else
                    percentChange = 0
                End If
                
                ' Conditional Formatting
                 If yearlyChange > 0 Then
                        ws.Range("J2:J3001").Interior.ColorIndex = 4
                    ElseIf yearlyChange < 0 Then
                        ws.Range("J2:J3001").Interior.ColorIndex = 3
                    End If
                    
                    If percentChange > 0 Then
                        ws.Range("k2:k3001").Interior.ColorIndex = 4
                    ElseIf yearlyChange < 0 Then
                        ws.Range("k2:k3001").Interior.ColorIndex = 3
                    End If
                    
               
                ws.Range("I" & summaryTableRowIndex).Value = ws.Cells(i, 1).Value
                ws.Range("J" & summaryTableRowIndex).Value = yearlyChange
                ws.Range("J" & summaryTableRowIndex).NumberFormat = "#0.00"
                ws.Range("K" & summaryTableRowIndex).Value = percentChange
                ws.Range("K" & summaryTableRowIndex).NumberFormat = "0.00%"
                ws.Range("L" & summaryTableRowIndex).Value = totalVolume
                
               
                summaryTableRowIndex = summaryTableRowIndex + 1
            End If
            
        Next i
        
    Next ws
    
End Sub


----

Sub Conditional_Formatting()
'
' Conditional_Formatting Macro
'
' Keyboard Shortcut: Ctrl+n
'
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, _
        Formula1:="=0"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 255
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, _
        Formula1:="=0"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 52224
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
End Sub
----

Sub Greatest()
'
' Greatest Macro
'
' Keyboard Shortcut: Ctrl+l
'
    Range("Q2").Select
    ActiveCell.FormulaR1C1 = "=MAX(C[-6])"
    Range("Q2:Q3").Select
    Selection.NumberFormat = "0.00%"
    Range("Q3").Select
    ActiveCell.FormulaR1C1 = "=MIN(C[-6])"
    Range("Q4").Select
    ActiveCell.FormulaR1C1 = "=MAX(C[-5])"
    Range("Q4").Select
    Selection.NumberFormat = "0.000000E+00"
    Selection.NumberFormat = "0.00000E+00"
    Selection.NumberFormat = "0.0000E+00"
    Selection.NumberFormat = "0.000E+00"
    Selection.NumberFormat = "0.00E+00"
    Range("P6").Select
End Sub
Sub Greatest_Ticker()
'
' Greatest_Ticker Macro
'

'
    ActiveCell.FormulaR1C1 = "=INDEX(C[-7],MATCH(MAX(C[-5]),C[-5],0))"
    Range("P3").Select
    ActiveCell.FormulaR1C1 = "=INDEX(C[-7],MATCH(MIN(C[-5]),C[-5],0))"
    Range("P4").Select
    ActiveCell.FormulaR1C1 = "=INDEX(C[-7],MATCH(MAX(C[-4]),C[-4],0))"
    Range("P5").Select
End Sub
