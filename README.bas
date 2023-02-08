Attribute VB_Name = "Module1"

Sub Stock_Analysis()
  Dim ws As Worksheet
  Dim ticker As String
  Dim year_open As Double
  Dim year_close As Double
  Dim yearly_change As Double
  Dim pct_change As Double
  Dim total_vol As Double
  Dim last_row As Long
  Dim i As Long
  Dim result_row As Long

  For Each ws In ThisWorkbook.Sheets
    last_row = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    result_row = 2
    
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    
    For i = 2 To last_row
      ticker = ws.Cells(i, "A").Value
      year_open = ws.Cells(i, "C").Value
      total_vol = total_vol + ws.Cells(i, "G").Value
      
      If ws.Cells(i + 1, "A").Value <> ticker Or i = last_row Then
        year_close = ws.Cells(i, "E").Value
        yearly_change = year_close - year_open
        If year_open <> 0 Then
          pct_change = yearly_change / year_open
        Else
          pct_change = 0
        End If
        
        ws.Range("I" & result_row).Value = ticker
        ws.Range("J" & result_row).Value = yearly_change
        ws.Range("K" & result_row).Value = pct_change
        ws.Range("L" & result_row).Value = total_vol
        
        ws.Range("J" & result_row).NumberFormat = "0.00"
        ws.Range("K" & result_row).NumberFormat = "0.00%"
        
        If yearly_change >= 0 Then
          ws.Range("J" & result_row).Interior.ColorIndex = 4
        Else
          ws.Range("J" & result_row).Interior.ColorIndex = 3
        End If
        
        result_row = result_row + 1
        total_vol = 0
      End If
    Next i
  Next ws
End Sub


