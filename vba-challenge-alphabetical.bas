Attribute VB_Name = "Module1"
Option Explicit
Sub stock_data()

        Dim ws As Worksheet
        Dim open_price As Double
        Dim close_price As Double
        Dim yearly_change As Double
        Dim change_ratio As Double
        Dim last_row As Long
        Dim output_row As Long
        Dim input_row As Long
        
        
        Dim ticker As String
        Dim total_volume As LongLong
        Dim max_ticker As String
        Dim min_ticker As String
        Dim greatest_total_name As String
        Dim greatest_increase As Double
        Dim greatest_decrease As Double
        Dim greatest_total_vol As LongLong
        
        
        
        For Each ws In Worksheets
        
            ws.Activate
        
            Range("I1").Value = "ticker"
            Range("J1").Value = "Yearly Change"
            Range("K1").Value = "Percent Change"
            Range("L1").Value = "Total Stock Volume"
            
            Range("P2").Value = "Greatest % Increase"
            Range("P3").Value = "Greatest % Decrease"
            Range("P4").Value = "Greatest Total Volume"
            Range("R1").Value = "Ticker"
            Range("S1").Value = "Value"
            
            
            'prepare for first stock
            total_volume = 0
            output_row = 2
            open_price = Cells(2, 3).Value
            
            greatest_increase = 0
            greatest_decrease = 0
            greatest_total_vol = 0
            
            last_row = Cells(Rows.Count, 1).End(xlUp).Row
            
            'variables
            
            For input_row = 2 To last_row
                ticker = Cells(input_row, 1).Value
                'volume
                total_volume = Cells(input_row, 7).Value
                
                'Last day of current stock
                If Cells(input_row + 1, 1).Value <> ticker Then
                    'input
                    close_price = Cells(input_row, 6).Value
                    total_volume = total_volume + Cells(input_row, 7).Value
                    
                    'calculations
                    yearly_change = close_price - open_price
                    change_ratio = yearly_change / open_price
                    'greatest vol
                    
                    'output
                    Cells(output_row, "I").Value = ticker
                    Cells(output_row, "L").Value = total_volume
                    
                    Cells(output_row, "J").Value = yearly_change
                    If yearly_change < 0 Then
                        Cells(output_row, "J").Interior.ColorIndex = 3
                    Else
                        Cells(output_row, "J").Interior.ColorIndex = 4
                    End If
                        Cells(output_row, "K").Value = FormatPercent(change_ratio)
                    
                    'prepare for the next stock
                    output_row = output_row + 1
                    open_price = Cells(input_row + 1, 3).Value
                    total_volume = 0
                     
                End If
                
              Next input_row
              
                'calculate min and max values
                Dim last_new_row As Long
                last_new_row = Range("K2").End(xlDown).Row
                max_ticker = " "
                min_ticker = " "
                greatest_total_name = " "
                greatest_increase = 0
                greatest_decrease = 0
                greatest_total_vol = 0
                
                Dim i As Integer
                
                
                For i = 2 To last_new_row
                
                    If Cells(i, "K").Value > greatest_increase Then
                        greatest_increase = Cells(i, "K").Value
                        Cells(2, "S").Value = greatest_increase
                        max_ticker = Cells(i, "I").Value
                        Cells(2, "R").Value = max_ticker
                    End If
                    
                    If Cells(i, "K").Value < greatest_decrease Then
                        greatest_decrease = Cells(i, "K").Value
                        Cells(3, "S").Value = greatest_decrease
                        min_ticker = Cells(i, "I").Value
                        Cells(3, "R").Value = min_ticker
                    End If
                    
                    If Cells(i, "L").Value > greatest_total_vol Then
                        greatest_total_vol = Cells(i, "L").Value
                        Cells(4, "S").Value = greatest_total_vol
                        greatest_total_name = Cells(i, "I").Value
                        Cells(4, "R").Value = greatest_total_name
                    End If
                        
                Next i
                
    
        Next ws
        
        MsgBox ("Done")
End Sub



