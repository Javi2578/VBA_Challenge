Option Explicit

Sub homework()
    Const FIRST_DATA_ROW As Integer = 2
    Dim I As Integer
    Dim input_row As Long
    Dim last_data_row As Long
    
    Dim totalvolume As Long
    Dim openprice As Double
    Dim rownumber As Integer
    Dim yearlychange As Double
    Dim closeprice As Double
    Dim percentchange As Double
    Dim currentticker As String
    Dim nextticker As String
    
   
    
    'Define collums names bellow use Range
    
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
    Columns("I:L").AutoFit
    
    ' last cell row will be as follow
    
    last_data_row = Cells(FIRST_DATA_ROW, 1).End(xlDown).Row
   
    
    MsgBox " Finish"
    
    'initialize variables; open price, total volume (0), summary row no (summary
    totalvolume = 0
    openprice = Cells(2, 3).Value
    
    rownumber = 2
    
    
    
    
    
    ' HOW IM GOING TO LOOK AT ALL THE DATA AND COMPARE IT
    For I = 2 To last_data_row
        
        currentticker = Cells(I, 1).Value
        nextticker = Cells(I + 1, 1).Value
        
        
        If currentticker = nextticker Then
            ' calculate total volume
            totalvolume = Cells(I, 7).Value + totalvolume
            
            
    
    
        Else
        
            closeprice = Cells(I, 6).Value
            'assign the close price to current row
            
            'calculate yearly change ( close price - open price )
             yearlychange = closeprice - openprice
            
        
            'CALculate percent change
            
            percentchange = yearlychange / openprice
            'output result to summary rows
            
            Cells(rownumber, 9).Value = currentticker
            Cells(rownumber, 10).Value = yearlychange
            Cells(rownumber, 11).Value = percentchange
            Cells(rownumber, 12).Value = totalvolume
            'assign the open price to next row
            openprice = Cells(I + 1, 3).Value
            'reset total volume to zero
            
            totalvolume = 0
            
        End If
        
    
    Next I
    
End Sub

