Sub HW()

For Each ws In Worksheets
    ws.Activate
    
'Declarations
Dim ticker, resumeticker, tickinc, tickertt, tickdeac As String
Dim summary As Integer
Dim lr, totalstock, tostrv, grtsv As LongLong
Dim greatinc, greatdec, greattval, perchange, openprice, closeprice, yrch As Double

'Headers
Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Stock Volume"
Cells(2, 15).Value = "Greatest % Increase"
Cells(3, 15).Value = "Greatest % Decrease"
Cells(4, 15).Value = "Greatest Total Volume"
Cells(1, 16).Value = "Ticker"
Cells(1, 17).Value = "Value"

'Here we go
lr = Cells(Rows.Count, 1).End(xlUp).Row
summary = 2
openprice = Cells(2, 3).Value


For b = 2 To lr

  If Cells(b, 1).Value <> Cells(b - 1, 1).Value Then
     openprice = Cells(b, 3).Value
  End If

  If Cells(b, 1).Value <> Cells(b + 1, 1).Value Then
    
    ticker = Cells(b, 1).Value
    closeprice = Cells(b, 6).Value
    yearchange = closeprice - openprice
        
        If openprice = 0 Then
            perchange = 0
        Else
          perchange = yearchange / openprice
        End If
        
        If yearchange > 0 Then
        Cells(summary, 10).Interior.ColorIndex = 4
        ElseIf yearchange < 0 Then
        Cells(summary, 10).Interior.ColorIndex = 3
          
     End If
     
        
        
    Cells(summary, 10).Value = yearchange
    Cells(summary, 9).Value = ticker
    Cells(summary, 11).Value = perchange
    Cells(summary, 11).NumberFormat = "0.00%"
    totalstock = totalstock + Cells(b, 7).Value
    Cells(summary, 12).Value = totalstock
    
    summary = summary + 1
    totalstock = 0
    openprice = Cells(b + 1, 3).Value
    
    
    Else
    
    totalstock = totalstock + Cells(b, 7).Value
    
         
End If

  
Next b


'Now the challenge
lrs = Cells(Rows.Count, 9).End(xlUp).Row

greatinc = 0
greatdec = 0

    For r = 2 To lrs
    
        resumeticker = Cells(r, 9).Value
        yrch = Cells(r, 11).Value
        tostrv = Cells(r, 12).Value
                
        If yrch > greatinc Then
            greatinc = yrch
            tickinc = resumeticker
        End If
          
        If yrch < greatdec Then
            greatdec = yrch
            tickdeac = resumeticker
        End If
        
        If tostrv > grtsv Then
            grtsv = tostrv
            tickertt = resumeticker
        End If
        
        
        
    
    
Next r
 
      Cells(2, 17).Value = greatinc
      Cells(2, 17).NumberFormat = "0.00%"
      Cells(3, 17).Value = greatdec
      Cells(3, 17).NumberFormat = "0.00%"
      Cells(4, 17).Value = grtsv
      Cells(2, 16).Value = tickinc
      Cells(3, 16).Value = tickdeac
      Cells(4, 16).Value = tickertt
      
      
      
      
Next ws


 
End Sub