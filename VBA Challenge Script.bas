Attribute VB_Name = "Module1"
Sub Challenge()

'Look through all sheets
For Each ws In Worksheets

    'Define variables
    Dim tickername As String
    Dim openamount As Double
    Dim closeamount As Double
    Dim stock As Double
    Dim tickersum As Integer
    Dim yearlychange As Double
    Dim percentchange As Double
    Dim strtOpn As Long
    
    'Bonus Variables
    Dim incticker As String
    Dim incvalue As Double
    Dim decticker As String
    Dim decvalue As Double
    Dim totalticker As String
    Dim totalvalue As Double
       
    tickersum = 2
    strOpn = 2
    
    incvalue = 0
    decvalue = 0
    totalvalue = 0
    
    'label summary table columns
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    ws.Range("Q1").Value = "Ticker"
    ws.Range("R1").Value = "Value"
    ws.Range("P2").Value = "Greatest % Increase"
    ws.Range("P3").Value = "Greatest % Decrease"
    ws.Range("P4").Value = "Greatest Total Volume"
    
    'determine last row
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
   
      
       
    'find out where ticker changes
    For i = 2 To lastrow
    
              
        If (ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1)) Then
            
            'copy across ticker name
            tickername = ws.Cells(i, 1).Value
            ws.Cells(tickersum, 9).Value = tickername
                       
            'open amount
            openamount = ws.Cells(strOpn, 3).Value
            
            'close amount
            closeamount = ws.Cells(i, 6).Value
            
            '__________________________________________
            
            'YEARLY CHANGE
            '__________________________________________
            
            'work out yearly change
            yearlychange = closeamount - openamount
            
            'print yearly change
            ws.Cells(tickersum, 10).Value = yearlychange
            
            
            'conditional format yearly change column
            If yearlychange >= 0 Then
                
                ws.Cells(tickersum, 10).Interior.ColorIndex = 4
                
            Else
                                        
                ws.Cells(tickersum, 10).Interior.ColorIndex = 3
                
            End If
                        
            '_____________________________________________
            
            'PERCENTAGE CHANGE
            '_____________________________________________
            'calculate percentage change
             percentagechange = yearlychange / openamount
             
             ' start of the next stock ticker
             strOpn = i + 1
             
            
                        
            'print percentage change
            ws.Cells(tickersum, 11).Value = percentagechange
            
               'format percentage column
            ws.Range("K2:K" & lastrow).NumberFormat = "0.00%"
            
                        
            '_____________________________________________
            
            'STOCK VOLUME
            '_____________________________________________
            'add up stock volume
            stock = stock + ws.Cells(i, 7).Value
            
            'print stock volume in summary table
            ws.Cells(tickersum, 12).Value = stock
            
            'move ticket summary down a row
            tickersum = tickersum + 1
            
            'Reset value to 0
            openamount = 0
            closeamount = 0
            stock = 0
            
        Else
     
        
        'add up stock volume
        stock = stock + ws.Cells(i, 7).Value
        
        
        End If
        
        '_____________________________________________
        
        'BONUS
        '_____________________________________________
        
        'find greatest % increase
        If (ws.Cells(i, 11).Value > incvalue) Then
            
            incvalue = ws.Cells(i, 11).Value
            incticker = ws.Cells(i, 9).Value
            
        'find greatest % decrease
        ElseIf (ws.Cells(i, 11).Value < decvalue) Then
            
            decvalue = ws.Cells(i, 11).Value
            decticker = ws.Cells(i, 9).Value
            
        'find greatest total stock
        ElseIf (ws.Cells(i, 12).Value > totalvalue) Then
            
            totalvalue = ws.Cells(i, 12).Value
            totalticker = ws.Cells(i, 9).Value
        
        End If
      
      
    Next i
    
    'print greatest % increase
    ws.Range("Q2") = incticker
    ws.Range("R2") = incvalue
    
    'print greatest % decrease
    ws.Range("Q3") = decticker
    ws.Range("R3") = decvalue
    
    'print total stock
    ws.Range("Q4") = totalticker
    ws.Range("R4") = totalvalue
    
    
    'Format columns
    ws.Columns("A:R").AutoFit
    ws.Range("R2:R3").NumberFormat = "0.00%"
    

   
Next ws


End Sub
