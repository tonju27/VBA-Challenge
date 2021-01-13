Attribute VB_Name = "Module1"
Sub stockticker():


' Loop through all sheets (Adopted from Bootcamp Activities)
    For Each ws In Worksheets

        ' Find the last row of each worksheet
        ' Subtract one to return the number of rows without header
        lastRow = ws.Cells(Rows.Count, "A").End(xlUp).Row - 1

 
Dim openprice As Double
Dim closeprice As Double
Dim ticker As String
Dim yearlychange As Double
Dim percentchange As Double
Dim totalvolume As Double
Dim currentticker As String
Dim nextticker As String
Dim volume As Double
Dim summaryrow As Integer
Dim counter As Integer
Dim greatestincrease As Double
Dim greatestincreaseticker As String
Dim greatestdecrease As Double
Dim greatestdecreaseticker As String
Dim greatestvolume As Double
Dim greatestvolumeticker As String



yearlychange = 0
percentchange = 0
greatestincrease = 0
greatestdecrease = 0
greatestvolume = 0
volume = 0
totalvolume = 0
summaryrow = 2
ticker = Cells(2, 1).Value
counter = 0
ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Stock Volume"
ws.Cells(1, 15).Value = "Ticker"
ws.Cells(1, 16).Value = "Value"
ws.Cells(2, 14).Value = "Greatest % increase:"
ws.Cells(3, 14).Value = "Greatest % decrease:"
ws.Cells(4, 14).Value = "Greatest volume:"



    For i = 2 To lastRow
        
        currentticker = ws.Cells(i, 1).Value
        nextticker = ws.Cells(i + 1, 1).Value
        
        ' Assign first opening price of ticker to openprice
        
        If counter = 0 Then
            openprice = ws.Cells(i, 3).Value
            counter = counter + 1
        End If
        
         
        If currentticker = nextticker Then
        
            
            volume = ws.Cells(i, 7).Value
            totalvolume = totalvolume + volume
        Else
            ' Add last row of ticker to summary variables
            
            closeprice = ws.Cells(i, 6).Value
            yearlychange = closeprice - openprice
            
            If openprice <> 0 Then
                percentchange = ((closeprice / openprice) - 1)
            End If
            
            
            volume = ws.Cells(i, 7).Value
            totalvolume = totalvolume + volume

            ' Print ticker values to summary table
            
            ws.Cells(summaryrow, 9).Value = ticker
            ws.Cells(summaryrow, 10).Value = yearlychange
            ws.Cells(summaryrow, 11).Value = percentchange
            ws.Cells(summaryrow, 12).Value = totalvolume
            
            ' Identify greatest changes
            
            If percentchange > greatestincrease Then
                greatestincrease = percentchange
               greatestincreaseticker = ticker
            End If
            
            If percentchange < greatestdecrease Then
                greatestdecrease = percentchange
                greatestdecreaseticker = ticker
            End If
            
            If totalvolume > greatestvolume Then
                greatestvolume = totalvolume
                greatestvolumeticker = ticker
            End If

                   
            ' initialize variables
            yearlychange = 0
            percentchange = 0
            totalvolume = 0
            counter = 0
            summaryrow = summaryrow + 1
            ticker = nextticker
            
        End If
    Next i
    
    ' Print greatest changes
    ws.Cells(2, 15).Value = greatestincreaseticker
    ws.Cells(2, 16).Value = greatestincrease
    ws.Cells(3, 15).Value = greatestdecreaseticker
    ws.Cells(3, 16).Value = greatestdecrease
    ws.Cells(4, 15).Value = greatestvolumeticker
    ws.Cells(4, 16).Value = greatestvolume
    

    Next ws
    
       
End Sub


