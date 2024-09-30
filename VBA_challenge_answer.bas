Attribute VB_Name = "Module1"
Sub stock()

' Declare Variables
Dim i As Long
Dim ticker As String
Dim volumeTotal As Double
Dim LR As Long
Dim stockSummaryTable As Integer
Dim ws As Worksheet
Dim GTV As Double ' GTV stands for "Greatest Total Volume"
Dim quarterChange As Double
Dim openPrice As Double
Dim closedPrice As Double
Dim increase As Double
Dim decrease As Double
Dim percentChange As Double
Dim greatestIncrease As Double
Dim greatestDecrease As Double



'Loop through the worksheets
For Each ws In Worksheets


' Initialize the varibles before loop if needed
LR = ws.Cells(Rows.Count, 1).End(xlUp).Row
stockSummaryTable = 2
volumeTotal = 0
GTV = 0
greatestIncrease = -1
greatestDecrease = 1


' Start a Loop
For i = 2 To LR


    ' Create an if statement to see if the cell before is the same as the current cell. If they are not then...
    If ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then
    
        ' Log the openPrice for calculations
        openPrice = ws.Cells(i, 3).Value
    
    ' Close the if statement
    End If

    ' Create an if statement to see if the next cell is the same stock or not. If they are not, then...
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
        ' Set the ticker
        ticker = ws.Cells(i, 1).Value
        
        ' Log the closedPrice
        closedPrice = ws.Cells(i, 6).Value
        
        ' Find the quarterChange
        quarterChange = closedPrice - openPrice
                
        ' Find the percentChange
        If openPrice <> 0 Then
    
            percentChange = (closedPrice - openPrice) / openPrice
        
        Else
    
            percentChange = 0

        End If
        
        ' Create an if statement to check if it is more than the greatestIncrease
        If percentChange > greatestIncrease Then
                
            ' Update the value of greaestIncrease
            greatestIncrease = percentChange
                
            ws.Range("Q2").Value = greatestIncrease
            ws.Range("Q2").NumberFormat = "0.00%"
            
            ' Put the corresponding ticker value into place
            ws.Range("P2").Value = ticker
            
        ' Close if statement
        End If
        
        ' Create another if statement to check if percent Change is less than the greatestDecrease
        If percentChange < greatestDecrease Then
            
            ' Update the value of greatestDecrease
            greatestDecrease = percentChange
                
            ws.Range("Q3").Value = greatestDecrease
            ws.Range("Q3").NumberFormat = "0.00%"
            
            ' Put the corresponding ticker value into place
            ws.Range("P3").Value = ticker
                
        ' Close if statement
        End If
            
        
        ' Add the volume value to the total volume variable
        volumeTotal = volumeTotal + ws.Cells(i, 7).Value
        
        ' Create and if statement that finds the greatest total volume before resetting
        If volumeTotal > GTV Then
            
            ' Update the value of the GTV
            GTV = volumeTotal
                
            ws.Range("Q4").Value = GTV
                
            ' Put the corresponding ticker value into place
            ws.Range("P4").Value = ticker
                
        ' Close the if statement
        End If
        
        ' Print ticker to the stockSummaryTable
        ws.Range("I" & stockSummaryTable).Value = ticker
        
        ' Print the quarterly change to the stockSummaryTable
        ws.Range("J" & stockSummaryTable).Value = quarterChange
            
            ' Create an ElseIf statement to color-code the Quarterly Change row
            If quarterChange = 0 Then
            
                ' Make the background color white if it is equal to 0
                ws.Range("J" & stockSummaryTable).Interior.ColorIndex = 2
            
            ElseIf quarterChange < 0 Then
            
                ' Make the background color red if it is smaller than 0
                ws.Range("J" & stockSummaryTable).Interior.ColorIndex = 3
                
            ElseIf quarterChange > 0 Then
            
                ' Make the background color green if it is bigger than 0
                ws.Range("J" & stockSummaryTable).Interior.ColorIndex = 4
                
            ' Close the if statement
            End If
        
        ' Print the percentChange to the stockSummaryTable
        ws.Range("K" & stockSummaryTable).Value = percentChange
        ws.Range("K" & stockSummaryTable).NumberFormat = "0.00%"
        
        ' Print the volumeTotal to the the stockSummaryTable
        ws.Range("L" & stockSummaryTable).Value = volumeTotal
        
        ' Add one to the stockSummaryTable so it moves on to the next row
        stockSummaryTable = stockSummaryTable + 1
        
        ' Reset the volumeTotal counter
        volumeTotal = 0
        
    ' If the next cell is the same stock then...
    Else
    
        ' Add the volume to volumeTotal
        volumeTotal = volumeTotal + ws.Cells(i, 7).Value
        
    ' Close the if statement
    End If

' Go to the next i
Next i

' Add other formating things you need before moving on to the next worksheet
' Add the headers
ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Quartly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"
ws.Range("O2").Value = "Greatest % Increase"
ws.Range("O3").Value = "Greates % Decrease"
ws.Range("O4").Value = "Greatest Total Volume"
ws.Range("P1").Value = "Ticker"
ws.Range("Q1").Value = "Value"

' Make the columns auto fit
ws.Columns("I:Q").AutoFit

' Move on to the next worksheet in the workbook
Next ws

End Sub

