VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub ticker():
    
    Dim w As Integer
    Dim ws_no As Integer
    Dim ws As Worksheet
    Dim starting_ws As Worksheet
    Dim startingcounter As Integer
    Dim volumecounter As Double
    Dim highesthigh As Double
    Dim lowestlow As Double
    Dim ticketindexcounter As Integer
    Dim openstockprice As Double
    Dim k As Long
    Dim currentticker As String
    Dim highticker As String
    Dim lowticker As String
    Dim bigticker As String
    
    
    Set starting_ws = ActiveSheet
    
    For Each ws In ThisWorkbook.Worksheets
        ws.Activate
        
  
    
    'initialize counters'
    volumecounter = 0
    ticketindexcounter = 2
    
    'some counters can't be empty for the 'not' part of the elseif loop part of the code,
    'so i seed them here for the first entry in case it would land it there initially'
    
    currentticker = Cells(2, 1)
    openstockprice = Cells(2, 3).Value
    
    'find last nonempty row to know how long for loops need to iterate'
    k = Cells(Rows.Count, 1).End(xlUp).Row
    
    For i = 2 To k + 1
    
        If Cells(i, 1).Value = Cells(i - 1, 1).Value Then
        
            'record additional volume'
            volumecounter = volumecounter + Cells(i, 7).Value
            
            'check if this is new highest stock price'
            If highesthigh < Cells(i, 4).Value Then
                highesthigh = Cells(i, 4).Value
                
                End If
            
            'or if this is the new lowest stock price'
            If lowestlow > Cells(i, 5).Value Then
                lowestlow = Cells(i, 5).Value
        
                End If
        
        ElseIf Cells(i, 1).Value <> Cells(i - 1, 1).Value Then
                
              
            'before moving on, record volume for stocks for this ticket..'
            Cells(ticketindexcounter, 12).Value = volumecounter
            
            '..as well as yearly and percent change'
            Cells(ticketindexcounter, 10).Value = openstockprice - Cells(i, 6).Value
            
            Cells(ticketindexcounter, 11) = Cells(ticketindexcounter, 10) * 100 / openstockprice
            
            
            'now, move down one row to enter new data'
            ticketindexcounter = ticketindexcounter + 1
            
            'reset variables'
            volumecounter = 0
            openstockprice = Cells(i, 3)
            highesthigh = Cells(i, 4)
            lowestlow = Cells(i, 5)
            
            'record new ticker name'
            Cells(ticketindexcounter, 9).Value = Cells(i, 1).Value
            
            'add to volumecounter now that it's been reset'
            volumecounter = volumecounter + Cells(i, 7).Value
            
            End If
            
       Next i
    
    'challenge portion was challenging
    'find last nonempty row to know how long for loops need to iterate to find'
    
    m = Cells(Rows.Count, 11).End(xlUp).Row
    
    'seed first percent change data, always positioned here
    
    highesthigh = Cells(3, 11).Value
    lowestlow = Cells(3, 11).Value
    biggestvolume = Cells(3, 12).Value
    
    For j = 3 To m - 1
        If Cells(j + 1, 11) > highesthigh Then
            'check if this is the new highest percent change
            highesthigh = Cells(j + 1, 11).Value
            highticker = Cells(j + 1, 1)
        End If
        
    Next j
            
    For j = 3 To m - 1
        If Cells(j + 1, 11) < lowestlow Then
            'check if this is the new lowest percent change'
            lowestlow = Cells(j + 1, 11).Value
            lowticker = Cells(j + 1, 1)
        End If
        
    Next j
        
    'And for total volume
    
    For j = 3 To m - 1
        
        If Cells(j + 1, 12) > biggestvolume Then
            'check if this is the new highest volume
            biggestvolume = Cells(j + 1, 12).Value
            bigticker = Cells(j + 1, 1)
                End If
    Next j
    
    'report winners
    
    Cells(2, 15) = highticker
    Cells(2, 16) = highesthigh
    Cells(3, 15) = lowticker
    Cells(3, 16) = lowestlow
    Cells(4, 15) = bigticker
    Cells(4, 16) = biggestvolume
            
    
    Next
  starting_ws.Activate
  
End Sub
