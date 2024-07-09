Attribute VB_Name = "Module1"
Option Explicit



Sub Stock_data_01()

'Define all of needed dimensions

    Dim ws As Worksheet
    Dim total As Double
    Dim i, find_value As Long
    Dim QuarterlyChange As Double
    Dim j As Integer
    Dim start As Long
    Dim LastRow As Long
    Dim PerChange As Double
    Dim days As Integer
    
   

'=======================================================
'Starting with loop through each worksheet (Next ws noted already)

    For Each ws In Worksheets
   
        ws.Activate
        
'Add New columns to to Worksheet with proper names and Titles
        
        Range("I1").Value = "Ticker"
        Range("J1").Value = "Quarterly Change"
        Range("K1").Value = "Percent Change"
        Range("L1").Value = "Total Stock Volume"
        Columns("J:L").ColumnWidth = 20
       
        
        
' Setting the Initial Values and finding the las Row
        
        
        j = 0
        total = 0
        QuarterlyChange = 0
        start = 2
        
'        LastRow = Cells(Rows.Count, "A").End(xlUp).Row

        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        
        For i = 2 To LastRow
        
' Check the ticker sign and if it changes to others

        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        
           total = total + Cells(i, 7).Value  'Stores results
          
 
           If total = 0 Then
           
               Range("I" & 2 + j).Value = Cells(i, 1).Value
               Range("J" & 2 + j).Value = 0
               Range("K" & 2 + j).Value = "%" & 0
               Range("L" & 2 + j).Value = 0
            
            Else       'Find non zero starting value
              
              If Cells(start, 3) = 0 Then
              
                
              
                For find_value = start To i
               
                       If Cells(find_value, 3).Value <> 0 Then
                       
                           start = find_value
                           
                           Exit For
                       End If
                Next find_value
                
              End If
               
'Calculation of changes

            QuarterlyChange = (Cells(i, 6) - Cells(start, 3))
            
            PerChange = QuarterlyChange / Cells(start, 3)
            
'Start of the next stock ticker and record results

            start = i + 1
            Range("I" & 2 + j).Value = Cells(i, 1).Value
            Range("J" & 2 + j).Value = QuarterlyChange
            Range("J" & 2 + j).NumberFormat = "0.00"
            Range("K" & 2 + j).Value = PerChange
            Range("K" & 2 + j).NumberFormat = "0.00%"
            Range("L" & 2 + j).Value = total
           
            
'Add color coding to QuarterlyChange

                If (QuarterlyChange > 0) Then
                
                    Range("J" & 2 + j).Interior.ColorIndex = 4
                
                ElseIf (QuarterlyChange <= 0) Then
                
                    Range("J" & 2 + j).Interior.ColorIndex = 3
                
                End If
                

               
               
            End If
                         
'Move on to next ticker symbol

           total = 0
           
           QuarterlyChange = 0
           
           j = j + 1
           
           days = 0
           
'Add results for each ticker symbol

       Else
       
           total = total + Cells(i, 7).Value
           
    End If
    
    Next i
   
   
    Next ws

End Sub




