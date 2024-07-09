VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub MY_Stock_Data_Q1()

    Dim i As Long
   
    Dim Stock_Name As String
    
    Dim Total_Stock_Volume As Double
    Total_Stock_Volume = 0
    
    Dim Quarterly_Change As Double
    Quarterly_Change = 0
    
    Dim Percent_Change As Double
    Percent_Change = 0
    
    Dim Summary_Table_Row As Double
    Summary_Table_Row = 2
    
    Dim open_price As Double
    open_price = Cells(2, 3).Value
    
    For i = 2 To 93001
       
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            
                Stock_Name = Cells(i, 1).Value
            
                Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value
                Quarterly_Change = Cells(i, 6).Value - open_price
                Percent_Change = Quarterly_Change / open_price
    
                Range("I" & Summary_Table_Row).Value = Stock_Name
                Range("J" & Summary_Table_Row).Value = Quarterly_Change
                Range("K" & Summary_Table_Row).Value = Percent_Change
                Range("L" & Summary_Table_Row).Value = Total_Stock_Volume
                
                Summary_Table_Row = Summary_Table_Row + 1
                Total_Stock_Volume = 0
                open_price = Cells(i + 1, 3).Value
                
            Else
          
                Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value
                
            End If
       
    Next i
   
    Dim t1 As String
    Dim t2 As String
    Dim t3 As String
    
    Dim val1 As Double
    Dim val2 As Double
    Dim val3 As Double
    
    t1 = Cells(2, 9).Value
    t2 = Cells(2, 9).Value
    t3 = Cells(2, 9).Value
    
    val1 = Cells(2, 11).Value
    val2 = Cells(2, 11).Value
    val3 = Cells(2, 12).Value
    
   
    For j = 2 To Summary_Table_Row
   
        If Cells(j, 11) > val1 Then
            t1 = Cells(j, 9).Value
            val1 = Cells(j, 11).Value
        End If
        
        If Cells(j, 11) < val2 Then
            t2 = Cells(j, 9).Value
            val2 = Cells(j, 11).Value
        End If
        
        If Cells(j, 12) > val3 Then
            t3 = Cells(j, 9).Value
            val3 = Cells(j, 12).Value
        End If
        
        Cells(2, 15).Value = t1
        Cells(3, 15).Value = t2
        Cells(4, 15).Value = t3
    
        Cells(2, 16).Value = val1
        Cells(3, 16).Value = val2
        Cells(4, 16).Value = val3
    Next j
    

End Sub

