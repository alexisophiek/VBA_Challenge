Attribute VB_Name = "VbaStocks"
Sub Stocks()
'loops for ws
Dim ws As Worksheet

For Each ws In ActiveWorkbook.Worksheets
ws.Activate

'label results table
Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Stock Volume"
Cells(1, 14).Value = "Results"
Cells(1, 15).Value = "Ticker"
Cells(1, 16).Value = "Values"
Columns("A:S").AutoFit

'define variables for storing data
lastusedrow = Sheet1.Cells(Rows.Count, 1).End(xlUp).Row


Dim ticker As String
Dim op As Double
Dim cl As Double


Dim difp As Double
difp = 0
Dim dif As Double
dif = 0

Dim sumv As Double
sumv = 0


'Set a rowcounter for placing data
Dim rc As Integer
rc = 2
     
'loop through tickers for each Data set
        op = Cells(2, 3).Value
        
        For x = 2 To lastusedrow
            If Cells(x + 1, 1).Value <> Cells(x, 1).Value Then
            cl = Cells(x, 6).Value
            ticker = Cells(x, 1).Value
            dif = cl - op
            sumv = sumv + Cells(x, 7).Value
            
            'for zero
            If op <> 0 Then
                difp = dif / op
                Else
                difp = 0
            End If
                                        
            Range("i" & rc).Value = ticker
            Range("j" & rc).Value = dif
            Range("l" & rc).Value = sumv
            Range("k" & rc).Value = difp
            rc = (rc + 1)
            
            op = Cells(x + 1, 3).Value
            
            End If
   
 Next x
 
'format data
lastresultsrow = Sheet1.Cells(Rows.Count, 10).End(xlUp).Row

    For i = 2 To lastresultsrow

    If Cells(i, 10).Value < 0 Then
        Cells(i, 10).Interior.ColorIndex = 3
        Cells(i, 11).NumberFormat = "0.00%"
        Else
    If Cells(i, 10).Value > 0 Then
        Cells(i, 10).Interior.ColorIndex = 4
        Cells(i, 11).NumberFormat = "0.00%"
        End If
        End If
    
Next i

'add the max value challenge
Dim maxinc As Double
Dim maxdec As Integer
Dim maxval As Long


maxinc = 0
maxdec = 0
maxval = 0

For v = 2 To lastresultsrow
 If maxinc < Cells(v, 11).Value Then
 maxinc = Cells(v, 11).Value
    If maxdec > Cells(v, 11).Value Then
    maxdec = Cells(v, 11).Value
        If maxval < Cells(v, 12).Value Then
        maxval = Cells(v, 12).Value
             
    End If
    End If
    End If
    Next v
  
Next ws

End Sub
