Sub Stock_code()

'makes the code run smoothly
Application.ScreenUpdating = False

'Set up for looping through worksheets
Dim ws As Worksheet
For Each ws In Worksheets
'Activate each sheet so the code can run on each
ws.Activate

'Enter in column headings
Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percentage Change"
Cells(1, 12).Value = "Total Stock Volume"

'Declaring variables
Dim Ticker_Name As String
Dim End_Value As Double
Dim Start_Value As Double
Dim Total_Stock As Double
Total_Stock = 0

'Find the last row in the first column so we include all data
Dim Last_Row As Long
Last_Row = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

'Keep track of the location the ticker value is in the summary
Dim Summary_Row As Integer
Summary_Row = 2

'Loop through all ticker options
For i = 2 To Last_Row

    'check to see where the change is. This is checking if the below line is the same as the current line.
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    
    'Setting the ticker name
    Ticker_Name = Cells(i, 1).Value
    
    'Setting the end value.
    End_Value = Cells(i, 6).Value
    
    'Add to the Total stock value
    Total_Stock = Total_Stock + Cells(i, 7).Value
    
    'Add the ticker name to the summary
    Range("I" & Summary_Row).Value = Ticker_Name
    
    'Add the total stock value to the summary
    Range("L" & Summary_Row).Value = Total_Stock
    
    'Add the yearly change to the summary
    Range("J" & Summary_Row).Value = End_Value - Start_Value
    
    'Add the % change, noting that 0 results in overflow
    If Not Start_Value = 0 Then
    Range("K" & Summary_Row).Value = (End_Value - Start_Value) / Start_Value
    Range("K" & Summary_Row).NumberFormat = "0.00%"
    
    Else
    Range("K" & Summary_Row).Value = 0
    
    End If
    
    'Add one to the ticker summary row so the next entry goes on the next line
    Summary_Row = Summary_Row + 1
    
    'Reset the total stock value because we want the whole value and not only 1 line
    Total_Stock = 0
    
    'check to see whether the row above is the same. To find the first line of each set.
    ElseIf Cells(i - 1, 1).Value <> Cells(i, 1).Value Then
    
    'Setting start value
    Start_Value = Cells(i, 3).Value
                
    Else
    
    'if the rows are the same ticker
    Total_Stock = Total_Stock + Cells(i, 7).Value
        
    End If
    
   Next i
   
'start of conditional formatting
'goes through each line and checks if it is a positive or negative number

Dim Last_Row2 As Long
Last_Row2 = ws.Cells(Rows.Count, 10).End(xlUp).Row

    For x = 2 To Last_Row2

    If Cells(x, 10).Value < 0 Then

    Cells(x, 10).Interior.ColorIndex = 3

    ElseIf Cells(x, 10).Value > 0 Then

    Cells(x, 10).Interior.ColorIndex = 4

    End If

Next x

'start of bonus question

'setting headings
Cells(2, 15).Value = "Greatest % increase"
Cells(3, 15).Value = "Greatest % decrease"
Cells(4, 15).Value = "Greatest Total Volume"
Cells(1, 16).Value = "Ticker"
Cells(1, 17).Value = "Value"

'goes through each row and checks whether the number is greater or less than the previous line and stores accordingly.
Dim Last_Row3 As Long
Last_Row3 = Cells(Rows.Count, 11).End(xlUp).Row

Dim Max As Double
Max = 0

Dim Min As Double
Min = 0

Dim Total As Variant
Total = 0

Dim Ticker_Name_Max As String
Ticker_Name_Max = ""

Dim Ticker_Name_Min As String
Ticker_Name_Min = ""

Dim Ticker_Name_Total As String
Ticker_Name_Total = ""

For k = 2 To Last_Row3

    If Cells(k, 11).Value > Max Then
    Max = Cells(k, 11).Value
    Ticker_Name_Max = Cells(k, 9).Value
    
    End If
    
    If Cells(k, 11).Value < Min Then
    Min = Cells(k, 11).Value
    Ticker_Name_Min = Cells(k, 9).Value

    End If
    
    If Cells(k, 12).Value > Total Then
    Total = Cells(k, 12).Value
    Ticker_Name_Total = Cells(k, 9).Value

    End If
    
Next k

Range("P2").Value = Ticker_Name_Max
Range("Q2").Value = Max
Range("Q2").NumberFormat = "0.00%"

Range("P3").Value = Ticker_Name_Min
Range("Q3").Value = Min
Range("Q3").NumberFormat = "0.00%"

Range("P4").Value = Ticker_Name_Total
Range("Q4").Value = Total

'changing the column size so all text fits nicer on the screen.
Columns("I:Q").Select
Columns("I:Q").EntireColumn.AutoFit
Range("I1").Select
    
   Next ws

End Sub




