Sub Insert_Columns()
	Dim ws as Worksheet
	For Each ws in Worksheets
		ws.Cells(1, 9).Value = "Ticker"
		ws.Cells(1,10).Value = "Yearly Change"
		ws.Cells(1,11).Value = "Percent Change"
		ws.Cells(1,12).Value = "Total Stock Volume"
		ws.Cells(2,15).Value = "Greatest % Increase"
		ws.Cells(3,15).Value = "Greatest % Decrease"
		ws.Cells(4,15).Value = "Greatest Total Volume"
		ws.Cells(1,16).Value = "Ticker"
		ws.Cells(1,17).Value = "Value"
	Next
End Sub

'Insert Ticker
Sub Insert_Ticker()
	Dim ws as Worksheet 
	Dim LastRow As Long
	For Each ws in Worksheets
		LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
		For Ticks = 2 To LastRow
			ws.Cells(Ticks, 9).Value = Cells(Ticks, 1).Value
		Next Ticks
	Next
End Sub

'Insert Percentage changes
Sub Insert_Percentages()
    Dim ws As Worksheet
    Dim LastRow As Long
    Dim Value_Subtraction As Double
    Dim Value_Percentage As Double
    Dim Value_Starting As Double
    For Each ws In Worksheets
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        For Ticks = 2 To LastRow
            Value_Starting = Round(ws.Cells(Ticks, 3).Value, 0)
            Value_Subtraction = ws.Cells(Ticks, 3).Value - ws.Cells(Ticks, 6).Value
            Value_Percentage = Value_Subtraction * 100 / 50
            ws.Cells(Ticks, 10).Value = Value_Subtraction
            ws.Cells(Ticks, 11).Value = Value_Percentage
        Next Ticks
    Next
End Sub

'Insert Percentages Color Code
Sub Insert_Percentage_Colors()
    Dim ws As Worksheet
    Dim LastRow As Long
    Dim Value_Subtraction As Double
    Dim Value_Percentage As Double
    For Each ws In Worksheets
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        For Ticks = 2 To LastRow
            'if percentage is negative color red
            'if percentage is positive color green
            'if percentage is neutral color yellow
            If ws.Cells(Ticks, 10).Value > 0 Then
                ws.Cells(Ticks, 10).Interior.ColorIndex = 4
            ElseIf ws.Cells(Ticks, 10).Value < 0 Then
                ws.Cells(Ticks, 10).Interior.ColorIndex = 3
            Else
                ws.Cells(Ticks, 10).Interior.ColorIndex = 6
            End If
        Next Ticks
    Next
End Sub

'Total stocks function
Sub Total_Stocks()
    Dim ws As Worksheet
    Dim LastRow As Long
    For Each ws In Worksheets
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        For Ticks = 2 To LastRow
            ws.Cells(Ticks, 12).Value = Cells(Ticks, 7).Value
        Next Ticks
    Next
End Sub

'Function to get highest percentage growth
Sub Get_Highest()
    Dim ws As Worksheet
    Dim LastRow As Long
    For Each ws In Worksheets
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row - 1
        For Ticks = 2 To LastRow
            If (ws.Cells(Ticks + 1, 10) > ws.Cells(Ticks, 10)) Then
                Cells(2, 16).Value = Cells(Ticks + 1, 9)
                Cells(2, 17).Value = Cells(Ticks + 1, 10)
                
            End If
        Next Ticks
    Next
End Sub


'Function to get lowest percentage growth
Sub Get_Lowest()
    Dim ws As Worksheet
    Dim LastRow As Long
    For Each ws In Worksheets
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row - 1
        For Ticks = 2 To LastRow
            If (ws.Cells(Ticks + 1, 10) < ws.Cells(Ticks, 10)) Then
                Cells(3, 16).Value = Cells(Ticks + 1, 9)
                Cells(3, 17).Value = Cells(Ticks + 1, 10)
                
            End If
        Next Ticks
    Next
End Sub


'Function to get the highest volume
Sub Get_Highest_Volume()
    Dim ws As Worksheet
    Dim LastRow As Long
    For Each ws In Worksheets
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row - 1
        For Ticks = 2 To LastRow
            If (ws.Cells(Ticks + 1, 12) > ws.Cells(Ticks, 12)) Then
                Cells(4, 16).Value = Cells(Ticks + 1, 9)
                Cells(4, 17).Value = Cells(Ticks + 1, 10)
                
            End If
        Next Ticks
    Next
End Sub