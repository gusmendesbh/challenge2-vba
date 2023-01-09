Attribute VB_Name = "final"
Sub woorksheetLoop()
    Dim xSh As Worksheet
    Application.ScreenUpdating = False
    For Each xSh In Worksheets
        xSh.Select
        Call stockAnalysis
    Next
    Application.ScreenUpdating = True
End Sub

Sub stockAnalysis()

Dim tick As String, newTick As String, highTick As String, lowTick As String, volumeTick As String
Dim openValue As Double, closeValue As Double, percent As Double, high As Double, low As Double
Dim total As LongLong, tempTotal As LongLong, volume As LongLong
Dim i As Long
Dim j As Integer

Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly_Change"
Cells(1, 11).Value = "Percent_Change"
Cells(1, 12).Value = "Total_Stock_Volume"
Cells(1, 16).Value = "Ticker"
Cells(1, 17).Value = "Value"
Cells(2, 15).Value = "Greatest_%_Increase"
Cells(3, 15).Value = "Greatest_%_Decrease"
Cells(4, 15).Value = "Greatest_Total_Volume"

tick = Cells(2, 1).Value
openValue = Cells(2, 3).Value

tempTotal = 0

high = 0
low = 0
volume = Cells(2, 12).Value

j = 2
For i = 2 To 753001
    

newTick = Cells(i, 1).Value
    If tick <> newTick Then
        closeValue = Cells(i - 1, 6).Value

        Cells(j, 10).Value = closeValue - openValue
            If Cells(j, 10).Value > 0 Then
                Cells(j, 10).Interior.ColorIndex = 4
            ElseIf Cells(j, 10).Value < 0 Then
                Cells(j, 10).Interior.ColorIndex = 3
            Else
                Cells(j, 10).Interior.ColorIndex = 6
            End If
            
            If Cells(j, 11).Value > 0 Then
                Cells(j, 11).Interior.ColorIndex = 4
            ElseIf Cells(j, 10).Value < 0 Then
                Cells(j, 11).Interior.ColorIndex = 3
            Else
                Cells(j, 11).Interior.ColorIndex = 6
            End If
            
            
        Cells(j, 11).Value = Format(((openValue - closeValue) / openValue) * -1, "0.00%")
       
        
        If high > Cells(j, 11).Value Then
            high = Cells(j, 11).Value
            highTick = Cells(j, 9).Value
        End If
        
        If low < Cells(j, 11).Value Then
            low = Cells(j, 11).Value
            lowTick = Cells(j, 9).Value
        End If
        
        If volume < Cells(j, 12).Value Then
            volume = Cells(j, 12).Value
            volumeTick = Cells(j, 9).Value
        End If
        
        Cells(2, 16).Value = highTick
        Cells(2, 17).Value = Format(high * -1, "0.00%")
        Cells(3, 16).Value = lowTick
        Cells(3, 17).Value = Format(low * -1, "0.00%")
        Cells(4, 16).Value = volumeTick
        Cells(4, 17).Value = volume
                      
                      
        Cells(j, 9).Value = tick
            
        tempTotal = tempTotal + Cells(i, 7).Value

        Cells(j, 12).Value = tempTotal

        j = j + 1
            
        tempTotal = 0
            
        openValue = Cells(i, 3).Value
        
        tick = newTick
               
            
    Else
        
        tempTotal = tempTotal + Cells(i, 7).Value
           
    End If
        
Next i

End Sub

