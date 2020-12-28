Attribute VB_Name = "Module1"
Sub StockSummary()
    
    Dim Ticker As String
    
    Dim Yearly_Change As Double
    
    Dim Year_Start As Double
    
    Dim Year_End As Double
    
    Dim Yearly_Change_Percent As Double
    
    Dim Vol_Total As Double
        Vol_Total = 0
        
    Dim MinDate As Long
        
    Dim MinDateRow As Long
        
    Dim MaxDate As Long
        
    Dim MaxDateRow As Long
        
    Dim Summary_Row As Integer
        Summary_Row = 2
    
    Dim i As Long
    
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
    
    For i = 2 To 70926
    
        ' Pull in the opening value for just the first ticker
    
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            
            Vol_Total = Vol_Total + Cells(i, 7)
            
            If Cells(i, 2).Value < MinDate Or MinDate = 0 Then
                MinDate = Cells(i, 2)
                MinDateRow = i
            End If
            If Cells(i, 2).Value > MaxDate Or MaxDate = 0 Then
                MaxDate = Cells(i, 2)
                MaxDateRow = i
            End If
            
            Year_Start = Cells(MinDateRow, 3).Value
            Year_End = Cells(MaxDateRow, 6).Value
            Yearly_Change = Year_End - Year_Start
            Yearly_Change_Percent = Yearly_Change / Year_Start
            
            Ticker = Cells(i, 1).Value
             
            Range("I" & Summary_Row).Value = Ticker

            Range("J" & Summary_Row).Value = Yearly_Change
            
            Range("K" & Summary_Row).Value = Yearly_Change_Percent

            Range("L" & Summary_Row).Value = Vol_Total
    
            Summary_Row = Summary_Row + 1
            
            ' Get the close value from the new ticker
            ' Run your calculations
            ' Insert that into the summary table
            
            ' Get the new open value, for the new ticker
            
            
            MinDate = 0
            MaxDate = 0
            Vol_Total = 0
            
        Else

            Vol_Total = Vol_Total + Cells(i, 7)
            
            If Cells(i, 2).Value < MinDate Or MinDate = 0 Then
                MinDate = Cells(i, 2)
                MinDateRow = i
            End If
            If Cells(i, 2).Value > MaxDate Or MaxDate = 0 Then
                MaxDate = Cells(i, 2)
                MaxDateRow = i
            End If
        
        End If
    
    Next i

End Sub

Sub MinMaxCheck()

    Dim MinDate As Long
    Dim MinDateRow As Long
    
    Dim MaxDate As Long
    Dim MaxDateRow As Long
    
    MinDate = 0
    MinDateRow = 0
    MaxDate = 0
    MaxDateRow = 0
    
    For i = 2 To 263
    
        If Cells(i, 2).Value < MinDate Or MinDate = 0 Then
            MinDate = Cells(i, 2)
            MinDateRow = i
            
        End If
        
        If Cells(i, 2).Value > MaxDate Or MaxDate = 0 Then
            MaxDate = Cells(i, 2)
            MaxDateRow = i
        End If
    
    Next i

    MsgBox (Str(MinDate))
    MsgBox (Str(MaxDate))
    MsgBox (Str(MinDateRow))
    MsgBox (Str(MaxDateRow))
    
End Sub

Sub LoopSheets()

    Dim WS As Worksheet
    
    For Each WS In Worksheets
' ----------------------------------------
'   Put the goods in here? It's not working!!!
' ----------------------------------------
    Next

End Sub