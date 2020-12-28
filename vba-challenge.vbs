Attribute VB_Name = "Module1"
Sub StockSummary()
    
    ' Declare variables that will be used to store and output summary data
    Dim Ticker As String
    Dim Yearly_Change As Double
    Dim Yearly_Change_Percent As Double
    Dim Vol_Total As Double
        Vol_Total = 0
    ' This will be used to indicate where the previous values will be written to
    Dim Summary_Row As Integer
        Summary_Row = 2
        
    ' These variables are used to calculate the Yearly_Change variable
    Dim Year_Start As Double
    Dim Year_End As Double
    
    ' These variables are used to determine what entry and it's corresponding row
    ' for the earliest and latest date for a ticker
    ' This accounts for the case that the ticker are not already in chronological
    ' order while also identifying the opening and closing date
    Dim MinDate As Long
    Dim MinDateRow As Long
    Dim MaxDate As Long
    Dim MaxDateRow As Long
    
    Dim i As Long
    
    ' Writing summary table headers to the specified location
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
    
    ' Looping through all the rows in the sheet
    ' Currently set to manually work for the 'A' sheet
    For i = 2 To 70926
    
        ' Applies actions if the ticker changes - assumed to be alphabetical
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            
            Vol_Total = Vol_Total + Cells(i, 7) ' Update Vol_Total for the last time
            
            ' Checks last row for ticket if new min or new max
            ' Test if this can be run outside of If statement - exists in Else
            If Cells(i, 2).Value < MinDate Or MinDate = 0 Then
                MinDate = Cells(i, 2)
                MinDateRow = i
            End If
            If Cells(i, 2).Value > MaxDate Or MaxDate = 0 Then
                MaxDate = Cells(i, 2)
                MaxDateRow = i
            End If
            
            ' Using Min and Max rows, pull in opening min and closing max data
            ' Use placeholders Year_Start and Year_End to get yearly change and %
            Year_Start = Cells(MinDateRow, 3).Value
            Year_End = Cells(MaxDateRow, 6).Value
            Yearly_Change = Year_End - Year_Start
            Yearly_Change_Percent = Yearly_Change / Year_Start
            
            Ticker = Cells(i, 1).Value ' Reads in ticker value before changing
            
            'Writes the stored summary statistics for the current ticket
            Range("I" & Summary_Row).Value = Ticker
            Range("J" & Summary_Row).Value = Yearly_Change
            Range("K" & Summary_Row).Value = Yearly_Change_Percent
            Range("L" & Summary_Row).Value = Vol_Total
    
            Summary_Row = Summary_Row + 1 ' Increments location for next ticker
            
            ' Reset MinDate and MaxDate - They will need to be reset in order to calc
            ' indicate which values for the new ticket are min and max
            ' MinDateRow and MaxDateRow are overridden when a new MinDate and MaxDate
            ' are determined so no need to reset
            MinDate = 0
            MaxDate = 0
            
            Vol_Total = 0 ' Reset to 0 because new ticker to sum/increment on
            
        Else

            ' Adds the volume for the current ticker onto a running total
            Vol_Total = Vol_Total + Cells(i, 7)
            
            ' Checks if row for current ticker if new min or new max
            ' Test if this can be run outside of If statement - exists in If
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

' Used to plan out tracking the min and max dates
' Was tested on just the 'A' ticker and then subsequently for the 'AA' ticker

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

' Used to plan out looping and applying the StockSummary() subroutine to all sheets
    
    Dim WS As Worksheet
    
    For Each WS In Worksheets
' ----------------------------------------
'   Put the goods in here? It's not working!!!
' ----------------------------------------
    Next

End Sub
