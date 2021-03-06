Attribute VB_Name = "Module1"
Sub StockSummary()
    
    ' Declare a worksheet variable to use to loop through each sheet in the workbook
    Dim WS As Worksheet
    ' Declare variables that will be used to store and output summary data
    Dim Ticker As String
    Dim Yearly_Change As Double
    Dim Yearly_Change_Percent As Double
    Dim Vol_Total As Double

    ' This will be used to indicate where the previous values will be written to
    Dim Summary_Row As Integer

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
    
    Dim LastRow As Long
    
    Dim i As Long
    
    For Each WS In Worksheets ' Open the Worksheet loop
    
        ' The following must be determined at the beginning of each new sheet.
        ' As such, it must take place inside the worksheet For Each loop.
        ' This must also take place outside the row For loop.
        
        ' Activates the current sheet held in WS to be worked through
        WS.Activate
        
        ' Determine what the last row in the data set is - bottom up
        LastRow = WS.Cells(WS.Rows.Count, 1).End(xlUp).Row
        
        ' Reset these variables for the beginning of each sheet
        Vol_Total = 0
        Summary_Row = 2
        
        ' Writing summary table headers to the specified location
        Range("I1").Value = "Ticker"
        Range("J1").Value = "Yearly Change"
        Range("K1").Value = "Percent Change"
        Range("L1").Value = "Total Stock Volume"
        
        ' Looping through all the rows in the sheet
        For i = 2 To LastRow
        
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
        
            ' Applies actions if the ticker changes - assumed to be alphabetical
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                      
                ' Using Min and Max rows, pull in opening min and closing max data
                ' Use placeholders Year_Start and Year_End to get yearly change and %
                Year_Start = Cells(MinDateRow, 3).Value
                Year_End = Cells(MaxDateRow, 6).Value
                Yearly_Change = Year_End - Year_Start
                ' Conditional to prevent DivBy0 Errory
        If Year_Start = 0 Then
                
                    Yearly_Change_Percent = 0 ' PLNT causing a DivBy0 Error
                    Range("K" & Summary_Row).Interior.Color = RGB(200, 200, 200) ' Draw attention to these cases
                Else
                    Yearly_Change_Percent = Yearly_Change / Year_Start
                
                End If
                
                Ticker = Cells(i, 1).Value ' Reads in ticker value before changing
                
                'Writes the stored summary statistics for the current ticket
                Range("I" & Summary_Row).Value = Ticker
                Range("J" & Summary_Row).Value = Yearly_Change
                Range("K" & Summary_Row).Value = Yearly_Change_Percent
                Range("L" & Summary_Row).Value = Vol_Total
                
                If Yearly_Change > 0 Then
                    Range("J" & Summary_Row).Interior.Color = RGB(0, 200, 0)
                ElseIf Yearly_Change < 0 Then
                    Range("J" & Summary_Row).Interior.Color = RGB(200, 0, 0)
                End If
        
                Summary_Row = Summary_Row + 1 ' Increments location for next ticker in summary
                
                ' Reset MinDate and MaxDate - They will need to be reset in order to calc
                ' indicate which values for the new ticket are min and max
                ' MinDateRow and MaxDateRow are overridden when a new MinDate and MaxDate
                ' are determined so no need to reset
                MinDate = 0
                MaxDate = 0
                
                Vol_Total = 0 ' Reset to 0 because new ticker to sum/increment on
            
            End If
        
        Next i ' Close for loop
            
    Next ' Close with loop

End Sub

Sub StockSummaryBonus()
    
    ' Declare a worksheet variable to use to loop through each sheet in the workbook
    Dim WS As Worksheet
    
    Dim LastRowSum As Long
    
    ' Declaring variables used to hold the largest/smallest values
    Dim MaxChange As Double
    Dim MaxChangeTicker As String
    Dim MinChange As Double
    Dim MinChangeTicker As String
    Dim MaxVol As Double
    Dim MaxVolTicker As String
    
    For Each WS In Worksheets ' Open the Worksheet loop
    
        ' Activates the current sheet held in WS to be worked through
        WS.Activate
        
        ' Reset these values at the beginning of each For Each loop
        ' This will allow each variable to be overridden by the first ticker
        ' This serves to prevent comparisons across sheets
        MaxChange = 0
        MinChange = 0
        MaxVol = 0
        
        Range("O2").Value = "Greatest % Increase"
        Range("O3").Value = "Greatest % Decrease"
        Range("O4").Value = "Greatest Total Volume"
        Range("P1").Value = "Ticker"
        Range("Q1").Value = "Value"
        
        
        ' Determine what the last row in the data set is - bottom up
        LastRowSum = WS.Cells(WS.Rows.Count, 9).End(xlUp).Row
    
        ' Looping through the summary table to identify greatest, least, and most vol
        For i = 2 To LastRowSum
        
            If Cells(i, 11).Value > MaxChange Or MaxChange = 0 Then
                MaxChange = Cells(i, 11)
                MaxChangeTicker = Cells(i, 9)
            End If
            If Cells(i, 11).Value < MinChange Or MinChange = 0 Then
                MinChange = Cells(i, 11)
                MinChangeTicker = Cells(i, 9)
            End If
            If Cells(i, 12).Value > MaxVol Or MaxVol = 0 Then
                MaxVol = Cells(i, 12)
                MaxVolTicker = Cells(i, 9)
            End If
        
        Next i
    
    ' Print titles for the bonus summary table
    Range("P2").Value = MaxChangeTicker
    Range("Q2").Value = MaxChange
    Range("P3").Value = MinChangeTicker
    Range("Q3").Value = MinChange
    Range("P4").Value = MaxVolTicker
    Range("Q4").Value = MaxVol
    
    ' As we have identified the last summary row in this subroutine, we do that here
    Range("K2:K" & LastRowSum).NumberFormat = "0.00%" ' Formats Percent Change column
    Range("Q2:Q3").NumberFormat = "0.00%" ' Format bonus percentage values
    
    Next

End Sub

'Sub MinMaxCheck()
'
'' Used to plan out tracking the min and max dates
'' Was tested on just the 'A' ticker and then subsequently for the 'AA' ticker
'
'    Dim MinDate As Long
'    Dim MinDateRow As Long
'
'    Dim MaxDate As Long
'    Dim MaxDateRow As Long
'
'    MinDate = 0
'    MinDateRow = 0
'    MaxDate = 0
'    MaxDateRow = 0
'
'    For i = 2 To 263
'
'        If Cells(i, 2).Value < MinDate Or MinDate = 0 Then
'            MinDate = Cells(i, 2)
'            MinDateRow = i
'
'        End If
'
'        If Cells(i, 2).Value > MaxDate Or MaxDate = 0 Then
'            MaxDate = Cells(i, 2)
'            MaxDateRow = i
'        End If
'
'    Next i
'
'    MsgBox (Str(MinDate))
'    MsgBox (Str(MaxDate))
'    MsgBox (Str(MinDateRow))
'    MsgBox (Str(MaxDateRow))
'
'End Sub
'
'Sub LoopSheets()
'
'    Dim WS As Worksheet
'
'    For Each WS In Worksheets
'
'        WS.Activate
'' ----------------------------------------
'' The goods go here... Avoid declaring a variable twice
'' Those goods will go outside of all loops
'' ----------------------------------------
'    Next
'
'End Sub
'
'Sub LastRow()
'
'    Dim Sht As Worksheet
'    Dim LastRow As Long
'
'    Set Sht = ActiveSheet
'
'    LastRow = Sht.Cells(Sht.Rows.Count, 1).End(xlUp).Row
'
'    MsgBox (Str(LastRow))
'
'End Sub
