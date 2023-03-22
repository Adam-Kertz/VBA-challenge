Attribute VB_Name = "StockData"
Sub StockData()

    'Reset Output Data Cells
    Worksheets("2018").Columns("H:Z").ClearContents
    Worksheets("2018").Columns("H:Z").ClearFormats
    
    Worksheets("2019").Columns("H:Z").ClearContents
    Worksheets("2019").Columns("H:Z").ClearFormats
    
    Worksheets("2020").Columns("H:Z").ClearContents
    Worksheets("2020").Columns("H:Z").ClearFormats
    
    'Declare loop counters
    Dim Input_Data As Double
    Dim Output_Data As Integer
    Dim Worksheet_Tracker As Integer
    
    'Declare Array for Worksheet Names
    Dim Worksheet_Track(2) As String

    'Create an array of all worksheet names
    Worksheet_Track(0) = "2018"
    Worksheet_Track(1) = "2019"
    Worksheet_Track(2) = "2020"

    'Initialize variables for tracking stock information
    Dim Ticker As String
    Dim Ticker_Old As String
    Dim Open_Price As Double
    Dim Close_Price As Double
    Dim Volume As Double
    Dim Change As Double
    Dim Percent_Change As Double
    Dim Max_Increase As Double
    Dim Ticker_Max As String
    Dim Max_Decrease As Double
    Dim Ticker_Min As String
    Dim Max_Volume As Double
    Dim Ticker_Volume As String

    'Loop through each worksheet
    For Worksheet_Tracker = 0 To UBound(Worksheet_Track)

        'Start at 2 because we're not interested in data in title rows of tables
        Input_Data = 2
        Output_Data = 2

        'Initialize variables for first loop pass for this worksheet
        Volume = 0
        Ticker_Old = Worksheets(Worksheet_Track(Worksheet_Tracker)).Cells(Input_Data, 1).Value
        Open_Price = Worksheets(Worksheet_Track(Worksheet_Tracker)).Cells(Input_Data, 3).Value
        Max_Increase = 0
        Max_Decrease = 0
        Max_Volume = 0

        'Create header of output table
        Worksheets(Worksheet_Track(Worksheet_Tracker)).Cells(1, 9).Value = "Ticker"
        Worksheets(Worksheet_Track(Worksheet_Tracker)).Cells(1, 10).Value = "Yearly Change"
        Worksheets(Worksheet_Track(Worksheet_Tracker)).Cells(1, 11).Value = "Percent Change"
        Worksheets(Worksheet_Track(Worksheet_Tracker)).Cells(1, 12).Value = "Total Stock Volume"
        
        'Creae headers of "greatest" summary table
        Worksheets(Worksheet_Track(Worksheet_Tracker)).Cells(1, 16).Value = "Ticker"
        Worksheets(Worksheet_Track(Worksheet_Tracker)).Cells(1, 17).Value = "Value"
        Worksheets(Worksheet_Track(Worksheet_Tracker)).Cells(2, 15).Value = "Greatest % Increase"
        Worksheets(Worksheet_Track(Worksheet_Tracker)).Cells(3, 15).Value = "Greatest % Decrease"
        Worksheets(Worksheet_Track(Worksheet_Tracker)).Cells(4, 15).Value = "Greatest Total Volume"
        
        'Resize columns and format header text
        Worksheets(Worksheet_Track(Worksheet_Tracker)).Columns("I:L").AutoFit
        Worksheets(Worksheet_Track(Worksheet_Tracker)).Range("I1:L1").Font.Bold = True
        Worksheets(Worksheet_Track(Worksheet_Tracker)).Columns("O").AutoFit
        Worksheets(Worksheet_Track(Worksheet_Tracker)).Range("P1:Q1").Font.Bold = True
        Worksheets(Worksheet_Track(Worksheet_Tracker)).Range("O2:O4").Font.Bold = True

        Do
            'Find current ticker
            Ticker = Worksheets(Worksheet_Track(Worksheet_Tracker)).Cells(Input_Data, 1).Value

            If (Ticker = Ticker_Old) Then 'If ticker is same as previous iteration, update volume
                Volume = Volume + Worksheets(Worksheet_Track(Worksheet_Tracker)).Cells(Input_Data, 7).Value
            Else    'If ticker is different from previous iteration, output previous ticker data and reset varabiles
            
                'Assign closing price and calculate year end changes
                Close_Price = Worksheets(Worksheet_Track(Worksheet_Tracker)).Cells(Input_Data - 1, 6).Value
                Change = Close_Price - Open_Price
                Percent_Change = Change / Open_Price

                'Output Ticker, Yearly Change, Percent Change, and Total Volume to worksheet
                Worksheets(Worksheet_Track(Worksheet_Tracker)).Cells(Output_Data, 9).Value = Ticker_Old
                Worksheets(Worksheet_Track(Worksheet_Tracker)).Cells(Output_Data, 10).Value = FormatNumber(Change, 2)
                Worksheets(Worksheet_Track(Worksheet_Tracker)).Cells(Output_Data, 11).Value = FormatPercent(Percent_Change, 2)
                Worksheets(Worksheet_Track(Worksheet_Tracker)).Cells(Output_Data, 12).Value = FormatNumber(Volume, 0)
                
                'Highlight Yearly and Percent Change in Green if Positive, and Red if zero or negative
                If Change > 0 Then
                    Worksheets(Worksheet_Track(Worksheet_Tracker)).Cells(Output_Data, 10).Interior.ColorIndex = 4
                    Worksheets(Worksheet_Track(Worksheet_Tracker)).Cells(Output_Data, 11).Interior.ColorIndex = 4
                Else
                    Worksheets(Worksheet_Track(Worksheet_Tracker)).Cells(Output_Data, 10).Interior.ColorIndex = 3
                    Worksheets(Worksheet_Track(Worksheet_Tracker)).Cells(Output_Data, 11).Interior.ColorIndex = 3
                End If
                
                'Update Max Percent Increase Variable
                If Percent_Change > Max_Increase Then
                    Ticker_Max = Ticker_Old
                    Max_Increase = Percent_Change
                End If
                    
                'Update Max Percent Decrease Variable
                If Percent_Change < Max_Decrease Then
                    Ticker_Min = Ticker_Old
                    Max_Decrease = Percent_Change
                End If
                
                'Update Max Volume Variable
                If Volume > Max_Volume Then
                    Ticker_Volume = Ticker_Old
                    Max_Volume = Volume
                End If
                
                'Reset Variables for next Ticker
                Volume = 0
                Open_Price = Worksheets(Worksheet_Track(Worksheet_Tracker)).Cells(Input_Data, 3).Value
                Output_Data = Output_Data + 1
            End If

            'Update variables for next iteration
            Input_Data = Input_Data + 1
            Ticker_Old = Ticker
        
        'Set loop to iterate one more time than the number of populated columns (to print the final tracker information)
        Loop Until IsEmpty(Worksheets(Worksheet_Track(k)).Cells(Input_Data - 1, 1).Value)
        
        'Output Largest Increase and Decrease (in percent) and Volume
        Worksheets(Worksheet_Track(Worksheet_Tracker)).Cells(2, 16).Value = Ticker_Max
        Worksheets(Worksheet_Track(Worksheet_Tracker)).Cells(2, 17).Value = FormatPercent(Max_Increase, 2)
        
        Worksheets(Worksheet_Track(Worksheet_Tracker)).Cells(3, 16).Value = Ticker_Min
        Worksheets(Worksheet_Track(Worksheet_Tracker)).Cells(3, 17).Value = FormatPercent(Max_Decrease, 2)
        
        Worksheets(Worksheet_Track(Worksheet_Tracker)).Cells(4, 16).Value = Ticker_Volume
        Worksheets(Worksheet_Track(Worksheet_Tracker)).Cells(4, 17).Value = FormatNumber(Max_Volume, 0)
        
        'Resize summary table to fit numerical contents
        Worksheets(Worksheet_Track(Worksheet_Tracker)).Columns("P:Q").AutoFit
        
    Next Worksheet_Tracker
End Sub
