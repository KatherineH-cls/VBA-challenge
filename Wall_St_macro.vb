Option Explicit

Sub WorksheetLoop()

    ' I copied this worksheet loop code from the microsoft website
    ' https://support.microsoft.com/en-us/topic/macro-to-loop-through-all-worksheets-in-a-workbook-feef14e3-97cf-00e2-538b-5da40186e2b0

         Dim WS_Count As Integer
         Dim WS As Integer

         ' Set WS_Count equal to the number of worksheets in the active
         ' workbook.
         WS_Count = ActiveWorkbook.Worksheets.Count

         ' Begin the loop.
         For WS = 1 To WS_Count

            ' Insert your code here.
            Worksheets(WS).Select

            Call Ticker_homework
            ' The following line shows how to reference a sheet within
            ' the loop by displaying the worksheet name in a dialog box.
            ' MsgBox ActiveWorkbook.Worksheets(WS).Name
            
         Next WS

      End Sub



Sub Ticker_homework()


'Code to execute homework
'Task to find opening and closing value for each stock over period

Dim Ticker As String
Dim row_in As Long
Dim row_out As Integer
Dim vol As LongLong
Dim openval As Double
Dim closeval As Double
Dim year_change As Double
Dim percent_change As Double

'Put output headings in excel sheet
Cells(1, "I").Value = "Ticker"
Cells(1, "J").Value = "Yearly Change"
Cells(1, "K").Value = "Percent Change"
Cells(1, "L").Value = "Total Stock Volume"


'start at first ticker
'initialise values
    row_in = 2
    row_out = 2
    
    Do

        'identify and store ticker
        'go to start value, collect and store opening
        'initialise vol
        Ticker = Cells(row_in, "A").Value
        openval = Cells(row_in, "C").Value
        vol = 0
    
        'start loop, ending when ticker code changes
        'add to stock volume
        
        Do
    
            vol = vol + Cells(row_in, "G").Value
            row_in = row_in + 1
            
        Loop Until Cells(row_in, "A").Value <> Ticker
     
        'collect and store closing
            closeval = Cells(row_in - 1, "F").Value
        'do calcs
            year_change = closeval - openval
            If openval <> 0 Then
                percent_change = year_change / openval
            Else
                ' to avoid division by zero error
                percent_change = 0
            End If
        'store output
            Cells(row_out, "I").Value = Ticker
            Cells(row_out, "J").Value = year_change
            Cells(row_out, "K").Value = percent_change
            Cells(row_out, "L").Value = vol
        'end loop for this ticker and start next one
        row_out = row_out + 1

    'if ticker cell is empty, end procedure
    ' (note would need a different test if there were any empty rows
    ' in the middle of the data)
    Loop Until Cells(row_in, "A").Value = ""
    
    'now format cells
    'format percent change as a percent
    Range(Cells(2, "K"), Cells(2, "K").End(xlDown)).NumberFormat = "0.00%"
    'increase column width of Total Stock Volume
    Range("L1").ColumnWidth = 15
    
    'format background colour of yearly change cells
    Dim rng As Range
        
    Set rng = Range(Cells(2, "J"), Cells(2, "J").End(xlDown))
    
    rng.FormatConditions.Delete
    
    rng.FormatConditions.Add xlCellValue, xlGreater, "=0"
    rng.FormatConditions(1).Interior.Color = RGB(0, 255, 0)
    
    rng.FormatConditions.Add xlCellValue, xlLess, "=0"
    rng.FormatConditions(2).Interior.Color = RGB(255, 0, 0)
    
    'Challenge exercise to create table of biggest changes
    
    'Put in table headings
    
    Cells(1, "P").Value = "Ticker"
    Cells(1, "Q").Value = "Value"
    Cells(2, "O").Value = "Greatest % Increase"
    Cells(3, "O").Value = "Greatest % Decrease"
    Cells(4, "O").Value = "Greatest Total Volume"
    Columns("O").AutoFit
    Columns("Q").ColumnWidth = 15
    Range("Q2:Q3").NumberFormat = "0.00 %"
    
    Dim maxincrease As Double
    Dim maxdecrease As Double
    Dim maxvolume As LongLong
    Dim Loc_maxincrease As Integer
    Dim Loc_maxdecrease As Integer
    Dim Loc_maxvol As Integer
    
    'find maximums
    maxincrease = WorksheetFunction.Max(Range("K:K"))
    maxdecrease = WorksheetFunction.Min(Range("K:K"))
    maxvolume = WorksheetFunction.Max(Range("L:L"))
    
    'identify ticker associated with maximums
    Loc_maxincrease = WorksheetFunction.Match(maxincrease, Range("K:K"), 0)
    Loc_maxdecrease = WorksheetFunction.Match(maxdecrease, Range("K:K"), 0)
    Loc_maxvol = WorksheetFunction.Match(maxvolume, Range("L:L"), 0)
    
    'Output biggest changes results to worksheet
    
    Cells(2, "Q").Value = maxincrease
    Cells(3, "Q").Value = maxdecrease
    Cells(4, "Q").Value = maxvolume
    
    Cells(2, "P") = Cells(Loc_maxincrease, "I").Value
    Cells(3, "P") = Cells(Loc_maxdecrease, "I").Value
    Cells(4, "P") = Cells(Loc_maxvol, "I").Value
    
 
End Sub
