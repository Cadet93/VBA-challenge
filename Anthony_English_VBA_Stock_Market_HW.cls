VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub stock_market()

'Declare Variables
Dim Rowcounter As Double
Dim Lastrow As Double
Dim Ticker As String
Dim Firstopen As Double
Dim Yearlyclose As Double
Dim Yearlychange As Double
Dim Totalstock As Double
Dim Percent_Change As Double
Dim Greatest_Increases As Double
Dim Greatest_Decrease As Double
Dim Greatest_Volume As Double
Dim GIR As Double
Dim GDR As Double
Dim GVR As Double



'Loop through each worksheet
For Each ws In Worksheets

    Rowcounter = 2
    Lastrow = ws.Range("A" & Rows.Count).End(xlUp).Row
    'MsgBox Lastrow
    Firstopen = ws.Range("C2").Value

    'Add Column Headings
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"

    'set column width to fit for columns I-Q
    ws.Columns("I:Q").EntireColumn.AutoFit

    For i = 2 To Lastrow

        'Determine when the Ticker changes to a new stock
        If ws.Cells(i, 1) <> ws.Cells(i + 1, 1) Then
    
            'Set Ticker Symbol and Total Stock Value
            Ticker = ws.Cells(i, 1)
            ws.Range("I" & Rowcounter) = ws.Cells(i, 1)
            Totalstock = Totalstock + ws.Cells(i, 7)
            ws.Range("L" & Rowcounter) = Totalstock
        
            'Determine Yearly Change in the Stock Value
            Yearlyclose = ws.Cells(i, 6).Value
            Yearlychange = Yearlyclose - Firstopen
            ws.Cells(Rowcounter, 10) = Yearlychange
        
            If Yearlychange >= 0 Then
                ws.Range("J" & Rowcounter).Interior.ColorIndex = 4
            Else
                ws.Range("J" & Rowcounter).Interior.ColorIndex = 3
            End If
            
        
        
            'Determine Percent Change of the stock for the year
            If Firstopen = 0 Then
        
                Percent_Change = Yearlyclose - Firstopen
            
            
            Else
        
                Percent_Change = Yearlychange / Firstopen
            
            End If
        
        
            ws.Cells(Rowcounter, 11) = FormatPercent(Percent_Change)
            Firstopen = ws.Cells(i + 1, 3).Value
            Rowcounter = Rowcounter + 1
            Totalstock = 0
        
        
        Else
    
            Totalstock = Totalstock + ws.Cells(i, 7)
    
    
        End If
    
    

    Next i

    'Determine Greatest Increase, Decrease, and Volume
    Greatest_Increase = Application.WorksheetFunction.Max(ws.Range("K:K"))

    GIR = Application.WorksheetFunction.Match(Greatest_Increase, ws.Range("K:K"), 0)


    Greatest_Decrease = Application.WorksheetFunction.Min(ws.Range("K:K"))

    GDR = Application.WorksheetFunction.Match(Greatest_Decrease, ws.Range("K:K"), 0)


    Greatest_Volume = Application.WorksheetFunction.Max(ws.Range("L:L"))

    GVR = Application.WorksheetFunction.Match(Greatest_Volume, ws.Range("L:L"), 0)

    'Populate Greatest Values
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("P2").Value = ws.Cells(GIR, 9).Value
    ws.Range("Q2").Value = FormatPercent(Greatest_Increase)
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("P3").Value = ws.Cells(GDR, 9).Value
    ws.Range("Q3").Value = FormatPercent(Greatest_Decrease)
    ws.Range("O4").Value = "Greatest Total Volume"
    ws.Range("P4").Value = ws.Cells(GVR, 9).Value
    ws.Range("Q4").Value = Greatest_Volume

Next ws


End Sub
