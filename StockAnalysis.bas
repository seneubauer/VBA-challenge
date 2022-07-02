Attribute VB_Name = "StockAnalysis"
Option Explicit

Sub RunAnalysis()
    
    ' declare variables
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim summaryWS As Worksheet
    Dim summarySheetName As String
    Dim k As Integer
    Dim refRow As Integer
    Dim refColumn As Integer
    
    ' set the workbook
    Set wb = ActiveWorkbook
    
    ' pause workbook updating
    Application.ScreenUpdating = False
    
    ' define the summary worksheet's name
    summarySheetName = "Summary"
    
    ' delete the pre-existing summary worksheet
    If WorksheetExists(summarySheetName, wb) Then
        Application.DisplayAlerts = False
        wb.Worksheets(summarySheetName).Delete
        Application.DisplayAlerts = True
    End If
    
    ' create the new summary worksheet
    wb.Worksheets.Add(Before:=wb.Worksheets(1)).Name = summarySheetName
    Set summaryWS = wb.Worksheets(summarySheetName)
    
    ' set the initial reference row and column values
    ' CAUTION: refRow cannot be lower than 2 because I'm lazy
    refRow = 2
    refColumn = 2
    
    ' iterate through the worksheets
    For Each ws In wb.Worksheets
        If ws.Name <> summarySheetName Then
            Call AnalyzeStocks(ws, summaryWS, refRow, refColumn)
            refColumn = refColumn + 5
        End If
    Next ws
    
    ' resume workbook updating
    Application.ScreenUpdating = True
    
End Sub

Sub AnalyzeStocks(ws As Worksheet, summary As Worksheet, refRow As Integer, refColumn As Integer)
    
    ' raw data values
    Dim ticker0 As String
    Dim ticker1 As String
    Dim ticker2 As String
    Dim openValue As Double
    Dim closeValue As Double
    Dim stockVolume As LongLong
    Dim greatestPercentIncrease As Double
    Dim greatestPercentDecrease As Double
    Dim greatestTotalVolume As LongLong
    
    ' summary table ranges
    Dim tickerRNG As Range
    Dim yearlyRNG As Range
    Dim percentRNG As Range
    Dim stockRNG As Range
    
    ' indexers
    Dim i As Long
    Dim LR As Long
    Dim rIndex As Integer
    Dim r As Range
    
    ' colors
    Dim red As Integer
    Dim green As Integer
    Dim blue As Integer
    
    ' set the colors
    red = 3
    green = 4
    blue = 5
    
    ' get the worksheet's last row
    LR = ws.Range("A" & ws.Rows.Count).End(xlUp).Row
    
    ' set the initial results index
    rIndex = 1
    
    ' create the dataset label
    summary.Cells(refRow - 1, refColumn + 0).Value = "Dataset:"
    summary.Cells(refRow - 1, refColumn + 1).Value = ws.Name
    
    ' create & format the new headers
    With summary.Cells(refRow + 5, refColumn + 0)
        .Value = "Ticker"
        .Font.Bold = True
        .Font.Italic = True
    End With
    With summary.Cells(refRow + 5, refColumn + 1)
        .Value = "Yearly Change"
        .Font.Bold = True
        .Font.Italic = True
    End With
    With summary.Cells(refRow + 5, refColumn + 2)
        .Value = "Percent Change"
        .Font.Bold = True
        .Font.Italic = True
    End With
    With summary.Cells(refRow + 5, refColumn + 3)
        .Value = "Total Stock Volume"
        .Font.Bold = True
        .Font.Italic = True
    End With
    
    ' create & format the new bonus headers
    With summary.Cells(refRow + 0, refColumn + 2)
        .Value = "Ticker"
        .Font.Bold = True
        .Font.Italic = True
    End With
    With summary.Cells(refRow + 0, refColumn + 3)
        .Value = "Value"
        .Font.Bold = True
        .Font.Italic = True
    End With
    With summary.Cells(refRow + 1, refColumn + 1)
        .Value = "Greatest % Increase"
        .Font.Bold = True
        .Font.Italic = True
    End With
    With summary.Cells(refRow + 2, refColumn + 1)
        .Value = "Greatest % Decrease"
        .Font.Bold = True
        .Font.Italic = True
    End With
    With summary.Cells(refRow + 3, refColumn + 1)
        .Value = "Greatest Total Volumne"
        .Font.Bold = True
        .Font.Italic = True
    End With
    
    ' populate the new table
    For i = 2 To LR
        
        ' acquire the tickers
        ticker0 = CStr(ws.Range("A" & i - 1).Value)
        ticker1 = CStr(ws.Range("A" & i).Value)
        ticker2 = CStr(ws.Range("A" & i + 1).Value)
        
        If ticker0 <> ticker1 And ticker1 = ticker2 Then        ' at the start of a new section
            
            openValue = CDbl(ws.Range("C" & i).Value)
            stockVolume = CLngLng(ws.Range("G" & i).Value)
            
        ElseIf ticker0 = ticker1 And ticker1 = ticker2 Then     ' in the middle of a section
            
            stockVolume = stockVolume + CLngLng(ws.Range("G" & i).Value)
            
        ElseIf ticker0 = ticker1 And ticker1 <> ticker2 Then    ' at the end of a section
            
            closeValue = CDbl(ws.Range("F" & i).Value)
            stockVolume = stockVolume + CLngLng(ws.Range("G" & i).Value)
            
            ' write the results row
            summary.Cells(refRow + 5 + rIndex, refColumn + 0).Value = ticker1
            summary.Cells(refRow + 5 + rIndex, refColumn + 1).Value = closeValue - openValue
            summary.Cells(refRow + 5 + rIndex, refColumn + 2).Value = (closeValue - openValue) / openValue
            summary.Cells(refRow + 5 + rIndex, refColumn + 3).Value = stockVolume
            
            ' due to the wording of the assignment, I used the Conditional Formatting property
            ' instead of directly assigning interior colors to the relevant cells
            ' if I were to used direct color assignment, I'd have used this commented code
            'If closeValue > openValue Then      ' positive change
            '    summary.Cells(refRow + 5 + rIndex, refColumn + 1).Interior.ColorIndex = green
            'ElseIf closeValue < openValue Then  ' negative change
            '    summary.Cells(refRow + 5 + rIndex, refColumn + 1).Interior.ColorIndex = red
            'Else                                ' no change
            '    summary.Cells(refRow + 5 + rIndex, refColumn + 1).Interior.ColorIndex = blue
            'End If
            
            ' increment the row index
            rIndex = rIndex + 1
        End If
    Next i
    
    ' define the columns as ranges
    Set tickerRNG = summary.Range(summary.Cells(refRow + 5 + 1, refColumn + 0), summary.Cells(refRow + 5 + rIndex - 1, refColumn + 0))
    Set yearlyRNG = summary.Range(summary.Cells(refRow + 5 + 1, refColumn + 1), summary.Cells(refRow + 5 + rIndex - 1, refColumn + 1))
    Set percentRNG = summary.Range(summary.Cells(refRow + 5 + 1, refColumn + 2), summary.Cells(refRow + 5 + rIndex - 1, refColumn + 2))
    Set stockRNG = summary.Range(summary.Cells(refRow + 5 + 1, refColumn + 3), summary.Cells(refRow + 5 + rIndex - 1, refColumn + 3))
    
    ' apply the formatting
    With tickerRNG
        .NumberFormat = "General"
    End With
    With yearlyRNG
        .NumberFormat = "#,##0.00"
        .FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, Formula1:="=0"
        .FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:="=0"
        .FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, Formula1:="=0"
        .FormatConditions(1).Interior.Color = vbGreen
        .FormatConditions(2).Interior.Color = vbRed
        .FormatConditions(3).Interior.Color = vbBlue
    End With
    With percentRNG
        .NumberFormat = "0.00%"
    End With
    With stockRNG
        .NumberFormat = "#,##0"
    End With
    
    ' calculate the maximum Greatest % Increase
    greatestPercentIncrease = MaxDbl(percentRNG)
    greatestPercentDecrease = MinDbl(percentRNG)
    greatestTotalVolume = MaxLngLng(stockRNG)
    
    ' locate the corresponding tickers
    For Each r In tickerRNG
        
        ' find the greatest percent increase
        If r.Offset(0, 2).Value = greatestPercentIncrease Then
            summary.Cells(refRow + 1, refColumn + 2).Value = r.Value
            summary.Cells(refRow + 1, refColumn + 3).Value = r.Offset(0, 2).Value
            summary.Cells(refRow + 1, refColumn + 3).NumberFormat = "0.00%"
        End If
        
        ' find the greatest percent decrease
        If r.Offset(0, 2).Value = greatestPercentDecrease Then
            summary.Cells(refRow + 2, refColumn + 2).Value = r.Value
            summary.Cells(refRow + 2, refColumn + 3).Value = r.Offset(0, 2).Value
            summary.Cells(refRow + 2, refColumn + 3).NumberFormat = "0.00%"
        End If
        
        ' find the greatest total volume
        If r.Offset(0, 3).Value = greatestTotalVolume Then
            summary.Cells(refRow + 3, refColumn + 2).Value = r.Value
            summary.Cells(refRow + 3, refColumn + 3).Value = r.Offset(0, 3).Value
            summary.Cells(refRow + 3, refColumn + 3).NumberFormat = "#,##0"
        End If
    Next r
    
    ' autosize the columns
    summary.Columns(refColumn + 0).AutoFit
    summary.Columns(refColumn + 1).AutoFit
    summary.Columns(refColumn + 2).AutoFit
    summary.Columns(refColumn + 3).AutoFit
    
End Sub

Function MaxDbl(rng As Range) As Double
    
    Dim c As Range
    Dim result As Double
    
    result = -1.79769313486231E+308
    
    For Each c In rng
        If CDbl(c.Value) > result Then
            result = CDbl(c.Value)
        End If
    Next c
    
    MaxDbl = result
    
End Function

Function MinDbl(rng As Range) As Double
    
    Dim c As Range
    Dim result As Double
    
    result = 1.79769313486231E+308
    
    For Each c In rng
        If CDbl(c.Value) < result Then
            result = CDbl(c.Value)
        End If
    Next c
    
    MinDbl = result
    
End Function

Function MaxLngLng(rng As Range) As LongLong
    
    Dim c As Range
    Dim result As LongLong
    
    ' i don't know the actual minimum value but zero is a safe starting point in this application
    result = 0
    
    For Each c In rng
        If CLngLng(c.Value) > result Then
            result = CLngLng(c.Value)
        End If
    Next c
    
    MaxLngLng = result
    
End Function

Function WorksheetExists(wsName As String, wb As Workbook) As Boolean
    
    Dim ws As Worksheet
    Dim result As Boolean
    result = False
    
    For Each ws In wb.Worksheets
        If ws.Name = wsName Then
            result = True
            Exit For
        End If
    Next ws
    
    WorksheetExists = result
    
End Function