Attribute VB_Name = "Module1_Liska"
Sub MedWallStreet()
'Assignment 2 for Data Science Bootcamp, by David Liska 6/9/2019.

Dim intSecondaryRow As Integer
Dim lngCurrentRow, lngPasteRow As Long
Dim dblInitialOpPrice, dblCollectedVolume, dblGreatPercInc, dblGreatPercDec, dblGreatTotVol As Double
Dim strTicker, strGreatPercIncTicker, strGreatPercDecTicker, strGreatTotVolTicker As String
Dim strSheetName As String
Dim ws As Worksheet

For Each ws In Worksheets
    

    'Initialize variables
    lngCurrentRow = 1
    lngPasteRow = 2
    strTicker = ""
    
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
    
    Do
        Do
            lngCurrentRow = lngCurrentRow + 1
        
            strTicker = ws.Cells(lngCurrentRow, 1).Value
            If ws.Cells(lngCurrentRow - 1, 1).Value <> strTicker Then dblInitialOpPrice = ws.Cells(lngCurrentRow, 3).Value
    
            dblCollectedVolume = ws.Cells(lngCurrentRow, 7).Value + dblCollectedVolume
    
        Loop While ws.Cells(lngCurrentRow + 1, 1).Value = strTicker
        
        ws.Cells(lngPasteRow, 9).Value = strTicker
        ws.Cells(lngPasteRow, 10).Value = ws.Cells(lngCurrentRow, 6).Value - dblInitialOpPrice
        If dblInitialOpPrice > 0 Then
            ws.Cells(lngPasteRow, 11).Value = Round((ws.Cells(lngCurrentRow, 6).Value - dblInitialOpPrice) / dblInitialOpPrice, 4)
        Else
            ws.Cells(lngPasteRow, 11).Value = 0
        End If
        ws.Cells(lngPasteRow, 12).Value = dblCollectedVolume
        
        dblCollectedVolume = 0
        
        lngPasteRow = lngPasteRow + 1
    
    Loop While ws.Cells(lngCurrentRow + 1, 1).Value <> ""
    
    'Generate greatest percents and volume table
    LastRow = ws.Cells(Rows.Count, 9).End(xlUp).Row
    dblGreatPercInc = 0
    dblGreatPercDec = 0
    dblGreatTotVol = 0
    
    For intSecondaryRow = 2 To LastRow
    
    
        If ws.Cells(intSecondaryRow, 11).Value > dblGreatPercInc Then
            dblGreatPercInc = ws.Cells(intSecondaryRow, 11).Value
            strGreatPercIncTicker = ws.Cells(intSecondaryRow, 9).Value
        End If
        
        If ws.Cells(intSecondaryRow, 11).Value < dblGreatPercDec Then
            dblGreatPercDec = ws.Cells(intSecondaryRow, 11).Value
            strGreatPercDecTicker = ws.Cells(intSecondaryRow, 9).Value
        End If
        
        If ws.Cells(intSecondaryRow, 12).Value > dblGreatTotVol Then
            dblGreatTotVol = ws.Cells(intSecondaryRow, 12).Value
            strGreatTotVolTicker = ws.Cells(intSecondaryRow, 9).Value
        End If
    
    
    Next intSecondaryRow
    
    strSheetName = ws.Name
    Sheets(strSheetName).Select
    
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
    
    
    ws.Range("P2").Value = strGreatPercIncTicker
    ws.Range("Q2").Value = dblGreatPercInc
           
    ws.Range("P3").Value = strGreatPercDecTicker
    ws.Range("Q3").Value = dblGreatPercDec
    
    ws.Range("P4").Value = strGreatTotVolTicker
    ws.Range("Q4").Value = dblGreatTotVol
    
    Columns("O:O").EntireColumn.AutoFit
    Columns("Q:Q").EntireColumn.AutoFit
    
    Sheets(strSheetName).Select
    ws.Range("K2").Select
    ws.Range(Selection, Selection.End(xlDown)).Select
    Selection.Style = "Percent"
    Selection.NumberFormat = "0.00%"
    
    ws.Range("Q2:Q3").Select
    Selection.Style = "Percent"
    Selection.NumberFormat = "0.00%"
    
    ws.Range("J2").Select
    ws.Range(Selection, Selection.End(xlDown)).Select
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, _
        Formula1:="=0"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 5287936
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, _
        Formula1:="=0"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 255
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False


Next ws
MsgBox "Macro has completed running.", vbOKOnly, "Macro Complete"
End Sub
