Attribute VB_Name = "Module1"
Sub AlphaTest()
    
    ' --------------------------------------------
    ' LOOP THROUGH ALL SHEETS
    ' --------------------------------------------
    For Each ws In Worksheets
    
        'Set an initial variables for ticker, Yearly change, Percent change and tot stock volume
        Dim ticker As String
        Dim ticker1 As String
        Dim YrlyChnge As Double
        Dim PctChnge As Double
        Dim TotstckVol As Double
        Dim Smry_Row As Integer
        Dim Smry_Row1 As Integer
        Dim column As Integer
        Dim openprice As Double
        Dim closeprice As Double
        Dim ColorRed As Integer
        Dim ColorGrn As Integer
        Dim rng1 As Range
        Dim rng2 As Range
        Dim max_pct As Double
        Dim min_pct As Double
        Dim max_vol As Double
    
        column = 1
        YrlyChnge = 0
        PctChnge = 0
        TotstckVol = 0
        Smry_Row = 2
        Smry_Row1 = 2
        ColorRed = 3
        ColorGrn = 4
        
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(1, 15).Value = "Ticker"
        ws.Cells(1, 16).Value = "Value"
        ws.Cells(2, 14).Value = "Greatest % Increase"
        ws.Cells(3, 14).Value = "Greatest % Decrease"
        ws.Cells(4, 14).Value = "Greatest Total Volume"
        
        
        ' Determine the Last Row
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        ' MsgBox ("The last row is " + Str(lastRow))
        
        ' Initial open price
        openprice = ws.Cells(2, 3).Value
        ' MsgBox ("This is the value of openprice " + Str(openprice))

        For i = 2 To lastRow
        ' For i = 2 To 2000
        
            ' Check to see if I am doing this right
            If ws.Cells(i + 1, column).Value <> ws.Cells(i, column).Value Then
                 ' MsgBox (Cells(i, column).Value)
                 ' MsgBox (Cells(i + 1, column).Value)
                
                ' Set ticker
                ticker = ws.Cells(i, column).Value
                
                ' Print the ticker in summary portion
                ws.Range("I" & Smry_Row).Value = ticker
                
                'Set closeprice
                closeprice = ws.Cells(i, column + 5).Value
                ' MsgBox ("This is the value of closeprice " + Str(closeprice))
                
                ' Calculate Yearly Change
                YrlyChnge = closeprice - openprice
                ' MsgBox ("This is the value of YrlyChnge " + Str(YrlyChnge))
                
                'Calculate the percent change
                If openprice = 0 Then
                    PctChnge = 0
                    
                Else
                    PctChnge = YrlyChnge / openprice
                    ' MsgBox ("This is the value of PctChnge " + Str(PctChnge))
                End If
                                
                ' Print YrlyChange
                ws.Range("J" & Smry_Row).Value = YrlyChnge
                
                ' Format cell for YrlyChnge Green or Red
                If ws.Cells(Smry_Row, 10).Value > 0 Then
                    ws.Cells(Smry_Row, 10).Interior.ColorIndex = ColorGrn
                    
                ElseIf ws.Cells(Smry_Row, 10).Value < 0 Then
                    ws.Cells(Smry_Row, 10).Interior.ColorIndex = ColorRed
                    
                End If
                
                ' Print PctChnge
                ws.Range("K" & Smry_Row).Value = PctChnge
                ws.Cells(Smry_Row, 11).NumberFormat = "0.00%"
                
                ' Add the stock volume
                TotstckVol = TotstckVol + ws.Cells(i, 7).Value
                
                ' Print the Stock Volume Total
                ws.Range("L" & Smry_Row).Value = TotstckVol
                
                ' Add 1 to the summary row
                Smry_Row = Smry_Row + 1
                
                ' Reset the TotstckVol to 0
                TotstckVol = 0
                
                ' Reset the openprice to for next ticker
                openprice = ws.Cells(i + 1, 3).Value
                ' MsgBox ("This is the value of openprice " + Str(openprice))
            
            ' If it is the same ticker in the else portion
            Else
                ' Add to the TotstckVol
                TotstckVol = TotstckVol + ws.Cells(i, 7).Value
            
            End If
            
        Next i
        
        Dim lastRow1 As Integer
        
        lastRow1 = 0
        
        lastRow1 = ws.Cells(Rows.Count, 11).End(xlUp).Row
        ' MsgBox ("The last row is " + Str(lastRow))
        Set rng1 = ws.Range("K2" & ":" & "K" & lastRow1)
        Set rng2 = ws.Range("L2" & ":" & "L" & lastRow1)
        
        min_pct = Application.WorksheetFunction.Min(rng1)
        max_pct = Application.WorksheetFunction.Max(rng1)
        max_vol = Application.WorksheetFunction.Max(rng2)
        ' MsgBox ("Min Percent is " + Str(min_pct))
        
        For q = 2 To lastRow1
            
            If ws.Cells(q, column + 10).Value = max_pct Then
            
                ticker1 = ws.Cells(q, column + 8).Value
                ws.Cells(Smry_Row1, 15).Value = ticker1
                ws.Cells(Smry_Row1, 16).Value = max_pct
                ws.Cells(Smry_Row1, 16).NumberFormat = "0.00%"
                
                ' Add 1 to the summary row
                Smry_Row1 = Smry_Row1 + 1

            End If
               
        Next q
        
        For o = 2 To lastRow1
        
            If ws.Cells(o, column + 10).Value = min_pct Then
            
                ticker1 = ws.Cells(o, column + 8).Value
                ws.Cells(Smry_Row1, 15).Value = ticker1
                ws.Cells(Smry_Row1, 16).Value = min_pct
                ws.Cells(Smry_Row1, 16).NumberFormat = "0.00%"
                
                ' Add 1 to the summary row
                Smry_Row1 = Smry_Row1 + 1

            End If
            
        Next o
        
        For Z = 2 To lastRow1
        
            If ws.Cells(Z, column + 11).Value = max_vol Then
            
                ticker1 = ws.Cells(Z, column + 8).Value
                ws.Cells(Smry_Row1, 15).Value = ticker1
                ws.Cells(Smry_Row1, 16).Value = max_vol
                
                ' Add 1 to the summary row
                Smry_Row1 = Smry_Row1 + 1

            End If
        
        Next Z
        
    ' --------------------------------------------
    ' FIXES COMPLETE
    ' --------------------------------------------
    Next ws

    MsgBox ("All Complete")

End Sub


