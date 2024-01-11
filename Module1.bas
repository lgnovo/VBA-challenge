Attribute VB_Name = "Module1"
Sub alphatest()
    Dim hw As Worksheet
    Dim i As Long
    Dim Beginning As Long
    Dim Total As Double
    Dim j As Integer
    Dim Perctchange As Single
    Dim RowCt As Long
    Dim Change As Single
    Dim Alpha As Variant
    Dim Start As Variant

For Each hw In Worksheets

    Total = 0
    j = 0
    Beginning = 2
    Change = 0

    hw.Range("I1").Value = "Ticker"
    hw.Range("J1").Value = "Yearly Change"
    hw.Range("K1").Value = "Percent Change"
    hw.Range("L1").Value = "Total Stock Volume"
    hw.Range("O2").Value = "Greatest % Increase"
    hw.Range("O3").Value = "Greatest % Decrease"
    hw.Range("O4").Value = "Greatest Total Volume"
    hw.Range("P1").Value = "Ticker"
    hw.Range("Q1").Value = "Value"

    RowCt = hw.Cells(Rows.Count, "A").End(xlUp).Row

    For i = 2 To RowCt
        If hw.Cells(i + 1, 1).Value <> hw.Cells(i, 1).Value Then
            Total = Total + hw.Cells(i, 7).Value
            Alpha = hw.Cells(i, 1).Value
            Start = hw.Cells(Beginning, 3).Value

            If Total = 0 Then
                hw.Range("I" & 2 + j).Value = Alpha
                hw.Range("J" & 2 + j).Value = 0
                hw.Range("K" & 2 + j).Value = "%" & 0
                hw.Range("L" & 2 + j).Value = 0
            Else
                If Start = 0 Then
                    For find_value = Beginning To i
                        If hw.Cells(find_value, 3).Value <> 0 Then
                            Beginning = find_value
                            Exit For
                        End If
                    Next find_value
                End If

                Change = hw.Cells(i, 6).Value - Start
                If Start <> 0 Then
                    Perctchange = Round((Change / Start) * 100, 2)
                Else
                    Perctchange = 0
                End If

                hw.Range("I" & 2 + j).Value = Alpha
                hw.Range("J" & 2 + j).Value = Round(Change, 2)
                hw.Range("K" & 2 + j).Value = "%" & Perctchange
                hw.Range("L" & 2 + j).Value = Total

    Select Case Change
         Case Is > 0
            hw.Range("J" & 2 + j).Interior.ColorIndex = 4
        Case Is < 0
            hw.Range("J" & 2 + j).Interior.ColorIndex = 3
        Case Else
            hw.Range("J" & 2 + j).Interior.ColorIndex = 0
    End Select
    End If


            Total = 0
            Change = 0
            j = j + 1
            Beginning = i + 1
        Else
            Total = Total + hw.Cells(i, 7).Value
        End If
    Next i

    Dim Kmax As Variant
    Dim Kmin As Variant
    Dim Lmax As Variant
    RowCount = RowCt - 1
    
    Kmax = WorksheetFunction.Max(hw.Range("K2:K" & RowCount))
    Kmin = WorksheetFunction.Min(hw.Range("K2:K" & RowCount))
    Lmax = WorksheetFunction.Max(hw.Range("L2:L" & RowCount))
    
    hw.Range("Q2").Value = "%" & Kmax * 100
    hw.Range("Q3").Value = "%" & Kmin * 100
    hw.Range("Q4").Value = Lmax

    ticker1 = WorksheetFunction.Match(Kmax, hw.Range("K2:K" & RowCount), 0)
    ticker2 = WorksheetFunction.Match(Kmin, hw.Range("K2:K" & RowCount), 0)
    ticker3 = WorksheetFunction.Match(Lmax, hw.Range("L2:L" & RowCount), 0)
    
    hw.Range("P2") = hw.Cells(ticker1 + 1, 9)
    hw.Range("P3") = hw.Cells(ticker2 + 1, 9)
    hw.Range("P4") = hw.Cells(ticker3 + 1, 9)
  

  Next hw
  
End Sub

