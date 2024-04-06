Attribute VB_Name = "Module1"
Sub Alphabetical()

    'variable declaration
    Dim i As Long
    Dim ticker As String
    Dim tablerow As Double
    Dim openprice, closeprice, yearlychange As Double
    Dim percentchange As Double
    Dim totalstockvolume As Double
    Dim maxincrease As Double
    Dim maxincreaseticker As String
    
    'headers
    Cells(1, "I") = "Ticker"
    Cells(1, "J") = "Yearly Change"
    Cells(1, "K") = "Percent Change"
    Cells(1, "L") = "Total Stock Volume"
    Cells(1, "P") = "Ticker"
    Cells(1, "Q") = "Value"
    Cells(2, "O") = "Greatest % Increase"
    Cells(3, "O") = "Greatest % Decrease "
    
    'variable initialization'
    tablerow = 2
    lastrow = Cells(Rows.Count, "A").End(xlUp).Row
    openprice = Cells(2, "C")
    totalstockvolume = 0
    Greatestincrease = 0
    
    
    For i = 2 To lastrow
            totalstockvolume = totalstockvolume + Cells(i, "G")
            If Cells(i, "A") <> Cells(i + 1, "A") Then
                ticker = Cells(i, "A")
                Cells(tablerow, "I") = ticker
                closeprice = Cells(i, "F")
                yearlychange = closeprice - openprice
                Cells(tablerow, "J") = yearlychange
                If Cells(tablerow, "J") > 0 Then
                    Cells(tablerow, "J").Interior.ColorIndex = 4
                Else
                    Cells(tablerow, "J").Interior.ColorIndex = 3
                End If
                If openprice <> 0 Then
                    percentchange = yearlychange / openprice
                Else
                    percentchange = 0
                End If
                Cells(tablerow, "K") = FormatPercent(percentchange, 2)
                Cells(tablerow, "L") = totalstockvolume
                
                If Cells(tablerow, "K") > maxincrease Then
                    maxincrease = Cells(tablerow, "K")
                    maxincreaseticker = Cells(tablerow, "I")
                End If
                
                
            openprice = Cells(i + 1, "C")
            tablerow = tablerow + 1
            totalstockvolume = 0
            End If
    Next i

End Sub
