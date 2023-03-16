Attribute VB_Name = "Module1"
Sub WallStreet()

Dim ws As Worksheet
Dim WorksheetName As String

For Each ws In ThisWorkbook.Worksheets

Dim Ticker As String
Dim lastrow, i, Printer As Long
Dim Volume As Double
Printer = 2
'to make sure all the values print at the same place (row 2) on each ws
Volume = 0
lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        'Debug.Print (WorksheetName)
WorksheetName = ws.Name
        'Debug.Print (WorksheetName)'

            ws.Range("I1").Value = "Ticker"
            ws.Range("J1").Value = "Open"
            ws.Range("K1").Value = "Close"
            ws.Range("L1").Value = "Yearly Change"
            ws.Range("M1").Value = "Percentage Change"
            ws.Range("N1").Value = "Total Stock Volume"


            For i = 2 To lastrow
            
                Ticker = ws.Cells(i, 1).Value
                    If ws.Cells(i - 1, 1).Value <> Ticker Then
                    'to get the open value for each ticker
                        ws.Range("I" & Printer).Value = Ticker
                        ws.Range("J" & Printer).Value = ws.Cells(i, 3).Value
                        Volume = ws.Cells(i, 7)
                    ElseIf ws.Cells(i + 1, 1).Value <> Ticker Then
                    'to get the closing value for each ticker
                        ws.Range("K" & Printer).Value = ws.Cells(i, 6).Value
                        ws.Range("N" & Printer).Value = Volume
                        Printer = Printer + 1
                        Volume = 0
                        
                    Else: Volume = Volume + ws.Cells(i, 7).Value
                    End If
            
            Next i

'another for loop to get Yearly Change and Percentage Change

Dim j As Integer
lastrow = ws.Cells(Rows.Count, 9).End(xlUp).Row
' Count, 9 bc only working with the new rows starting at column 9
    'Debug.Print (lastrow)

                For j = 2 To lastrow
                    ws.Cells(j, 12) = ws.Cells(j, 11) - ws.Cells(j, 10)
                    ws.Cells(j, 13) = ws.Cells(j, 12) / ws.Cells(j, 10)
                Next j
        
Next ws
    
End Sub

Sub WSformatting()

Dim ws As Worksheet
Dim WorksheetName As String

For Each ws In ThisWorkbook.Worksheets

Dim j As Integer
lastrow = ws.Cells(Rows.Count, 9).End(xlUp).Row
WorksheetName = ws.Name

    For j = 2 To lastrow
        ws.Cells(j, 13) = FormatPercent(Cells(j, 13))
        If ws.Cells(j, 12) <= 0 Then
            ws.Cells(j, 12).Interior.ColorIndex = 3
        ElseIf ws.Cells(j, 12) > 0 Then
            ws.Cells(j, 12).Interior.ColorIndex = 4
        'Else: ws.Cells(j, 12).Interior.ColorIndex = 6
        'in case share holders wanted to see where there was no net loss or gain aka where total stock volume = 0
        End If
    Next j
    
Next ws

End Sub
Sub OpenClose()
'removes "Open" and "Close" columns to match hw image and make the sheet look cleaner

Dim ws As Worksheet
Dim WorksheetName As String

For Each ws In ThisWorkbook.Worksheets
WorksheetName = ws.Name
  ws.Columns("J:K").Delete
Next ws

End Sub
