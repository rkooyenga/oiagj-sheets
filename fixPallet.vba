'fixPallet by Ray Kooyenga
'creates pallet sort id to correct sort
'fixes: now adjusts for pallets ending in A and AA but needs verification that result is optimal for staff
Sub fixPallet()
Dim s As String
Dim FirstDash As Long
Dim LastDash As Long
Dim Row As Long
Dim LastRow As Long
Dim Column As Long
Dim newStr As String
Dim endStr As String
Dim endStrLength As Long
Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets("ALL INV")

Column = 9

'get last row
ws.Range("H1").Select
Selection.End(xlDown).Select
LastRow = ActiveCell.Row

For Row = 2 To LastRow
    s = ws.Cells(Row, 6).Value
    If s Like "*-*" Then
            FirstDash = InStr(1, s, "-")
            LastDash = InStrRev(s, "-")
            endStr = Right(s, Len(s) - FirstDash)
            endStrLength = Len(endStr)
            Select Case True
              Case endStr Like "*AA"
                Select Case endStrLength
                   Case 3
                      endStr = "000" & endStr
                   Case 4
                      endStr = "00" & endStr
                   Case 5
                      endStr = "0" & endStr
                End Select
              Case endStr Like "*A"
                Select Case endStrLength
                   Case 2
                      endStr = "000" & endStr
                   Case 3
                      endStr = "00" & endStr
                   Case 4
                      endStr = "0" & endStr
                End Select
              Case Else
                Select Case endStrLength
                   Case 1
                      endStr = "000" & endStr
                   Case 2
                      endStr = "00" & endStr
                   Case 3
                      endStr = "0" & endStr
                End Select
            End Select
            newStr = Left(s, LastDash - 1) & "-" & endStr
            ws.Cells(Row, Column).Value = newStr
        Else
            ws.Cells(Row, Column).Value = s
        End If
Next

'tidy up
ws.Range("A:J").Columns.AutoFit
ws.Range("A:I").Borders.LineStyle = xlContinuous

With ws.Range("I:I")
    .NumberFormat = "General"
    .HorizontalAlignment = xlCenter
End With

With ws.Range("I1")
    .Value = "PalletSortID"
    .Interior.ColorIndex = 50
    .Font.Color = vbWhite
    .Font.Bold = True
End With

'new sort
With ws.Sort
     .SortFields.Clear
     .SortFields.Add Key:=ws.Range("C1"), Order:=xlAscending
     .SortFields.Add Key:=ws.Range("D1"), Order:=xlAscending
     .SortFields.Add Key:=ws.Range("I1"), Order:=xlAscending
     .SetRange ws.Range("A1:I" & LastRow)
     .Header = xlYes
     .Apply
End With

ws.Range("A1").Select
Selection.End(xlUp).Select

End Sub
