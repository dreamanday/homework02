Sub stocks()

Dim a As Long
Dim a2 As Long
Dim b As String
Dim c As Integer
Dim d As String
Dim e As Double
Dim f As Double

Range("I1").Value = "Ticker"
Range("J1").Value = "Yearly Change"
Range("K1").Value = "Percent Change"
Range("L1").Value = "Total Stock Volume"

Range("A2").Select
Selection.End(xlDown).Select
a = ActiveCell.Row

Range("A2:A" & a).Copy
Range("I2").PasteSpecial
ActiveSheet.Range("$I$2:$I" & a).RemoveDuplicates Columns:=1, Header:=xlNo

Range("I2").Select
Selection.End(xlDown).Select
a2 = ActiveCell.Row

Range("A2").Select

For t = 1 To a2

b = Range("I" & t + 1).Value

For q = 1 To a

If ActiveCell.Value = b Then
c = c + 1
ActiveCell.Offset(1, 0).Select
Else
Exit For
End If
Next q

e = ActiveCell.Offset(-c, 2).Value
f = ActiveCell.Offset(-1, 5).Value
Range("J" & t + 1).Value = f - e
    If Range("J" & t + 1).Value < 0 Then
    Range("J" & t + 1).Interior.ColorIndex = 3
    Else
    Range("J" & t + 1).Interior.ColorIndex = 4
    End If
Range("K" & t + 1).Value = (1 - (f / e)) * -1
    Range("K" & t + 1).Style = "Percent"
    Range("K" & t + 1).NumberFormat = "0.00%"
Range("L" & t + 1).Value = Application.Sum(Range(Cells(ActiveCell.Row - c, 7), Cells(ActiveCell.Row - 1, 7)))

c = 0
Next t


End Sub

Sub scores()

Dim a2 As Long
Dim b As Double
Dim c As Double
Dim d As Variant
Dim e As String

Range("N3").Value = "Greatest % of Increase"
Range("N4").Value = "Greatest % of Decrease"
Range("N5").Value = "Greatest Total Volume"
Range("O2").Value = "Ticker"
Range("P2").Value = "Value"

Range("I2").Select
Selection.End(xlDown).Select
a2 = ActiveCell.Row

b = Application.WorksheetFunction.Max(Range("K2:K" & a2))
Range("P3").Value = b

c = Application.WorksheetFunction.Min(Range("K2:K" & a2))
Range("P4").Value = c

d = Application.WorksheetFunction.Max(Range("L2:L" & a2))
Range("P5").Value = d

Range("K2").Select

For i = 1 To a2

If ActiveCell.Value = b Then
e = ActiveCell.Offset(0, -2).Value
Range("O3").Value = e
Exit For
Else
ActiveCell.Offset(1, 0).Select
End If
Next i

Range("K2").Select

For i = 1 To a2

If ActiveCell.Value = c Then
e = ActiveCell.Offset(0, -2).Value
Range("O4").Value = e
Exit For
Else
ActiveCell.Offset(1, 0).Select
End If
Next i

Range("L2").Select

For i = 1 To a2

If ActiveCell.Value = d Then
e = ActiveCell.Offset(0, -3).Value
Range("O5").Value = e
Exit For
Else
ActiveCell.Offset(1, 0).Select
End If
Next i

End Sub




