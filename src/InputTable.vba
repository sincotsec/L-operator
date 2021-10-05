Attribute VB_Name = "InputTable"
Option Explicit

Sub fillInputSheet()
   If ActiveWorkbook.Worksheets.Count = 1 Then ActiveWorkbook.Worksheets.Add
   Sheets(1).Select
   Sheets(1).Name = "String Factors"
   Sheets(2).Name = "Result"
   Cells.Clear
   Call fillTitle(2, 9)
   Call redrawTable
End Sub

Private Sub fillTitle(NumberOfFactors As Integer, NumberOfDegrees As Integer)
   With Range("A1:B2")
      .Borders(xlEdgeRight).LineStyle = xlContinuous
      .Font.Name = "Arial Narrow"
      .Font.Size = 18
      .HorizontalAlignment = xlCenter
      .VerticalAlignment = xlCenter
   End With
   Range(Cells(1, 1), Cells(2, NumberOfDegrees + 2)).Borders(xlEdgeBottom).LineStyle = xlContinuous
   Range("A1") = "Number of factors": Range("B1") = NumberOfFactors
   Range("A2") = "Number of degrees": Range("B2") = NumberOfDegrees
   Range("A1:A2").Font.Size = 12
   Cells.EntireColumn.AutoFit
   Range("B1").ColumnWidth = 5
End Sub

Private Sub fillString(FactorIndex As Integer, NumberOfDegrees As Integer)
   Dim i As Integer
   Sheets(1).Select
   With Range(Cells(FactorIndex + 2, 1), Cells(FactorIndex + 2, NumberOfDegrees + 2))
      .HorizontalAlignment = xlCenter
      .VerticalAlignment = xlCenter
      .Font.Name = "Arial Narrow"
      .Font.Size = 18
   End With
   Cells(FactorIndex + 2, 2).Borders(xlEdgeRight).LineStyle = xlContinuous
   Cells(FactorIndex + 2, 1) = "Factor " & FactorIndex
   Range(Cells(FactorIndex + 2, 3), Cells(FactorIndex + 2, NumberOfDegrees + 2)) = 0
End Sub

Sub redrawTable()
   Sheets(1).Select
   Dim NumberOfFactors As Integer
   Dim NumberOfDegrees As Integer
   Dim i As Integer
   NumberOfFactors = Cells(1, 2)
   NumberOfDegrees = Cells(2, 2)
   Cells.Clear
   Call fillTitle(NumberOfFactors, NumberOfDegrees)
   For i = 1 To NumberOfFactors
      Call fillString(i, NumberOfDegrees)
   Next i
   Range(Cells(1, 3), Cells(1, NumberOfDegrees + 2)).ColumnWidth = 5
End Sub
