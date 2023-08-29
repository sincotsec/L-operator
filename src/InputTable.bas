Attribute VB_Name = "inputTable"
Option Explicit

Public Sub prepareSheetBefore()
   ActiveWindow.WindowState = xlMaximized
   ActiveWindow.FreezePanes = False
   With Cells
      .Clear
      .ColumnWidth = 2
      .Interior.Pattern = xlNone
      .Font.ColorIndex = xlAutomatic
      .Font.Bold = False
      .Font.Size = 15
      .Font.Name = "Century Gothic"
      .HorizontalAlignment = xlCenter
      .VerticalAlignment = xlCenter
   End With
End Sub

Public Sub fillInputSheet()
   Sheets(1).Name = "L-operator"
   Call prepareSheetBefore
   Call fillTitle(2, 9)
   Call redrawTable
End Sub

Private Sub fillTitle(NumberOfFactors As Integer, NumberOfDegrees As Integer)
   Range("A1") = "Number of factors"
   Range("B1") = NumberOfFactors
   Range("A2") = "Number of degrees"
   Range("B2") = NumberOfDegrees
End Sub

Private Sub fillString(FactorIndex As Integer, NumberOfDegrees As Integer)
   Dim i As Integer
   Cells(FactorIndex, 4) = "L["
   Range(Cells(FactorIndex, 5), Cells(FactorIndex, NumberOfDegrees + 5)) = 0
   Cells(FactorIndex, NumberOfDegrees + 5) = "]"
End Sub

Public Sub redrawTable()
   Dim NumberOfFactors As Integer
   Dim NumberOfDegrees As Integer
   Dim i As Integer
   NumberOfFactors = Cells(1, 2)
   NumberOfDegrees = Cells(2, 2)
   Call prepareSheetBefore
   Call fillTitle(NumberOfFactors, NumberOfDegrees)
   For i = 1 To NumberOfFactors
      Call fillString(i, NumberOfDegrees)
   Next i
   Cells.EntireColumn.AutoFit
End Sub
