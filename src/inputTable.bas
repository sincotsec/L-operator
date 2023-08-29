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

Private Sub fillTitle(numberOfFactors As Integer, numberOfDegrees As Integer)
   Range("A1") = "Number of factors"
   Range("B1") = numberOfFactors
   Range("A2") = "Number of degrees"
   Range("B2") = numberOfDegrees
End Sub

Private Sub fillString(factorIndex As Integer, numberOfDegrees As Integer)
   Dim i As Integer
   Cells(factorIndex, 4) = "L["
   Range(Cells(factorIndex, 5), Cells(factorIndex, numberOfDegrees + 5)) = 0
   Cells(factorIndex, numberOfDegrees + 5) = "]"
End Sub

Public Sub redrawTable()
   Dim numberOfFactors As Integer
   Dim numberOfDegrees As Integer
   Dim i As Integer
   numberOfFactors = Cells(1, 2)
   numberOfDegrees = Cells(2, 2)
   Call prepareSheetBefore
   Call fillTitle(numberOfFactors, numberOfDegrees)
   For i = 1 To numberOfFactors
      Call fillString(i, numberOfDegrees)
   Next i
   Cells.EntireColumn.AutoFit
End Sub
