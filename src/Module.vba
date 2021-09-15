Attribute VB_Name = "Module"
Option Explicit

Enum dgItem
   dgRepetition = 0
   dgDegree = 1
End Enum

Type Group
   Degree As Integer
   Repetition As Integer
End Type

Function getMaximum(Number1 As Integer, Number2 As Integer) As Integer
   getMaximum = Number2
   If Number1 >= Number2 Then getMaximum = Number1
End Function

Function getMinimum(Number1 As Integer, Number2 As Integer) As Integer
   getMinimum = Number2
   If Number1 <= Number2 Then getMinimum = Number1
End Function

Function ColorFromHSL(ByVal H As Double, S As Double, L As Double) As Long
   ' H in [0, 360]
   ' S in [0, 100]
   ' L in [0, 100]
   Dim Q As Double
   Dim P As Double
   Dim Tr As Double
   Dim Tg As Double
   Dim Tb As Double
   Dim R, G, B As Double
   H = H / 360
   S = S / 100
   L = L / 100
   If L < 0.5 Then
      Q = L * (1 + S)
   Else
      Q = L + S - (L * S)
   End If
   P = 2 * L - Q
   Tr = H + 0.3333333333333
   Tg = H
   Tb = H - 0.3333333333333
   If Tr < 0 Then Tr = Tr + 1
   If Tr > 1 Then Tr = Tr - 1
   If Tr < 0.16666666666666 Then
      R = P + ((Q - P) * 6 * Tr)
   ElseIf Tr >= 0.16666666666666 And Tr < 0.5 Then
      R = Q
   ElseIf Tr >= 0.5 And Tr < 0.6666666666666 Then
      R = P + ((Q - P) * (0.6666666666666 - Tr) * 6)
   Else
      R = P
   End If
   If Tg < 0 Then Tg = Tg + 1
   If Tg > 1 Then Tg = Tg - 1
   If Tg < 0.16666666666666 Then
      G = P + ((Q - P) * 6 * Tg)
   ElseIf Tg >= 0.16666666666666 And Tg < 0.5 Then
      G = Q
   ElseIf Tg >= 0.5 And Tg < 0.6666666666666 Then
      G = P + ((Q - P) * (0.6666666666666 - Tg) * 6)
   Else
      G = P
   End If
   If Tb < 0 Then Tb = Tb + 1
   If Tb > 1 Then Tb = Tb - 1
   If Tb < 0.16666666666666 Then
      B = P + ((Q - P) * 6 * Tb)
   ElseIf Tb >= 0.16666666666666 And Tb < 0.5 Then
      B = Q
   ElseIf Tb >= 0.5 And Tb < 0.6666666666666 Then
      B = P + ((Q - P) * (0.6666666666666 - Tb) * 6)
   Else
      B = P
   End If
   R = R * 255
   G = G * 255
   B = B * 255
   ColorFromHSL = RGB(R, G, B)
End Function

Sub fillInputSheet()
   If ActiveWorkbook.Worksheets.Count = 1 Then ActiveWorkbook.Worksheets.Add
   Sheets(1).Select
   Sheets(1).Name = "StringFactors"
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

Private Sub prepareOutputSheet()
   ActiveWindow.WindowState = xlMaximized
   ActiveWindow.FreezePanes = False
   With Sheets(2).Cells
      .Clear
      .ColumnWidth = 2
      .Interior.Pattern = xlNone
      .Interior.Color = ColorFromHSL(WorksheetFunction.RandBetween(0, 360), 70, 40)
      .Font.ColorIndex = xlAutomatic
      .Font.Bold = False
      .Font.Size = 15
      .Font.Name = "Arial Narrow"
      .HorizontalAlignment = xlCenter
      .VerticalAlignment = xlCenter
   End With
End Sub

Sub multiplyOperators()
   Application.Calculation = xlCalculationManual
   Application.ScreenUpdating = False
   Application.EnableEvents = False
   Call prepareOutputSheet
   Dim Cls As Multiplication
   Set Cls = New Multiplication
   Cls.allocateMemory Sheets(1).Cells(1, 2), Sheets(1).Cells(2, 2)
   Cls.fillDegreesOfDenominator
   Cls.setColumns
   Cls.doMultiplication
   Cls.prepareSheetAfter
   Set Cls = Nothing
   Application.Calculation = xlCalculationAutomatic
   Application.ScreenUpdating = True
   Application.EnableEvents = True
End Sub
