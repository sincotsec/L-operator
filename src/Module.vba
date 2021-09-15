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

Function ColorFromHSL(ByVal H As Double, S As Double, L As Double) As Long ' H from [0, 360], S from [0, 100], L from [0, 100]
   Dim i As Integer
   Dim Q As Double
   Dim P As Double
   Const OneThird = 0.33333333333333
   Const OneSixth = 0.16666666666666
   Const TwoThirds = 0.66666666666666
   Dim Tcolor(3) As Double
   Dim RGBcolor(3) As Double
   H = H / 360
   S = S / 100
   L = L / 100
   If L < 0.5 Then
      Q = L * (1 + S)
   Else
      Q = L + S - (L * S)
   End If
   P = 2 * L - Q
   Tcolor(0) = H + OneThird
   Tcolor(1) = H
   Tcolor(2) = H - OneThird
   For i = 0 To 2
      If Tcolor(i) < 0 Then Tcolor(i) = Tcolor(i) + 1
      If Tcolor(i) > 1 Then Tcolor(i) = Tcolor(i) - 1
      If Tcolor(i) < OneSixth Then
         RGBcolor(i) = P + ((Q - P) * 6 * Tcolor(i))
      ElseIf Tcolor(i) >= OneSixth And Tcolor(i) < 0.5 Then
         RGBcolor(i) = Q
      ElseIf Tcolor(i) >= 0.5 And Tcolor(i) < TwoThirds Then
         RGBcolor(i) = P + ((Q - P) * (TwoThirds - Tcolor(i)) * 6)
      Else
         RGBcolor(i) = P
      End If
      RGBcolor(i) = RGBcolor(i) * 255
   Next i
   ColorFromHSL = RGB(RGBcolor(0), RGBcolor(1), RGBcolor(2))
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

Sub multiplyOperators()
   Application.Calculation = xlCalculationManual
   Application.ScreenUpdating = False
   Application.EnableEvents = False
   Dim Cls As Multiplication
   Set Cls = New Multiplication
   Cls.prepareSheetBefore
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
