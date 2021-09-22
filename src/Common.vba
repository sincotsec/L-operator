Attribute VB_Name = "Common"
Option Explicit

Enum dgItem
   dgRepetition = 0
   dgDegree = 1
End Enum

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
