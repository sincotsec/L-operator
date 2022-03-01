Attribute VB_Name = "Main"
Option Explicit

Sub multiplyOperators()
   Dim Cls As Union
   Set Cls = New Union
   Cls.allocateMemory Sheets(1).Cells(1, 2), Sheets(1).Cells(2, 2)
   Cls.fillFactors
   Cls.prepareEquation
   Cls.fillDegreesOfDenominator
   Cls.doMultiplication
   Cls.prepareSheetAfter
   Set Cls = Nothing
End Sub
