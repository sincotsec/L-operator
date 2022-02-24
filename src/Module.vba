Attribute VB_Name = "Module"
Option Explicit

Sub multiplyOperators()
   Application.Calculation = xlCalculationManual
   Application.ScreenUpdating = False
   Application.EnableEvents = False
   Dim Cls As Union
   Set Cls = New Union
   Cls.prepareSheetBefore
   Cls.allocateMemory Sheets(1).Cells(1, 2), Sheets(1).Cells(2, 2)
   Cls.fillFactors
   Cls.prepareEquation
   Cls.fillDegreesOfDenominator
   Cls.doMultiplication
   Cls.prepareSheetAfter
   Set Cls = Nothing
   Application.Calculation = xlCalculationAutomatic
   Application.ScreenUpdating = True
   Application.EnableEvents = True
End Sub
