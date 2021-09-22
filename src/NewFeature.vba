Attribute VB_Name = "NewFeature"
Option Explicit

Sub testNewFeature()
   Application.Calculation = xlCalculationManual
   Application.ScreenUpdating = False
   Application.EnableEvents = False
   Dim Cls As Multiplication
   Set Cls = New Multiplication
   Cls.prepareSheetBefore
   Cls.allocateMemory Sheets(1).Cells(1, 2), Sheets(1).Cells(2, 2)
   Cls.fillDegreesOfDenominator
   Cls.setColumns
   'Cls.doMultiplication
   Cls.prepareSheetAfter
   Set Cls = Nothing
   Application.Calculation = xlCalculationAutomatic
   Application.ScreenUpdating = True
   Application.EnableEvents = True
End Sub
