Attribute VB_Name = "Main"
Option Explicit

Sub multiplyOperators()
   Dim ESO As EquationSystem
   Set ESO = New EquationSystem
   ESO.fillArrays Sheets(1).Cells(1, 2), Sheets(1).Cells(2, 2)
   ESO.prepareSolution
   ESO.fillDegreesOfDenominator
   ESO.doMultiplication
   Set ESO = Nothing
End Sub
