Attribute VB_Name = "Main"
Option Explicit

Sub multiplyOperators()
   Dim ESO As EquationSystem
   Set ESO = New EquationSystem
   ESO.fillArrays Cells(1, 2), Cells(2, 2)
   ESO.fillDegrees
   ESO.fillLetters
   ESO.prepareSolution
   ESO.fillDegreesOfDenominator
   ESO.printHeaders
   ESO.doMultiplication
   ESO.prepareSheetAfter
   Set ESO = Nothing
End Sub
