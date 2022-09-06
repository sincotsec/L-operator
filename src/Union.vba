VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Union"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim NumberOfFactors As Integer
Dim NumberOfDegrees As Integer
Dim ESO As EquationSystem

Const MaxRow = 1500

' Methods

Public Sub allocateMemory(parNumberOfFactors As Integer, parNumberOfDegrees As Integer)
   NumberOfFactors = parNumberOfFactors
   NumberOfDegrees = parNumberOfDegrees
End Sub

Public Sub prepareEquation()
   Set ESO = New EquationSystem
   ESO.fillArrays NumberOfFactors, NumberOfDegrees
   ESO.prepareSolution
   ESO.fillDegreesOfDenominator
End Sub

Public Sub doMultiplication()
   Dim LastRow As Long
   Dim LastColumn As Integer
   Sheets(1).Select
   Call prepareSheetBefore
   Cells(NumberOfFactors + 1, 1).EntireRow.Font.Bold = True
   Range("A1") = "Number of factors"
   Range("B1") = NumberOfFactors
   Range("A2") = "Number of degrees"
   Range("B2") = NumberOfDegrees
   ESO.printUngroupedDegrees
   
   LastRow = NumberOfFactors + 1
   LastColumn = NumberOfDegrees + 7
   ESO.printNumeratorDegrees LastRow, LastColumn
   LastColumn = LastColumn + ESO.getNumberOfNumeratorDegrees + 1
   ESO.printPointersOfDenominator LastColumn
   ESO.printDenominatorDegrees LastRow, LastColumn
   LastColumn = LastColumn + ESO.NumberOfUnknowns
   Do
      LastColumn = NumberOfDegrees + 6
      ESO.fillUnknowns
      ESO.groupRepetitionsFromDenominator
      ESO.fillDegreesOfResult
      LastRow = LastRow + 1
      Cells(LastRow, LastColumn) = "("
      ESO.printNumeratorRepetitions LastRow, LastColumn + 1
      LastColumn = LastColumn + ESO.getNumberOfNumeratorDegrees + 1
      Cells(LastRow, LastColumn) = ") : ("
      ESO.printUnknowns LastRow, LastColumn + 1
      LastColumn = LastColumn + ESO.NumberOfUnknowns + 1
      Cells(LastRow, LastColumn) = ") L["
      ESO.printResultDegrees LastRow, LastColumn + 1
      LastColumn = LastColumn + NumberOfDegrees + 1
      Cells(LastRow, LastColumn) = "]"
   Loop Until (LastRow >= MaxRow Or ESO.isDone())
End Sub

Public Sub prepareSheetAfter()
   Sheets(1).Select
   ActiveWindow.WindowState = xlMaximized
   ActiveWindow.ScrollColumn = 1
   Cells(NumberOfFactors + 2, 1).Select
   ActiveWindow.FreezePanes = False
   ActiveWindow.FreezePanes = True
   Sheets(1).Cells.EntireColumn.AutoFit
End Sub

' Destructor

Private Sub Class_Terminate()
   Set ESO = Nothing
End Sub
