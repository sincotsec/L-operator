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

'Dim Factors() As New Operator
'Dim StringFactors() As New Operator
'Dim Denominator As Operator
'Dim Numerator As Operator
'Dim Result As Operator

'Dim Conformity() As Integer
Dim ESO As EquationSystem

Const MaxRow = 1500

' Methods

Public Sub allocateMemory(parNumberOfFactors As Integer, parNumberOfDegrees As Integer)
   NumberOfFactors = parNumberOfFactors
   NumberOfDegrees = parNumberOfDegrees
'   ReDim StringFactors(NumberOfFactors - 1)
'   ReDim Factors(NumberOfFactors - 1)
'   Set Numerator = New Operator
'   Set Denominator = New Operator
'   Set Result = New Operator
End Sub

'Public Sub fillFactors()
'   Dim i As Integer
'   For i = 0 To NumberOfFactors - 1
'      StringFactors(i).allocateMemory NumberOfDegrees
'      StringFactors(i).fillStringFactor 1 + i
'      Factors(i).groupDegreesFromOperator StringFactors(i), Conformity
'      Factors(i).groupRepetitionsFromOperator StringFactors(i), Conformity
'   Next i
'End Sub

Public Sub prepareEquation()
   Set ESO = New EquationSystem
   ESO.fillArrays NumberOfFactors, NumberOfDegrees
   ESO.prepareSolution
   ESO.fillDegreesOfDenominator
End Sub

'Public Sub fillDegreesOfDenominator()
'   Dim FactorIndex As Integer
'   Dim GroupIndex As Integer
'   Dim FactorGroupIndexes() As Integer
'   Denominator.allocateMemory ESO.NumberOfUnknowns
'   For GroupIndex = 0 To Denominator.NumberOfGroups - 1
'      FactorGroupIndexes = ESO.getLetterIndexes(GroupIndex)
'      For FactorIndex = 0 To NumberOfFactors - 1
'         Denominator.Degree(GroupIndex) = Denominator.Degree(GroupIndex) + Factors(FactorIndex).Degree(FactorGroupIndexes(FactorIndex))
'      Next FactorIndex
'      Denominator.Repetition(GroupIndex) = 1
'   Next GroupIndex
'   Erase FactorGroupIndexes
'   Numerator.groupDegreesFromOperator Denominator, Conformity
'   Result.allocateMemory NumberOfDegrees
'End Sub

'Public Sub printStringFactors()
'   Dim i As Integer
'   For i = 0 To NumberOfFactors - 1
'      Cells(i + 1, 4) = "L["
'      StringFactors(i).printItemOfGroup dgDegree, False, i + 1, 5
'      Cells(i + 1, NumberOfDegrees + 5) = "]"
'   Next i
'End Sub

'Public Sub fillRepetitionsOfDenominator()
'   Dim GroupIndex As Integer
'   Dim UnknownArray() As Integer
'   UnknownArray = ESO.getUnknownArray
'   For GroupIndex = 0 To Denominator.NumberOfGroups - 1
'      Denominator.Repetition(GroupIndex) = UnknownArray(GroupIndex)
'   Next GroupIndex
'   Erase UnknownArray
'End Sub

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
'   Dim FactorIndex As Integer
'   Dim Index As Integer
'   For FactorIndex = 0 To NumberOfFactors - 1
'      Set StringFactors(FactorIndex) = Nothing
'      Set Factors(FactorIndex) = Nothing
'   Next FactorIndex
'   Set Denominator = Nothing
'   Set Numerator = Nothing
'   Set Result = Nothing
   Set ESO = Nothing
'   Erase Factors
'   Erase StringFactors
'   Erase Conformity
End Sub
