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

Dim Factors() As New Operator
Dim StringFactors() As New Operator
Dim Denominator As Operator
Dim Numerator As Operator
Dim Result As Operator

Dim Conformity() As Integer
Dim ESO As EquationSystem

Const MaxRow = 1500

' Methods

Public Sub allocateMemory(parNumberOfFactors As Integer, parNumberOfDegrees As Integer)
   NumberOfFactors = parNumberOfFactors
   NumberOfDegrees = parNumberOfDegrees
   ReDim StringFactors(NumberOfFactors - 1)
   ReDim Factors(NumberOfFactors - 1)
   Set Numerator = New Operator
   Set Denominator = New Operator
   Set Result = New Operator
End Sub

Public Sub fillFactors()
   Dim i As Integer
   For i = 0 To NumberOfFactors - 1
      StringFactors(i).allocateMemory NumberOfDegrees
      StringFactors(i).fillStringFactor 1 + i
      Factors(i).groupDegreesFromOperator StringFactors(i), Conformity
      Factors(i).groupRepetitionsFromOperator StringFactors(i), Conformity
   Next i
End Sub

Public Sub prepareEquation()
   Set ESO = New EquationSystem
   ESO.fillArrays NumberOfFactors, NumberOfDegrees
   'MsgBox ESO.getLetterInfo
   
   ESO.prepareSolution
   
End Sub

Public Sub fillDegreesOfDenominator()
   Dim FactorIndex As Integer
   Dim GroupIndex As Integer
   Dim FactorGroupIndexes() As Integer
   Denominator.allocateMemory ESO.NumberOfUnknowns
   For GroupIndex = 0 To Denominator.NumberOfGroups - 1
      FactorGroupIndexes = ESO.getLetterIndexes(GroupIndex)
      For FactorIndex = 0 To NumberOfFactors - 1
         Denominator.Degree(GroupIndex) = Denominator.Degree(GroupIndex) + Factors(FactorIndex).Degree(FactorGroupIndexes(FactorIndex))
      Next FactorIndex
      Denominator.Repetition(GroupIndex) = 1
   Next GroupIndex
   Erase FactorGroupIndexes
   Numerator.groupDegreesFromOperator Denominator, Conformity
   Result.allocateMemory NumberOfDegrees
End Sub

Public Sub printPointersOfDenominator(ByVal ColumnIndex As Integer)
   Dim i As Integer
   Dim j As Integer
   Dim FactorGroupIndexes() As Integer
   For i = 0 To Denominator.NumberOfGroups - 1
      FactorGroupIndexes = ESO.getLetterIndexes(i)
      For j = 0 To NumberOfFactors - 1
         Sheets(1).Cells(j + 1, ColumnIndex + i) = Factors(j).Degree(FactorGroupIndexes(j))
      Next j
   Next i
End Sub

Public Sub printStringFactors()
   Dim i As Integer
   For i = 0 To NumberOfFactors - 1
      Cells(i + 1, 4) = "L["
      StringFactors(i).printItemOfGroup dgDegree, False, i + 1, 5
      Cells(i + 1, NumberOfDegrees + 5) = "]"
   Next i
End Sub

Public Sub fillRepetitionsOfDenominator()
   Dim GroupIndex As Integer
   Dim UnknownArray() As Integer
   UnknownArray = ESO.getUnknownArray
   For GroupIndex = 0 To Denominator.NumberOfGroups - 1
      Denominator.Repetition(GroupIndex) = UnknownArray(GroupIndex)
   Next GroupIndex
   Erase UnknownArray
End Sub

Public Sub doMultiplication()
   Dim LastRow As Long
   Dim LastColumn As Integer
   Sheets(1).Select
   Cells(NumberOfFactors + 1, 1).EntireRow.Font.Bold = True
   Range("A1") = "Number of factors"
   Range("B1") = NumberOfFactors
   Range("A2") = "Number of degrees"
   Range("B2") = NumberOfDegrees
   printStringFactors
   LastRow = NumberOfFactors + 1
   LastColumn = NumberOfDegrees + 7
   LastColumn = Numerator.printItemOfGroup(dgDegree, False, LastRow, LastColumn) + 1
   printPointersOfDenominator LastColumn
   LastColumn = Denominator.printItemOfGroup(dgDegree, False, LastRow, LastColumn)
   Do
      LastColumn = NumberOfDegrees + 6
      ESO.fillUnknowns
      fillRepetitionsOfDenominator
      Numerator.groupRepetitionsFromOperator Denominator, Conformity
      Result.degroupDegreesFromOperator Numerator
      LastRow = LastRow + 1
      Cells(LastRow, LastColumn) = "("
      LastColumn = LastColumn + 1
      LastColumn = Numerator.printItemOfGroup(dgRepetition, True, LastRow, LastColumn)
      Cells(LastRow, LastColumn) = ") : ("
      LastColumn = LastColumn + 1
      LastColumn = Denominator.printItemOfGroup(dgRepetition, True, LastRow, LastColumn)
      Cells(LastRow, LastColumn) = ") L["
      LastColumn = LastColumn + 1
      LastColumn = Result.printItemOfGroup(dgDegree, False, LastRow, LastColumn)
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
   Dim FactorIndex As Integer
   Dim Index As Integer
   For FactorIndex = 0 To NumberOfFactors - 1
      Set StringFactors(FactorIndex) = Nothing
      Set Factors(FactorIndex) = Nothing
   Next FactorIndex
   Set Denominator = Nothing
   Set Numerator = Nothing
   Set Result = Nothing
   Set ESO = Nothing
   Erase Factors
   Erase StringFactors
   Erase Conformity
End Sub
