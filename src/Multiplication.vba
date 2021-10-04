VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Multiplication"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim NumberOfFactors As Integer
Dim NumberOfDegrees As Integer

Dim Factors() As New Operator
Dim StringFactors() As New Operator
Dim ColumnOperators() As Operator
Dim Denominator As Operator
Dim Numerator As Operator
Dim Result As Operator

Dim Conformity() As Integer
Dim EqObj As Equation

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
   ReDim ColumnOperators(3)
   Set ColumnOperators(0) = StringFactors(0)
   Set ColumnOperators(1) = Numerator
   Set ColumnOperators(2) = Denominator
   Set ColumnOperators(3) = Result
End Sub

Public Sub fillFactors()
   Dim i As Integer
   For i = 0 To NumberOfFactors - 1
      StringFactors(i).allocateMemory NumberOfDegrees
      StringFactors(i).fillStringFactor 3 + i
      Factors(i).groupDegreesFromOperator StringFactors(i), Conformity
      Factors(i).groupRepetitionsFromOperator StringFactors(i), Conformity
   Next i
End Sub

Public Sub prepareEquation()
   Set EqObj = New Equation
   EqObj.allocateMemory NumberOfFactors, NumberOfDegrees
   EqObj.fillArray Factors
   EqObj.prepareSolution
End Sub

Public Sub fillDegreesOfDenominator()
   Dim FactorIndex As Integer
   Dim GroupIndex As Integer
   Dim FactorGroupIndexes() As Integer
   Denominator.allocateMemory EqObj.NumberOfUnknowns
   For GroupIndex = 0 To Denominator.NumberOfGroups - 1
      FactorGroupIndexes = EqObj.getLetterIndexes(GroupIndex)
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
      FactorGroupIndexes = EqObj.getLetterIndexes(i)
      For j = 0 To NumberOfFactors - 1
         Sheets(2).Cells(j + 1, ColumnIndex + i) = Factors(j).Degree(FactorGroupIndexes(j))
      Next j
   Next i
End Sub

Public Sub printStringFactors()
   Dim i As Integer
   For i = 0 To NumberOfFactors - 1
      Cells(i + 1, 1) = "L["
      StringFactors(i).printItemOfGroup dgDegree, False, i + 1, 2
      Cells(i + 1, NumberOfDegrees + 2) = "]"
   Next i
End Sub

Public Sub fillRepetitionsOfDenominator()
   Dim GroupIndex As Integer
   Dim UnknownArray() As Integer
   UnknownArray = EqObj.getUnknownArray
   For GroupIndex = 0 To Denominator.NumberOfGroups - 1
      Denominator.Repetition(GroupIndex) = UnknownArray(GroupIndex)
   Next GroupIndex
   Erase UnknownArray
End Sub

Public Sub prepareSheetBefore()
   ActiveWindow.WindowState = xlMaximized
   ActiveWindow.FreezePanes = False
   With Sheets(2).Cells
      .Clear
      .ColumnWidth = 2
      .Interior.Pattern = xlNone
      .Interior.Color = ColorFromHSL(WorksheetFunction.RandBetween(0, 360), 85, 60)
      .Font.ColorIndex = xlAutomatic
      .Font.Bold = False
      .Font.Size = 15
      .Font.Name = "Century Gothic"
      .HorizontalAlignment = xlCenter
      .VerticalAlignment = xlCenter
   End With
End Sub

Public Sub doMultiplication()
   Dim i As Integer
   Dim LastRow As Long
   Dim LastColumn As Integer
   Sheets(2).Select
   Cells(NumberOfFactors + 1, 1).EntireRow.Font.Bold = True
   printStringFactors
   LastRow = NumberOfFactors + 1
   LastColumn = NumberOfDegrees + 3
   LastColumn = Numerator.printItemOfGroup(dgDegree, False, LastRow, LastColumn) + 1
   printPointersOfDenominator LastColumn
   LastColumn = Denominator.printItemOfGroup(dgDegree, False, LastRow, LastColumn)
   Do
      LastColumn = NumberOfDegrees + 2
      EqObj.fillUnknowns
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
   Loop Until (LastRow >= MaxRow Or EqObj.isDone())
End Sub

Public Sub prepareSheetAfter()
   Sheets(2).Select
   ActiveWindow.WindowState = xlMaximized
   Cells(NumberOfFactors + 2, 1).Select
   ActiveWindow.FreezePanes = False
   ActiveWindow.FreezePanes = True
   Sheets(2).Cells.EntireColumn.AutoFit
End Sub

' Destructor

Private Sub Class_Terminate()
   Dim FactorIndex As Integer
   Dim Index As Integer
   For FactorIndex = 0 To NumberOfFactors - 1
      Set StringFactors(FactorIndex) = Nothing
      Set Factors(FactorIndex) = Nothing
   Next FactorIndex
   For Index = 0 To 3
      Set ColumnOperators(Index) = Nothing
   Next Index
   Set Denominator = Nothing
   Set Numerator = Nothing
   Set Result = Nothing
   Set EqObj = Nothing
   Erase Factors
   Erase StringFactors
   Erase ColumnOperators
   Erase Conformity
End Sub
