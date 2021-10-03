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

Public Sub printPointersOfDenominator()
   Dim i As Integer
   Dim j As Integer
   Dim FactorGroupIndexes() As Integer
   For i = 0 To Denominator.NumberOfGroups - 1
      FactorGroupIndexes = EqObj.getLetterIndexes(i)
      For j = 0 To NumberOfFactors - 1
         Sheets(2).Cells(j + 1, Denominator.FirstColumn + i) = Factors(j).Degree(FactorGroupIndexes(j))
      Next j
   Next i
End Sub

Public Sub setColumns()
   Dim i As Integer
   Dim ColumnIndex As Integer
   Dim HueNumber As Integer
   Dim TitleRow As Integer
   TitleRow = NumberOfFactors + 1
   ColumnIndex = 0
   HueNumber = WorksheetFunction.RandBetween(0, 360)
   For i = 0 To 3
      ColumnOperators(i).setColumns ColumnIndex, HueNumber
      ColumnOperators(i).prepareTitle TitleRow
      ColumnIndex = ColumnOperators(i).LastColumn
      HueNumber = HueNumber + 30
      If HueNumber > 360 Then HueNumber = HueNumber - 360
   Next i
   Numerator.printItemOfGroup dgDegree, TitleRow
   Denominator.printItemOfGroup dgDegree, TitleRow
   ColumnIndex = 0
   For i = 0 To NumberOfFactors - 1
      StringFactors(i).setColumns ColumnIndex, 0
      StringFactors(i).printItemOfGroup dgDegree, i + 1
   Next i
   printPointersOfDenominator
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
      .Interior.Color = ColorFromHSL(WorksheetFunction.RandBetween(0, 360), 70, 40)
      .Font.ColorIndex = xlAutomatic
      .Font.Bold = False
      .Font.Size = 15
      .Font.Name = "Century Gothic"
      .HorizontalAlignment = xlCenter
      .VerticalAlignment = xlCenter
   End With
End Sub

Public Sub doMultiplication()
   Dim LastRow As Long
   LastRow = NumberOfFactors + 1
   Do
      EqObj.fillUnknowns
      Call fillRepetitionsOfDenominator
      Numerator.groupRepetitionsFromOperator Denominator, Conformity
      Result.degroupDegreesFromOperator Numerator
      LastRow = LastRow + 1
      Numerator.printItemOfGroup dgRepetition, LastRow
      Denominator.printItemOfGroup dgRepetition, LastRow
      Result.printItemOfGroup dgDegree, LastRow
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
