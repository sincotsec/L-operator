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
   ReDim Factors(NumberOfFactors + 2)
   Set Numerator = Factors(NumberOfFactors)
   Set Denominator = Factors(NumberOfFactors + 1)
   Set Result = Factors(NumberOfFactors + 2)
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
'   For i = 0 To NumberOfFactors + 2
'      Factors(i).setColumns ColumnIndex, HueNumber
'      Factors(i).prepareTitle TitleRow
'      ColumnIndex = Factors(i).LastColumn
'      HueNumber = HueNumber + 30
'      If HueNumber > 360 Then HueNumber = HueNumber - 360
'   Next i
   StringFactors(0).setColumns ColumnIndex, HueNumber
   StringFactors(0).prepareTitle TitleRow
   ColumnIndex = StringFactors(0).LastColumn
   HueNumber = HueNumber + 30
   If HueNumber > 360 Then HueNumber = HueNumber - 360
   
   Numerator.setColumns ColumnIndex, HueNumber
   Numerator.prepareTitle TitleRow
   ColumnIndex = Numerator.LastColumn
   HueNumber = HueNumber + 30
   If HueNumber > 360 Then HueNumber = HueNumber - 360
   
   Denominator.setColumns ColumnIndex, HueNumber
   Denominator.prepareTitle TitleRow
   ColumnIndex = Denominator.LastColumn
   HueNumber = HueNumber + 30
   If HueNumber > 360 Then HueNumber = HueNumber - 360
   
   Result.setColumns ColumnIndex, HueNumber
   Result.prepareTitle TitleRow
   ColumnIndex = Result.LastColumn
   HueNumber = HueNumber + 30
   If HueNumber > 360 Then HueNumber = HueNumber - 360

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
   Dim FactorIndex As Integer
   Dim LastRow As Long
   LastRow = NumberOfFactors + 1
   Do
      EqObj.fillUnknowns
      Call fillRepetitionsOfDenominator
      Numerator.groupRepetitionsFromOperator Denominator, Conformity
      Result.degroupDegreesFromOperator Numerator
      LastRow = LastRow + 1
'      For FactorIndex = 0 To NumberOfFactors + 1
'         Factors(FactorIndex).printItemOfGroup dgRepetition, LastRow
'      Next FactorIndex
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
   For FactorIndex = 0 To NumberOfFactors - 1
      Set StringFactors(FactorIndex) = Nothing
   Next FactorIndex
   For FactorIndex = 0 To NumberOfFactors + 2
      Set Factors(FactorIndex) = Nothing
   Next FactorIndex
   Set Denominator = Nothing
   Set Numerator = Nothing
   Set Result = Nothing
   Set EqObj = Nothing
   Erase Factors
   Erase StringFactors
   Erase Conformity
End Sub
