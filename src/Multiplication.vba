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

Public NumberOfFactors As Integer
Public NumberOfDegrees As Integer

Dim Factors() As New Operator
Dim StringFactors() As New Operator
Dim Denominator As Operator
Dim Numerator As Operator

Dim Conformity() As Integer
Dim EqObj As Equation

Const MaxRow = 1500

' Methods

Public Sub allocateMemory(parNumberOfFactors As Integer, parNumberOfDegrees As Integer)
   NumberOfFactors = parNumberOfFactors
   NumberOfDegrees = parNumberOfDegrees
   ReDim StringFactors(NumberOfFactors - 1)
   ReDim Factors(NumberOfFactors + 1)
   Set Numerator = Factors(NumberOfFactors)
   Set Denominator = Factors(NumberOfFactors + 1)
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

Public Function getFactorGroupIndexes(ByVal DenominatorGroupIndex As Integer) As Integer()
   Dim i As Integer
   Dim TempIndex As Integer
   Dim FactorGroupsIndexes() As Integer
   ReDim FactorGroupsIndexes(NumberOfFactors)
   TempIndex = DenominatorGroupIndex
   For i = NumberOfFactors - 1 To 1 Step -1
      FactorGroupsIndexes(i) = TempIndex Mod Factors(i).NumberOfGroups
      TempIndex = TempIndex \ Factors(i).NumberOfGroups
   Next i
   FactorGroupsIndexes(0) = TempIndex
   getFactorGroupIndexes = FactorGroupsIndexes()
   Erase FactorGroupsIndexes
End Function

Public Function getDenominatorGroupIndex(ByRef FactorGroupIndexes() As Integer) As Integer
   Dim i As Integer
   getDenominatorGroupIndex = FactorGroupIndexes(0)
   For i = 1 To NumberOfFactors - 1
      getDenominatorGroupIndex = getDenominatorGroupIndex * Factors(i).NumberOfGroups
      getDenominatorGroupIndex = getDenominatorGroupIndex + FactorGroupIndexes(i)
   Next i
End Function

Public Sub fillDegreesOfDenominator()
   Dim FactorIndex As Integer
   Dim GroupIndex As Integer
   Dim FactorGroupIndexes() As Integer
   
   Denominator.allocateMemory EqObj.NumberOfUnknowns
   For GroupIndex = 0 To Denominator.NumberOfGroups - 1
      FactorGroupIndexes = getFactorGroupIndexes(GroupIndex)
      For FactorIndex = 0 To NumberOfFactors - 1
         Denominator.Degree(GroupIndex) = Denominator.Degree(GroupIndex) + Factors(FactorIndex).Degree(FactorGroupIndexes(FactorIndex))
         Denominator.Repetition(GroupIndex) = 1
      Next FactorIndex
   Next GroupIndex
   Erase FactorGroupIndexes
   Numerator.groupDegreesFromOperator Denominator, Conformity
End Sub

Public Sub printPointersOfDenominator()
   Dim i As Integer
   Dim j As Integer
   Dim FactorGroupIndexes() As Integer
   For i = 0 To Denominator.NumberOfGroups - 1
      FactorGroupIndexes = getFactorGroupIndexes(i)
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
   TitleRow = 1
   ColumnIndex = 0
   HueNumber = WorksheetFunction.RandBetween(0, 360)
   For i = 0 To NumberOfFactors + 1
      Factors(i).setColumns ColumnIndex, HueNumber
      Factors(i).prepareTitle TitleRow
      Factors(i).printItemOfGroup dgDegree, TitleRow
      ColumnIndex = Factors(i).LastColumn
      HueNumber = HueNumber + 45
      If HueNumber > 360 Then HueNumber = HueNumber - 360
   Next i
End Sub

Public Sub fillRepetitionsOfDenominator()
   Dim GroupIndex As Integer
   Dim UnknownArray() As Integer
   UnknownArray = EqObj.getUnknownArray
   For GroupIndex = 0 To Denominator.NumberOfGroups - 1
      Denominator.Repetition(GroupIndex) = UnknownArray(GroupIndex)
   Next GroupIndex
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
      .Font.Name = "Arial Narrow"
      .HorizontalAlignment = xlCenter
      .VerticalAlignment = xlCenter
   End With
End Sub

Public Sub doMultiplication()
   Dim FactorIndex As Integer
   Dim LastRow As Long
   LastRow = 1
   Do
      EqObj.fillUnknowns
      Call fillRepetitionsOfDenominator
      Numerator.groupRepetitionsFromOperator Denominator, Conformity
      LastRow = LastRow + 1
      For FactorIndex = 0 To NumberOfFactors + 1
         Factors(FactorIndex).printItemOfGroup dgRepetition, LastRow
      Next FactorIndex
   Loop Until (LastRow >= MaxRow Or EqObj.isDone())
End Sub

Public Sub prepareSheetAfter()
   Sheets(2).Select
   ActiveWindow.WindowState = xlMaximized
   Cells(2, 1).Select
   ActiveWindow.FreezePanes = False
   ActiveWindow.FreezePanes = True
   Sheets(2).Cells.EntireColumn.AutoFit
End Sub

' Destructor

Private Sub Class_Terminate()
   Dim FactorIndex As Integer
   For FactorIndex = 0 To NumberOfFactors - 1
      Set StringFactors(FactorIndex) = Nothing
      Set Factors(FactorIndex) = Nothing
   Next FactorIndex
   Set Factors(NumberOfFactors) = Nothing
   Set Factors(NumberOfFactors + 1) = Nothing
   Set Denominator = Nothing
   Set Numerator = Nothing
   Set EqObj = Nothing
   Erase Factors
   Erase StringFactors
   Erase Conformity
End Sub
