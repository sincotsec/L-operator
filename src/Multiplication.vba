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

Dim Conformity() As Integer

Dim UpperBounds() As Integer
Dim LowerBounds() As Integer
Dim DiminishingGroupIndex As Integer

Const MaxRow = 1500

' Methods

Public Sub allocateMemory(parNumberOfFactors As Integer, parNumberOfDegrees As Integer)
   Dim i As Integer
   Dim j As Integer
   Dim NewConformity() As Integer
   NumberOfFactors = parNumberOfFactors
   NumberOfDegrees = parNumberOfDegrees
   ReDim StringFactors(NumberOfFactors - 1)
   ReDim Factors(NumberOfFactors + 1)
   Set Numerator = Factors(NumberOfFactors)
   Set Denominator = Factors(NumberOfFactors + 1)
   For i = 0 To NumberOfFactors - 1
      StringFactors(i).allocateMemory NumberOfDegrees
      StringFactors(i).fillStringFactor 3 + i
      Factors(i).groupDegreesFromOperator StringFactors(i), NewConformity
      For j = 0 To StringFactors(i).NumberOfGroups - 1
         Factors(i).Repetition(NewConformity(j)) = Factors(i).Repetition(NewConformity(j)) + 1
      Next j
   Next i
End Sub

Public Function getFactorGroupIndexes(ByVal DenominatorGroupIndex As Integer) As Integer()
   Dim i As Integer
   Dim TempIndex As Integer
   Dim IndexString As String
   Dim FactorGroupsIndexes() As Integer
   ReDim FactorGroupsIndexes(NumberOfFactors)
   TempIndex = DenominatorGroupIndex
   For i = NumberOfFactors - 1 To 1 Step -1
      FactorGroupsIndexes(i) = TempIndex Mod Factors(i).NumberOfGroups
      TempIndex = TempIndex \ Factors(i).NumberOfGroups
   Next i
   FactorGroupsIndexes(0) = TempIndex
   For i = 0 To NumberOfFactors - 1
      IndexString = IndexString & " " & FactorGroupsIndexes(i) & " "
   Next i
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
   Dim NumberOfDenominatorGroups As Integer
   Dim FactorGroupIndexes() As Integer
   
   NumberOfDenominatorGroups = 1
   For FactorIndex = 0 To NumberOfFactors - 1
      NumberOfDenominatorGroups = NumberOfDenominatorGroups * Factors(FactorIndex).NumberOfGroups
   Next FactorIndex
   Denominator.allocateMemory NumberOfDenominatorGroups
   
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

Public Sub setColumns()
   Dim i As Integer
   Dim ColumnIndex As Integer
   Dim HueNumber As Integer
   Dim TitleRow As Integer
   TitleRow = NumberOfFactors + 1
   ColumnIndex = 0
   HueNumber = WorksheetFunction.RandBetween(0, 360)
   For i = 0 To NumberOfFactors + 1
      Factors(i).setColumns ColumnIndex, HueNumber
      Factors(i).prepareTitle TitleRow
      Factors(i).printItemOfGroup dgDegree, TitleRow
      ColumnIndex = Factors(i).LastColumn
      HueNumber = HueNumber + 45
   Next i
   printPointersOfDenominator
End Sub

Function isFirstIndex(FactorIndex As Integer, ByVal DenominatorGroupIndex As Integer) As Boolean
   Dim FactorGroupIndexes() As Integer
   Dim i As Integer
   FactorGroupIndexes = getFactorGroupIndexes(DenominatorGroupIndex)
   isFirstIndex = True
   For i = 0 To NumberOfFactors - 1
      If (i <> FactorIndex) And (FactorGroupIndexes(i) <> 0) Then
         isFirstIndex = False
         Exit For
      End If
   Next i
End Function

Private Function getPreviousIndex(FactorIndex As Integer, ByVal DenominatorGroupIndex As Integer) As Integer
   Dim i As Integer
   Dim PreviousFactorGroupIndexes() As Integer
   ReDim PreviousFactorGroupIndexes(NumberOfFactors)
   PreviousFactorGroupIndexes = getFactorGroupIndexes(DenominatorGroupIndex)
   For i = NumberOfFactors - 1 To 0 Step -1
      If i <> FactorIndex Then
         If PreviousFactorGroupIndexes(i) > 0 Then
            PreviousFactorGroupIndexes(i) = PreviousFactorGroupIndexes(i) - 1
            Exit For
         Else
            PreviousFactorGroupIndexes(i) = Factors(i).NumberOfGroups - 1
         End If
      End If
   Next i
   getPreviousIndex = getDenominatorGroupIndex(PreviousFactorGroupIndexes)
End Function

Public Sub fillRepetitionUpperBounds(DenominatorGroupIndex As Integer)
   Dim FactorIndex As Integer
   Dim i As Integer
   Dim FactorGroupIndexes() As Integer
   Dim PreviousIndex As Integer
   FactorGroupIndexes = getFactorGroupIndexes(DenominatorGroupIndex)
   For FactorIndex = 0 To NumberOfFactors - 1
      If isFirstIndex(FactorIndex, DenominatorGroupIndex) Then
         UpperBounds(FactorIndex, DenominatorGroupIndex) = Factors(FactorIndex).Repetition(FactorGroupIndexes(FactorIndex))
      Else
         PreviousIndex = getPreviousIndex(FactorIndex, DenominatorGroupIndex)
         UpperBounds(FactorIndex, DenominatorGroupIndex) = UpperBounds(FactorIndex, PreviousIndex) - Denominator.Repetition(PreviousIndex)
      End If
   Next FactorIndex
   Erase FactorGroupIndexes
End Sub

Public Sub fillRepetitionsOfDenominator()
   Dim DenominatorGroupIndex As Integer
   If DiminishingGroupIndex <> -1 Then
      Denominator.Repetition(DiminishingGroupIndex) = Denominator.Repetition(DiminishingGroupIndex) - 1
   End If
   For DenominatorGroupIndex = 0 To Denominator.NumberOfGroups - 1
      If DenominatorGroupIndex > DiminishingGroupIndex Then
         Call fillRepetitionUpperBounds(DenominatorGroupIndex)
         Denominator.Repetition(DenominatorGroupIndex) = getRepetition(DenominatorGroupIndex)
         LowerBounds(DenominatorGroupIndex) = getMu(DenominatorGroupIndex)
         If LowerBounds(DenominatorGroupIndex) < 0 Then LowerBounds(DenominatorGroupIndex) = 0
      End If
   Next DenominatorGroupIndex
End Sub

Public Function getRepetition(DenominatorGroupIndex As Integer) As Integer
   Dim i As Integer
   getRepetition = UpperBounds(0, DenominatorGroupIndex)
   For i = 0 To NumberOfFactors - 1
      getRepetition = getMinimum(getRepetition, UpperBounds(i, DenominatorGroupIndex))
   Next i
End Function

Public Function getMu(DenominatorGroupIndex As Integer) As Integer
   Dim i As Integer
   Dim j As Integer
   Dim LettersASum As Integer
   Dim TempIndexes() As Integer
   Dim FactorGroupIndexes() As Integer
   FactorGroupIndexes = getFactorGroupIndexes(DenominatorGroupIndex)
   ReDim TempIndexes(NumberOfFactors)
   getMu = (1 - NumberOfFactors) * NumberOfDegrees
   LettersASum = 0
   For i = 0 To FactorGroupIndexes(0)
      LettersASum = LettersASum + Factors(0).Repetition(i)
   Next i
   getMu = getMu + (NumberOfFactors - 2) * LettersASum + 2 * UpperBounds(0, DenominatorGroupIndex)
   For i = 0 To NumberOfFactors - 1
      TempIndexes(i) = 0
   Next i
   For i = 0 To NumberOfFactors - 1
      For j = 0 To FactorGroupIndexes(i)
         getMu = getMu + UpperBounds(i, getDenominatorGroupIndex(TempIndexes))
         TempIndexes(i) = TempIndexes(i) + 1
      Next j
      TempIndexes(i) = TempIndexes(i) - 1
      getMu = getMu - UpperBounds(0, getDenominatorGroupIndex(TempIndexes))
   Next i
End Function

Public Function getDiminishingNumberIndex() As Integer
   Dim DenominatorGroupIndex As Integer
   getDiminishingNumberIndex = -1
   For DenominatorGroupIndex = Denominator.NumberOfGroups - 1 To 0 Step -1
      If Denominator.Repetition(DenominatorGroupIndex) > LowerBounds(DenominatorGroupIndex) Then
         getDiminishingNumberIndex = DenominatorGroupIndex
         Exit For
      End If
   Next DenominatorGroupIndex
End Function

Public Sub fillRepetitionsOfNumerator()
   Dim i As Integer
   For i = 0 To Numerator.NumberOfGroups - 1
      Numerator.Repetition(i) = 0
   Next i
   For i = 0 To Denominator.NumberOfGroups - 1
      Numerator.Repetition(Conformity(i)) = Numerator.Repetition(Conformity(i)) + Denominator.Repetition(i)
   Next i
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

Public Sub doMultiplication()
   Dim FactorIndex As Integer
   Dim LastRow As Long
   ReDim UpperBounds(NumberOfFactors, Denominator.NumberOfGroups)
   ReDim LowerBounds(Denominator.NumberOfGroups)
   LastRow = NumberOfFactors + 1
   DiminishingGroupIndex = -1
   Do
      Call fillRepetitionsOfDenominator
      DiminishingGroupIndex = getDiminishingNumberIndex()
      Call fillRepetitionsOfNumerator
      LastRow = LastRow + 1
      For FactorIndex = 0 To NumberOfFactors + 1
         Factors(FactorIndex).printItemOfGroup dgRepetition, LastRow
      Next FactorIndex
   Loop Until (LastRow >= MaxRow Or DiminishingGroupIndex = -1)
End Sub

Sub prepareSheetAfter()
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
      Set Factors(FactorIndex) = Nothing
   Next FactorIndex
   Set Factors(NumberOfFactors) = Nothing
   Set Factors(NumberOfFactors + 1) = Nothing
   Set Denominator = Nothing
   Set Numerator = Nothing
   Erase Factors
   Erase StringFactors
   Erase Conformity
   Erase UpperBounds
   Erase LowerBounds
End Sub
