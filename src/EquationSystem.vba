VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EquationSystem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Const MaxRow = 1500

Dim NumberOfLayers As Integer
Dim SumOfLetters As Integer
Dim Letters() As Variant
Dim Degrees() As Variant
Dim FactorDegrees() As Variant
Dim NumberOfSections() As Integer

Dim Unknowns() As Integer
Dim NumberOfUnknowns As Integer
Dim UpperBounds() As Integer
Dim LowerBounds() As Integer
Dim DiminishingUnknownIndex As Integer

Dim NumeratorDegrees() As Integer
Dim NumeratorRepetitions() As Integer
Dim ResultDegrees() As Integer
Dim DenominatorDegrees() As Integer

Dim NumberOfNumeratorDegrees As Integer

Dim ConformityArray() As Integer

' Basic methods

Public Sub fillArrays(NumberOfFactors As Integer, NumberOfDegrees As Integer)
   Dim TempArrayFrom() As Integer
   Dim TempArrayTo() As Integer
   Dim TempArray() As Integer
   Dim LayerIndex As Integer
   Dim DegreeIndex As Integer

   NumberOfLayers = NumberOfFactors
   SumOfLetters = NumberOfDegrees
   ReDim Letters(NumberOfLayers - 1)
   ReDim Degrees(NumberOfLayers - 1)
   ReDim NumberOfSections(NumberOfLayers - 1)
   ReDim FactorDegrees(NumberOfLayers - 1)
   ReDim ConformityArray(SumOfLetters - 1)

   For LayerIndex = 0 To NumberOfLayers - 1
      ReDim TempArray(SumOfLetters - 1)
      FactorDegrees(LayerIndex) = TempArray
      For DegreeIndex = 0 To SumOfLetters - 1
         FactorDegrees(LayerIndex)(DegreeIndex) = Cells(1 + LayerIndex, 5 + DegreeIndex)
      Next DegreeIndex

      TempArrayFrom = FactorDegrees(LayerIndex)
      groupArrays TempArrayFrom, SumOfLetters, TempArrayTo, NumberOfSections(LayerIndex)
      Degrees(LayerIndex) = TempArrayTo

      Letters(LayerIndex) = TempArray
      For DegreeIndex = 0 To SumOfLetters - 1
         Letters(LayerIndex)(ConformityArray(DegreeIndex)) = Letters(LayerIndex)(ConformityArray(DegreeIndex)) + 1
      Next DegreeIndex
   Next LayerIndex

   Erase TempArray
   Erase TempArrayFrom
   Erase TempArrayTo
End Sub

Public Sub prepareSolution()
   Dim i As Integer
   NumberOfUnknowns = 1
   For i = 0 To NumberOfLayers - 1
      NumberOfUnknowns = NumberOfUnknowns * NumberOfSections(i)
   Next i
   ReDim Unknowns(NumberOfUnknowns - 1)
   ReDim UpperBounds(NumberOfLayers - 1, NumberOfUnknowns - 1)
   ReDim LowerBounds(NumberOfUnknowns - 1)
   DiminishingUnknownIndex = -1
End Sub

Public Sub fillDegreesOfDenominator()
   Dim FactorIndex As Integer
   Dim GroupIndex As Integer
   Dim FactorGroupIndexes() As Integer
   ReDim DenominatorDegrees(NumberOfUnknowns - 1)
   For GroupIndex = 0 To NumberOfUnknowns - 1
      FactorGroupIndexes = getLetterIndexes(GroupIndex)
      For FactorIndex = 0 To NumberOfLayers - 1
         DenominatorDegrees(GroupIndex) = DenominatorDegrees(GroupIndex) + Degrees(FactorIndex)(FactorGroupIndexes(FactorIndex))
      Next FactorIndex
   Next GroupIndex
   Erase FactorGroupIndexes
   groupArrays DenominatorDegrees, NumberOfUnknowns, NumeratorDegrees, NumberOfNumeratorDegrees
End Sub

Public Sub printHeaders()
   Dim LastRow As Long
   Dim LastColumn As Integer
   Sheets(1).Select
   prepareSheetBefore
   Cells(NumberOfLayers + 1, 1).EntireRow.Font.Bold = True
   Range("A1") = "Number of factors"
   Range("B1") = NumberOfLayers
   Range("A2") = "Number of degrees"
   Range("B2") = SumOfLetters
   printFactorDegrees
   LastRow = NumberOfLayers + 1
   LastColumn = SumOfLetters + 7
   printArray NumeratorDegrees, NumberOfNumeratorDegrees, False, LastRow, LastColumn
   LastColumn = LastColumn + NumberOfNumeratorDegrees + 1
   printPointersOfDenominator LastColumn
   printArray DenominatorDegrees, NumberOfUnknowns, False, LastRow, LastColumn
End Sub

Public Sub doMultiplication()
   Dim LastRow As Long
   Dim LastColumn As Integer
   Sheets(1).Select
   LastRow = NumberOfLayers + 1
   
   Do
      fillUnknowns
      groupRepetitionsFromDenominator
      fillDegreesOfResult
      
      LastRow = LastRow + 1
      LastColumn = SumOfLetters + 6
      Cells(LastRow, LastColumn) = "("
      printArray NumeratorRepetitions, NumberOfNumeratorDegrees, True, LastRow, LastColumn + 1
      LastColumn = LastColumn + NumberOfNumeratorDegrees + 1
      Cells(LastRow, LastColumn) = ") : ("
      printArray Unknowns, NumberOfUnknowns, True, LastRow, LastColumn + 1
      LastColumn = LastColumn + NumberOfUnknowns + 1
      Cells(LastRow, LastColumn) = ") L["
      printArray ResultDegrees, SumOfLetters, False, LastRow, LastColumn + 1
      LastColumn = LastColumn + SumOfLetters + 1
      Cells(LastRow, LastColumn) = "]"
   Loop Until (LastRow >= MaxRow Or DiminishingUnknownIndex = -1)
   
   ActiveWindow.WindowState = xlMaximized
   ActiveWindow.ScrollColumn = 1
   Cells(NumberOfLayers + 2, 1).Select
   ActiveWindow.FreezePanes = False
   ActiveWindow.FreezePanes = True
   Cells(NumberOfLayers + 1, 1).EntireRow.Borders(xlEdgeBottom).LineStyle = xlContinuous
   Sheets(1).Cells.EntireColumn.AutoFit
End Sub

' Intermediate methods

Private Sub groupArrays(ArrayFrom() As Integer, CountFrom As Integer, ArrayTo() As Integer, CountTo As Integer)
   Dim isFound As Boolean
   Dim IndexFrom As Integer
   Dim IndexTo As Integer
   ReDim ConformityArray(CountFrom - 1)
   ReDim ArrayTo(CountFrom - 1)
   CountTo = 0
   
   For IndexFrom = 0 To CountFrom - 1
      isFound = False
      For IndexTo = 0 To CountTo - 1
         If ArrayTo(IndexTo) = ArrayFrom(IndexFrom) Then
            ConformityArray(IndexFrom) = IndexTo
            isFound = True
            Exit For
         End If
      Next IndexTo
      If (Not isFound) Then
         CountTo = CountTo + 1
         ArrayTo(CountTo - 1) = ArrayFrom(IndexFrom)
         ConformityArray(IndexFrom) = CountTo - 1
      End If
   Next IndexFrom
End Sub

Private Function getLetterIndexes(ByVal UnknownIndex As Integer) As Integer()
   Dim i As Integer
   Dim TempIndex As Integer
   Dim LetterIndexes() As Integer
   ReDim LetterIndexes(NumberOfLayers - 1)
   TempIndex = UnknownIndex
   For i = NumberOfLayers - 1 To 1 Step -1
      LetterIndexes(i) = TempIndex Mod NumberOfSections(i)
      TempIndex = TempIndex \ NumberOfSections(i)
   Next i
   LetterIndexes(0) = TempIndex
   getLetterIndexes = LetterIndexes()
   Erase LetterIndexes
End Function

Private Function getUnknownIndex(ByRef LetterIndexes() As Integer) As Integer
   Dim i As Integer
   getUnknownIndex = LetterIndexes(0)
   For i = 1 To NumberOfLayers - 1
      getUnknownIndex = getUnknownIndex * NumberOfSections(i)
      getUnknownIndex = getUnknownIndex + LetterIndexes(i)
   Next i
End Function

Private Function isFirstIndex(ByVal LayerIndex As Integer, ByVal UnknownIndex As Integer) As Boolean
   Dim LetterIndexes() As Integer
   Dim i As Integer
   LetterIndexes = getLetterIndexes(UnknownIndex)
   isFirstIndex = True
   For i = 0 To NumberOfLayers - 1
      If (i <> LayerIndex) And (LetterIndexes(i) <> 0) Then
         isFirstIndex = False
         Exit For
      End If
   Next i
End Function

Private Function getPreviousIndex(LayerIndex As Integer, ByVal UnknownIndex As Integer) As Integer
   Dim i As Integer
   Dim PreviousLetterIndexes() As Integer
   ReDim PreviousLetterIndexes(NumberOfLayers - 1)
   PreviousLetterIndexes = getLetterIndexes(UnknownIndex)
   For i = NumberOfLayers - 1 To 0 Step -1
      If (i <> LayerIndex) And (PreviousLetterIndexes(i) > 0) Then
         PreviousLetterIndexes(i) = PreviousLetterIndexes(i) - 1
         Exit For
      ElseIf (i <> LayerIndex) And (PreviousLetterIndexes(i) = 0) Then
         PreviousLetterIndexes(i) = NumberOfSections(i) - 1
      End If
   Next i
   getPreviousIndex = getUnknownIndex(PreviousLetterIndexes)
   Erase PreviousLetterIndexes
End Function

Private Function getUnknown(UnknownIndex As Integer) As Integer
   Dim LayerIndex As Integer
   Dim PreviousIndex As Integer
   Dim LetterIndexes() As Integer
   LetterIndexes = getLetterIndexes(UnknownIndex)
   For LayerIndex = 0 To NumberOfLayers - 1
      If isFirstIndex(LayerIndex, UnknownIndex) Then
         UpperBounds(LayerIndex, UnknownIndex) = Letters(LayerIndex)(LetterIndexes(LayerIndex))
      Else
         PreviousIndex = getPreviousIndex(LayerIndex, UnknownIndex)
         UpperBounds(LayerIndex, UnknownIndex) = UpperBounds(LayerIndex, PreviousIndex) - Unknowns(PreviousIndex)
      End If
   Next LayerIndex
   getUnknown = UpperBounds(0, UnknownIndex)
   For LayerIndex = 0 To NumberOfLayers - 1
      getUnknown = WorksheetFunction.Min(getUnknown, UpperBounds(LayerIndex, UnknownIndex))
   Next LayerIndex
   Erase LetterIndexes
End Function

Private Function getMu(UnknownIndex As Integer) As Integer
   Dim LayerIndex As Integer
   Dim SectionIndex As Integer
   Dim LettersASum As Integer
   Dim TempIndexes() As Integer
   Dim LetterIndexes() As Integer
   LetterIndexes = getLetterIndexes(UnknownIndex)
   ReDim TempIndexes(NumberOfLayers - 1)
   getMu = (1 - NumberOfLayers) * SumOfLetters
   LettersASum = 0
   For SectionIndex = 0 To LetterIndexes(0)
      LettersASum = LettersASum + Letters(0)(SectionIndex)
   Next SectionIndex
   getMu = getMu + (NumberOfLayers - 2) * LettersASum + 2 * UpperBounds(0, UnknownIndex)
   For LayerIndex = 0 To NumberOfLayers - 1
      TempIndexes(LayerIndex) = 0
   Next LayerIndex
   For LayerIndex = 0 To NumberOfLayers - 1
      For SectionIndex = 0 To LetterIndexes(LayerIndex)
         getMu = getMu + UpperBounds(LayerIndex, getUnknownIndex(TempIndexes))
         TempIndexes(LayerIndex) = TempIndexes(LayerIndex) + 1
      Next SectionIndex
      TempIndexes(LayerIndex) = TempIndexes(LayerIndex) - 1
      getMu = getMu - UpperBounds(0, getUnknownIndex(TempIndexes))
   Next LayerIndex
   Erase LetterIndexes
   Erase TempIndexes
End Function

Private Sub fillDiminishingUnknownIndex()
   Dim UnknownIndex As Integer
   DiminishingUnknownIndex = -1
   For UnknownIndex = NumberOfUnknowns - 1 To 0 Step -1
      If Unknowns(UnknownIndex) > LowerBounds(UnknownIndex) Then
         DiminishingUnknownIndex = UnknownIndex
         Exit For
      End If
   Next UnknownIndex
End Sub

Private Sub fillUnknowns()
   Dim UnknownIndex As Integer
   If DiminishingUnknownIndex <> -1 Then
      Unknowns(DiminishingUnknownIndex) = Unknowns(DiminishingUnknownIndex) - 1
   End If
   For UnknownIndex = 0 To NumberOfUnknowns - 1
      If UnknownIndex > DiminishingUnknownIndex Then
         Unknowns(UnknownIndex) = getUnknown(UnknownIndex)
         LowerBounds(UnknownIndex) = getMu(UnknownIndex)
         If LowerBounds(UnknownIndex) < 0 Then LowerBounds(UnknownIndex) = 0
      End If
   Next UnknownIndex
   fillDiminishingUnknownIndex
End Sub

Private Sub printArray(dgArray() As Integer, dgSize As Integer, ByVal dgFactorial As Boolean, ByVal RowIndex As Integer, ByVal ColumnIndex As Integer)
   Dim i As Integer
   Dim Factorial As String
   Factorial = ""
   If dgFactorial Then Factorial = "!"
   For i = 0 To dgSize - 1
      Sheets(1).Cells(RowIndex, ColumnIndex + i) = dgArray(i) & Factorial
   Next i
End Sub

Private Sub printPointersOfDenominator(ByVal ColumnIndex As Integer)
   Dim i As Integer
   Dim j As Integer
   Dim FactorGroupIndexes() As Integer
   For i = 0 To NumberOfUnknowns - 1
      FactorGroupIndexes = getLetterIndexes(i)
      For j = 0 To NumberOfLayers - 1
         Sheets(1).Cells(j + 1, ColumnIndex + i) = Degrees(j)(FactorGroupIndexes(j))
      Next j
   Next i
End Sub

Private Sub printFactorDegrees()
   Dim i As Integer
   Dim j As Integer
   For i = 0 To NumberOfLayers - 1
      Cells(i + 1, 4) = "L["
      For j = 0 To SumOfLetters - 1
         Sheets(1).Cells(i + 1, 5 + j) = FactorDegrees(i)(j)
      Next j
      Cells(i + 1, SumOfLetters + 5) = "]"
   Next i
End Sub

Private Sub groupRepetitionsFromDenominator()
   Dim GroupIndex As Integer
   ReDim NumeratorRepetitions(NumberOfNumeratorDegrees - 1)
   For GroupIndex = 0 To NumberOfNumeratorDegrees - 1
      NumeratorRepetitions(GroupIndex) = 0
   Next GroupIndex
   For GroupIndex = 0 To NumberOfUnknowns - 1
      NumeratorRepetitions(ConformityArray(GroupIndex)) = NumeratorRepetitions(ConformityArray(GroupIndex)) + Unknowns(GroupIndex)
   Next GroupIndex
End Sub

Private Sub fillDegreesOfResult()
   Dim GroupIndex As Integer
   Dim j As Integer
   Dim k As Integer
   ReDim ResultDegrees(SumOfLetters - 1)
   k = 0
   For GroupIndex = 0 To NumberOfNumeratorDegrees - 1
      For j = 1 To NumeratorRepetitions(GroupIndex)
         ResultDegrees(k) = NumeratorDegrees(GroupIndex)
         k = k + 1
      Next j
   Next GroupIndex
   'Debug.Print getInfo()
End Sub

' String functions

Public Function getInfo() As String
   Dim DebugString As String
   Dim i As Integer, j As Integer
   DebugString = "NumberOfLayers = " & NumberOfLayers & vbLf & "SumOfLetters = " & SumOfLetters
   For i = 0 To NumberOfLayers - 1
      DebugString = DebugString & vbLf & i & ":"
      For j = 0 To NumberOfSections(i) - 1
         DebugString = DebugString & " " & Letters(i)(j)
      Next j
   Next i
   DebugString = DebugString & vbLf & "NumberOfUnknowns = " & NumberOfUnknowns
   getInfo = DebugString
End Function

Public Function getUnknownInfo() As String
   Dim i As Integer
   Dim DebugString As String
   For i = 0 To NumberOfUnknowns - 1
      DebugString = DebugString & " " & Unknowns(i)
   Next i
   DebugString = DebugString & " | " & DiminishingUnknownIndex
   getUnknownInfo = DebugString
End Function

Public Function getLetterInfo() As String
   Dim DebugString As String
   Dim i As Integer
   Dim j As Integer
   DebugString = "Array of letters" & vbLf
   For i = 0 To NumberOfLayers - 1
      DebugString = DebugString & "Factor " & i + 1 & ". Groups: " & NumberOfSections(i) & ". Repetitions:"
      For j = 0 To NumberOfSections(i) - 1
         DebugString = DebugString & " " & Letters(i)(j)
      Next j
      DebugString = DebugString & vbLf
   Next i
   getLetterInfo = DebugString
End Function

' Destructor

Private Sub Class_Terminate()
   Erase UpperBounds
   Erase LowerBounds
   Erase Letters
   Erase Degrees
   Erase NumberOfSections
   Erase Unknowns
   Erase NumeratorDegrees
   Erase NumeratorRepetitions
   Erase ResultDegrees
   Erase DenominatorDegrees
   Erase ConformityArray
   Erase FactorDegrees
End Sub
