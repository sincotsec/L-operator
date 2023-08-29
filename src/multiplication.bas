Attribute VB_Name = "multiplication"
Option Explicit

Const maxRow = 1500
Dim lastRow As Long

Dim numberOfLayers As Integer
Dim sumOfLetters As Integer
Dim letters() As Variant
Dim degrees() As Variant
Dim factorConformity() As Variant
Dim factorDegrees() As Variant
Dim numberOfSections() As Integer

Dim unknowns() As Integer
Dim numberOfUnknowns As Integer
Dim upperBounds() As Integer
Dim lowerBounds() As Integer
Dim diminishingUnknownIndex As Integer

Dim numeratorDegrees() As Integer
Dim numeratorRepetitions() As Integer
Dim resultDegrees() As Integer
Dim denominatorDegrees() As Integer

Dim numberOfNumeratorDegrees As Integer
Dim numeratorConformity() As Integer

Public Sub multiplyOperators()
   fillArrays Cells(1, 2), Cells(2, 2)
   fillDegrees
   fillLetters
   fillDenominatorDegrees
   printHeaders
   doMultiplication
   prepareSheetAfter
   eraseArrays
End Sub

Private Sub fillArrays(numberOfFactors As Integer, numberOfDegrees As Integer)
   Dim TempArray() As Integer
   Dim LayerIndex As Integer
   Dim DegreeIndex As Integer
   
   numberOfLayers = numberOfFactors
   sumOfLetters = numberOfDegrees
   ReDim factorDegrees(numberOfLayers - 1)
   ReDim factorConformity(numberOfLayers - 1)

   For LayerIndex = 0 To numberOfLayers - 1
      ReDim TempArray(sumOfLetters - 1)
      factorDegrees(LayerIndex) = TempArray
      For DegreeIndex = 0 To sumOfLetters - 1
         factorDegrees(LayerIndex)(DegreeIndex) = Cells(1 + LayerIndex, 5 + DegreeIndex)
      Next DegreeIndex
   Next LayerIndex

   Erase TempArray
End Sub

Private Sub fillDegrees()
   Dim TempArrayFrom() As Integer
   Dim TempArrayTo() As Integer
   Dim LayerIndex As Integer
   Dim DegreeIndex As Integer
   Dim TempConformity() As Integer
   ReDim TempConformity(sumOfLetters - 1)
   ReDim degrees(numberOfLayers - 1)
   ReDim numberOfSections(numberOfLayers - 1)
   
   For LayerIndex = 0 To numberOfLayers - 1
      TempArrayFrom = factorDegrees(LayerIndex)
      groupArrays TempArrayFrom, sumOfLetters, TempArrayTo, numberOfSections(LayerIndex), TempConformity
      degrees(LayerIndex) = TempArrayTo
      factorConformity(LayerIndex) = TempConformity
   Next LayerIndex
   
   Erase TempArrayFrom
   Erase TempArrayTo
   Erase TempConformity
End Sub

Private Sub fillLetters()
   Dim LayerIndex As Integer
   Dim DegreeIndex As Integer
   Dim TempArray() As Integer
   ReDim letters(numberOfLayers - 1)
   For LayerIndex = 0 To numberOfLayers - 1
      ReDim TempArray(sumOfLetters - 1)
      letters(LayerIndex) = TempArray
      
      For DegreeIndex = 0 To numberOfSections(LayerIndex) - 1
         letters(LayerIndex)(DegreeIndex) = 0
      Next DegreeIndex
      
      For DegreeIndex = 0 To sumOfLetters - 1
         letters(LayerIndex)(factorConformity(LayerIndex)(DegreeIndex)) _
            = letters(LayerIndex)(factorConformity(LayerIndex)(DegreeIndex)) _
            + 1
      Next DegreeIndex
   Next LayerIndex
   Erase TempArray
End Sub

Private Sub fillDenominatorDegrees()
   Dim factorIndex As Integer
   Dim GroupIndex As Integer
   Dim FactorGroupIndexes() As Integer
   
   numberOfUnknowns = 1
   For factorIndex = 0 To numberOfLayers - 1
      numberOfUnknowns = numberOfUnknowns * numberOfSections(factorIndex)
   Next factorIndex
   
   ReDim denominatorDegrees(numberOfUnknowns - 1)
   ReDim numeratorConformity(numberOfUnknowns - 1)
   For GroupIndex = 0 To numberOfUnknowns - 1
      FactorGroupIndexes = getLetterIndexes(GroupIndex)
      denominatorDegrees(GroupIndex) = 0
      For factorIndex = 0 To numberOfLayers - 1
         denominatorDegrees(GroupIndex) _
            = denominatorDegrees(GroupIndex) _
            + degrees(factorIndex)(FactorGroupIndexes(factorIndex))
      Next factorIndex
   Next GroupIndex
   Erase FactorGroupIndexes
   groupArrays denominatorDegrees, numberOfUnknowns, numeratorDegrees, numberOfNumeratorDegrees, numeratorConformity
End Sub

Private Sub printHeaders()
   Dim LastColumn As Integer
   Dim LayerIndex As Integer
   Dim TempArray() As Integer
   prepareSheetBefore
   Cells(numberOfLayers + 1, 1).EntireRow.Font.Bold = True
   Range("A1") = "Number of factors"
   Range("B1") = numberOfLayers
   Range("A2") = "Number of degrees"
   Range("B2") = sumOfLetters
   For LayerIndex = 0 To numberOfLayers - 1
      TempArray = factorDegrees(LayerIndex)
      Cells(LayerIndex + 1, 4) = "L["
      printArray TempArray, sumOfLetters, False, LayerIndex + 1, 5
      Cells(LayerIndex + 1, sumOfLetters + 5) = "]"
   Next LayerIndex
   lastRow = numberOfLayers + 1
   LastColumn = sumOfLetters + 7
   printArray numeratorDegrees, numberOfNumeratorDegrees, False, lastRow, LastColumn
   LastColumn = LastColumn + numberOfNumeratorDegrees + 1
   printPointersOfDenominator LastColumn
   printArray denominatorDegrees, numberOfUnknowns, False, lastRow, LastColumn
   Erase TempArray
End Sub

Private Sub doMultiplication()
   ReDim unknowns(numberOfUnknowns - 1)
   ReDim upperBounds(numberOfLayers - 1, numberOfUnknowns - 1)
   ReDim lowerBounds(numberOfUnknowns - 1)
   diminishingUnknownIndex = -1
   Do
      fillUnknowns
      fillDiminishingUnknownIndex
      groupRepetitionsFromDenominator
      fillDegreesOfResult
      printTerm
   Loop Until (lastRow >= maxRow Or diminishingUnknownIndex = -1)
End Sub

Private Sub prepareSheetAfter()
   ActiveWindow.WindowState = xlMaximized
   ActiveWindow.ScrollColumn = 1
   Cells(numberOfLayers + 2, 1).Select
   ActiveWindow.FreezePanes = False
   ActiveWindow.FreezePanes = True
   Cells(numberOfLayers + 1, 1).EntireRow.Borders(xlEdgeBottom).LineStyle = xlContinuous
   Cells.EntireColumn.AutoFit
End Sub

' Intermediate methods

Private Sub printTerm()
   Dim LastColumn As Integer
   lastRow = lastRow + 1
   LastColumn = sumOfLetters + 6
   Cells(lastRow, LastColumn) = "("
   printArray numeratorRepetitions, numberOfNumeratorDegrees, True, lastRow, LastColumn + 1
   LastColumn = LastColumn + numberOfNumeratorDegrees + 1
   Cells(lastRow, LastColumn) = ") : ("
   printArray unknowns, numberOfUnknowns, True, lastRow, LastColumn + 1
   LastColumn = LastColumn + numberOfUnknowns + 1
   Cells(lastRow, LastColumn) = ") L["
   printArray resultDegrees, sumOfLetters, False, lastRow, LastColumn + 1
   LastColumn = LastColumn + sumOfLetters + 1
   Cells(lastRow, LastColumn) = "]"
End Sub

Private Sub groupArrays(ArrayFrom() As Integer, CountFrom As Integer, ArrayTo() As Integer, CountTo As Integer, ConformityArray() As Integer)
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

Private Function getLetterIndexes(UnknownIndex As Integer) As Integer()
   Dim i As Integer
   Dim TempIndex As Integer
   Dim LetterIndexes() As Integer
   ReDim LetterIndexes(numberOfLayers - 1)
   
   TempIndex = UnknownIndex
   For i = numberOfLayers - 1 To 1 Step -1
      LetterIndexes(i) = TempIndex Mod numberOfSections(i)
      TempIndex = TempIndex \ numberOfSections(i)
   Next i
   LetterIndexes(0) = TempIndex
   getLetterIndexes = LetterIndexes()
   Erase LetterIndexes
End Function

Private Function getUnknownIndex(LetterIndexes() As Integer) As Integer
   Dim i As Integer
   getUnknownIndex = LetterIndexes(0)
   For i = 1 To numberOfLayers - 1
      getUnknownIndex = getUnknownIndex * numberOfSections(i)
      getUnknownIndex = getUnknownIndex + LetterIndexes(i)
   Next i
End Function

Private Function isFirstIndex(LayerIndex As Integer, UnknownIndex As Integer) As Boolean
   Dim LetterIndexes() As Integer
   Dim i As Integer
   LetterIndexes = getLetterIndexes(UnknownIndex)
   isFirstIndex = True
   For i = 0 To numberOfLayers - 1
      If (i <> LayerIndex) And (LetterIndexes(i) <> 0) Then
         isFirstIndex = False
         Exit For
      End If
   Next i
End Function

Private Function getPreviousIndex(LayerIndex As Integer, UnknownIndex As Integer) As Integer
   Dim i As Integer
   Dim PreviousLetterIndexes() As Integer
   ReDim PreviousLetterIndexes(numberOfLayers - 1)
   PreviousLetterIndexes = getLetterIndexes(UnknownIndex)
   For i = numberOfLayers - 1 To 0 Step -1
      If (i <> LayerIndex) And (PreviousLetterIndexes(i) > 0) Then
         PreviousLetterIndexes(i) = PreviousLetterIndexes(i) - 1
         Exit For
      ElseIf (i <> LayerIndex) And (PreviousLetterIndexes(i) = 0) Then
         PreviousLetterIndexes(i) = numberOfSections(i) - 1
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
   
   For LayerIndex = 0 To numberOfLayers - 1
      If isFirstIndex(LayerIndex, UnknownIndex) Then
         upperBounds(LayerIndex, UnknownIndex) = letters(LayerIndex)(LetterIndexes(LayerIndex))
      Else
         PreviousIndex = getPreviousIndex(LayerIndex, UnknownIndex)
         upperBounds(LayerIndex, UnknownIndex) = upperBounds(LayerIndex, PreviousIndex) - unknowns(PreviousIndex)
      End If
   Next LayerIndex
   getUnknown = upperBounds(0, UnknownIndex)
   
   For LayerIndex = 0 To numberOfLayers - 1
      getUnknown = WorksheetFunction.Min(getUnknown, upperBounds(LayerIndex, UnknownIndex))
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
   ReDim TempIndexes(numberOfLayers - 1)
   For LayerIndex = 0 To numberOfLayers - 1
      TempIndexes(LayerIndex) = 0
   Next LayerIndex
   
   LettersASum = 0
   For SectionIndex = 0 To LetterIndexes(0)
      LettersASum = LettersASum + letters(0)(SectionIndex)
   Next SectionIndex
   
   getMu = (1 - numberOfLayers) * sumOfLetters _
      + (numberOfLayers - 2) * LettersASum _
      + 2 * upperBounds(0, UnknownIndex)
   
   For LayerIndex = 0 To numberOfLayers - 1
      For SectionIndex = 0 To LetterIndexes(LayerIndex)
         getMu = getMu + upperBounds(LayerIndex, getUnknownIndex(TempIndexes))
         TempIndexes(LayerIndex) = TempIndexes(LayerIndex) + 1
      Next SectionIndex
      TempIndexes(LayerIndex) = TempIndexes(LayerIndex) - 1
      getMu = getMu - upperBounds(0, getUnknownIndex(TempIndexes))
   Next LayerIndex
   
   Erase LetterIndexes
   Erase TempIndexes
End Function

Private Sub fillUnknowns()
   Dim UnknownIndex As Integer
   If diminishingUnknownIndex <> -1 Then
      unknowns(diminishingUnknownIndex) = unknowns(diminishingUnknownIndex) - 1
   End If
   For UnknownIndex = diminishingUnknownIndex + 1 To numberOfUnknowns - 1
      unknowns(UnknownIndex) = getUnknown(UnknownIndex)
      lowerBounds(UnknownIndex) = getMu(UnknownIndex)
      If lowerBounds(UnknownIndex) < 0 Then lowerBounds(UnknownIndex) = 0
   Next UnknownIndex
End Sub

Private Sub fillDiminishingUnknownIndex()
   Dim UnknownIndex As Integer
   diminishingUnknownIndex = -1
   For UnknownIndex = numberOfUnknowns - 1 To 0 Step -1
      If unknowns(UnknownIndex) > lowerBounds(UnknownIndex) Then
         diminishingUnknownIndex = UnknownIndex
         Exit For
      End If
   Next UnknownIndex
End Sub

Private Sub printArray(dgArray() As Integer, dgSize As Integer, dgFactorial As Boolean, RowIndex As Long, ColumnIndex As Integer)
   Dim i As Integer
   Dim Factorial As String
   Factorial = ""
   If dgFactorial Then Factorial = "!"
   For i = 0 To dgSize - 1
      Cells(RowIndex, ColumnIndex + i) = dgArray(i) & Factorial
   Next i
End Sub

Private Sub printPointersOfDenominator(ColumnIndex As Integer)
   Dim i As Integer
   Dim j As Integer
   Dim FactorGroupIndexes() As Integer
   For i = 0 To numberOfUnknowns - 1
      FactorGroupIndexes = getLetterIndexes(i)
      For j = 0 To numberOfLayers - 1
         Cells(j + 1, ColumnIndex + i) = degrees(j)(FactorGroupIndexes(j))
      Next j
   Next i
End Sub

Private Sub groupRepetitionsFromDenominator()
   Dim GroupIndex As Integer
   ReDim numeratorRepetitions(numberOfNumeratorDegrees - 1)
   For GroupIndex = 0 To numberOfNumeratorDegrees - 1
      numeratorRepetitions(GroupIndex) = 0
   Next GroupIndex
   For GroupIndex = 0 To numberOfUnknowns - 1
      numeratorRepetitions(numeratorConformity(GroupIndex)) = numeratorRepetitions(numeratorConformity(GroupIndex)) + unknowns(GroupIndex)
   Next GroupIndex
End Sub

Private Sub fillDegreesOfResult()
   Dim GroupIndex As Integer
   Dim j As Integer
   Dim k As Integer
   ReDim resultDegrees(sumOfLetters - 1)
   k = 0
   For GroupIndex = 0 To numberOfNumeratorDegrees - 1
      For j = 1 To numeratorRepetitions(GroupIndex)
         resultDegrees(k) = numeratorDegrees(GroupIndex)
         k = k + 1
      Next j
   Next GroupIndex
   'Debug.Print getInfo()
End Sub

' String functions

Public Function getInfo() As String
   Dim DebugString As String
   Dim i As Integer, j As Integer
   DebugString = "NumberOfLayers = " & numberOfLayers & vbLf & "SumOfLetters = " & sumOfLetters
   For i = 0 To numberOfLayers - 1
      DebugString = DebugString & vbLf & i & ":"
      For j = 0 To numberOfSections(i) - 1
         DebugString = DebugString & " " & letters(i)(j)
      Next j
   Next i
   DebugString = DebugString & vbLf & "NumberOfUnknowns = " & numberOfUnknowns
   getInfo = DebugString
End Function

Public Function getUnknownInfo() As String
   Dim i As Integer
   Dim DebugString As String
   For i = 0 To numberOfUnknowns - 1
      DebugString = DebugString & " " & unknowns(i)
   Next i
   DebugString = DebugString & " | " & diminishingUnknownIndex
   getUnknownInfo = DebugString
End Function

Public Function getLetterInfo() As String
   Dim DebugString As String
   Dim i As Integer
   Dim j As Integer
   DebugString = "Array of letters" & vbLf
   For i = 0 To numberOfLayers - 1
      DebugString = DebugString & "Factor " & i + 1 & ". Groups: " & numberOfSections(i) & ". Repetitions:"
      For j = 0 To numberOfSections(i) - 1
         DebugString = DebugString & " " & letters(i)(j)
      Next j
      DebugString = DebugString & vbLf
   Next i
   getLetterInfo = DebugString
End Function

' Destructor

Private Sub eraseArrays()
   Erase upperBounds
   Erase lowerBounds
   Erase letters
   Erase degrees
   Erase numberOfSections
   Erase unknowns
   Erase numeratorDegrees
   Erase numeratorRepetitions
   Erase resultDegrees
   Erase denominatorDegrees
   Erase factorDegrees
   Erase factorConformity
   Erase numeratorConformity
End Sub

