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
   Dim tempArray() As Integer
   Dim layerIndex As Integer
   Dim degreeIndex As Integer
   
   numberOfLayers = numberOfFactors
   sumOfLetters = numberOfDegrees
   ReDim factorDegrees(numberOfLayers - 1)
   ReDim factorConformity(numberOfLayers - 1)

   For layerIndex = 0 To numberOfLayers - 1
      ReDim tempArray(sumOfLetters - 1)
      factorDegrees(layerIndex) = tempArray
      For degreeIndex = 0 To sumOfLetters - 1
         factorDegrees(layerIndex)(degreeIndex) = Cells(1 + layerIndex, 5 + degreeIndex)
      Next degreeIndex
   Next layerIndex

   Erase tempArray
End Sub

Private Sub fillDegrees()
   Dim tempArrayFrom() As Integer
   Dim tempArrayTo() As Integer
   Dim layerIndex As Integer
   Dim degreeIndex As Integer
   Dim tempConformity() As Integer
   ReDim tempConformity(sumOfLetters - 1)
   ReDim degrees(numberOfLayers - 1)
   ReDim numberOfSections(numberOfLayers - 1)
   
   For layerIndex = 0 To numberOfLayers - 1
      tempArrayFrom = factorDegrees(layerIndex)
      groupArrays tempArrayFrom, sumOfLetters, tempArrayTo, numberOfSections(layerIndex), tempConformity
      degrees(layerIndex) = tempArrayTo
      factorConformity(layerIndex) = tempConformity
   Next layerIndex
   
   Erase tempArrayFrom
   Erase tempArrayTo
   Erase tempConformity
End Sub

Private Sub fillLetters()
   Dim layerIndex As Integer
   Dim degreeIndex As Integer
   Dim tempArray() As Integer
   ReDim letters(numberOfLayers - 1)
   For layerIndex = 0 To numberOfLayers - 1
      ReDim tempArray(sumOfLetters - 1)
      letters(layerIndex) = tempArray
      
      For degreeIndex = 0 To numberOfSections(layerIndex) - 1
         letters(layerIndex)(degreeIndex) = 0
      Next degreeIndex
      
      For degreeIndex = 0 To sumOfLetters - 1
         letters(layerIndex)(factorConformity(layerIndex)(degreeIndex)) _
            = letters(layerIndex)(factorConformity(layerIndex)(degreeIndex)) _
            + 1
      Next degreeIndex
   Next layerIndex
   Erase tempArray
End Sub

Private Sub fillDenominatorDegrees()
   Dim factorIndex As Integer
   Dim groupIndex As Integer
   Dim factorGroupIndexes() As Integer
   
   numberOfUnknowns = 1
   For factorIndex = 0 To numberOfLayers - 1
      numberOfUnknowns = numberOfUnknowns * numberOfSections(factorIndex)
   Next factorIndex
   
   ReDim denominatorDegrees(numberOfUnknowns - 1)
   ReDim numeratorConformity(numberOfUnknowns - 1)
   For groupIndex = 0 To numberOfUnknowns - 1
      factorGroupIndexes = getLetterIndexes(groupIndex)
      denominatorDegrees(groupIndex) = 0
      For factorIndex = 0 To numberOfLayers - 1
         denominatorDegrees(groupIndex) _
            = denominatorDegrees(groupIndex) _
            + degrees(factorIndex)(factorGroupIndexes(factorIndex))
      Next factorIndex
   Next groupIndex
   Erase factorGroupIndexes
   groupArrays denominatorDegrees, numberOfUnknowns, numeratorDegrees, numberOfNumeratorDegrees, numeratorConformity
End Sub

Private Sub printHeaders()
   Dim lastColumn As Integer
   Dim layerIndex As Integer
   Dim tempArray() As Integer
   prepareSheetBefore
   Cells(numberOfLayers + 1, 1).EntireRow.Font.Bold = True
   Range("A1") = "Number of factors"
   Range("B1") = numberOfLayers
   Range("A2") = "Number of degrees"
   Range("B2") = sumOfLetters
   For layerIndex = 0 To numberOfLayers - 1
      tempArray = factorDegrees(layerIndex)
      Cells(layerIndex + 1, 4) = "L["
      printArray tempArray, sumOfLetters, False, layerIndex + 1, 5
      Cells(layerIndex + 1, sumOfLetters + 5) = "]"
   Next layerIndex
   lastRow = numberOfLayers + 1
   lastColumn = sumOfLetters + 7
   printArray numeratorDegrees, numberOfNumeratorDegrees, False, lastRow, lastColumn
   lastColumn = lastColumn + numberOfNumeratorDegrees + 1
   printPointersOfDenominator lastColumn
   printArray denominatorDegrees, numberOfUnknowns, False, lastRow, lastColumn
   Erase tempArray
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
   Dim lastColumn As Integer
   lastRow = lastRow + 1
   lastColumn = sumOfLetters + 6
   Cells(lastRow, lastColumn) = "("
   printArray numeratorRepetitions, numberOfNumeratorDegrees, True, lastRow, lastColumn + 1
   lastColumn = lastColumn + numberOfNumeratorDegrees + 1
   Cells(lastRow, lastColumn) = ") : ("
   printArray unknowns, numberOfUnknowns, True, lastRow, lastColumn + 1
   lastColumn = lastColumn + numberOfUnknowns + 1
   Cells(lastRow, lastColumn) = ") L["
   printArray resultDegrees, sumOfLetters, False, lastRow, lastColumn + 1
   lastColumn = lastColumn + sumOfLetters + 1
   Cells(lastRow, lastColumn) = "]"
End Sub

Private Sub groupArrays(arrayFrom() As Integer, countFrom As Integer, arrayTo() As Integer, countTo As Integer, conformityArray() As Integer)
   Dim isFound As Boolean
   Dim indexFrom As Integer
   Dim indexTo As Integer
   ReDim conformityArray(countFrom - 1)
   ReDim arrayTo(countFrom - 1)
   countTo = 0
   
   For indexFrom = 0 To countFrom - 1
      isFound = False
      For indexTo = 0 To countTo - 1
         If arrayTo(indexTo) = arrayFrom(indexFrom) Then
            conformityArray(indexFrom) = indexTo
            isFound = True
            Exit For
         End If
      Next indexTo
      If (Not isFound) Then
         countTo = countTo + 1
         arrayTo(countTo - 1) = arrayFrom(indexFrom)
         conformityArray(indexFrom) = countTo - 1
      End If
   Next indexFrom
End Sub

Private Function getLetterIndexes(unknownIndex As Integer) As Integer()
   Dim i As Integer
   Dim tempIndex As Integer
   Dim letterIndexes() As Integer
   ReDim letterIndexes(numberOfLayers - 1)
   
   tempIndex = unknownIndex
   For i = numberOfLayers - 1 To 1 Step -1
      letterIndexes(i) = tempIndex Mod numberOfSections(i)
      tempIndex = tempIndex \ numberOfSections(i)
   Next i
   letterIndexes(0) = tempIndex
   getLetterIndexes = letterIndexes()
   Erase letterIndexes
End Function

Private Function getUnknownIndex(letterIndexes() As Integer) As Integer
   Dim i As Integer
   getUnknownIndex = letterIndexes(0)
   For i = 1 To numberOfLayers - 1
      getUnknownIndex = getUnknownIndex * numberOfSections(i)
      getUnknownIndex = getUnknownIndex + letterIndexes(i)
   Next i
End Function

Private Function isFirstIndex(layerIndex As Integer, unknownIndex As Integer) As Boolean
   Dim letterIndexes() As Integer
   Dim i As Integer
   letterIndexes = getLetterIndexes(unknownIndex)
   isFirstIndex = True
   For i = 0 To numberOfLayers - 1
      If (i <> layerIndex) And (letterIndexes(i) <> 0) Then
         isFirstIndex = False
         Exit For
      End If
   Next i
End Function

Private Function getPreviousIndex(layerIndex As Integer, unknownIndex As Integer) As Integer
   Dim i As Integer
   Dim previousLetterIndexes() As Integer
   ReDim previousLetterIndexes(numberOfLayers - 1)
   previousLetterIndexes = getLetterIndexes(unknownIndex)
   For i = numberOfLayers - 1 To 0 Step -1
      If (i <> layerIndex) And (previousLetterIndexes(i) > 0) Then
         previousLetterIndexes(i) = previousLetterIndexes(i) - 1
         Exit For
      ElseIf (i <> layerIndex) And (previousLetterIndexes(i) = 0) Then
         previousLetterIndexes(i) = numberOfSections(i) - 1
      End If
   Next i
   getPreviousIndex = getUnknownIndex(previousLetterIndexes)
   Erase previousLetterIndexes
End Function

Private Function getUnknown(unknownIndex As Integer) As Integer
   Dim layerIndex As Integer
   Dim previousIndex As Integer
   Dim letterIndexes() As Integer
   letterIndexes = getLetterIndexes(unknownIndex)
   
   For layerIndex = 0 To numberOfLayers - 1
      If isFirstIndex(layerIndex, unknownIndex) Then
         upperBounds(layerIndex, unknownIndex) = letters(layerIndex)(letterIndexes(layerIndex))
      Else
         previousIndex = getPreviousIndex(layerIndex, unknownIndex)
         upperBounds(layerIndex, unknownIndex) = upperBounds(layerIndex, previousIndex) - unknowns(previousIndex)
      End If
   Next layerIndex
   getUnknown = upperBounds(0, unknownIndex)
   
   For layerIndex = 0 To numberOfLayers - 1
      getUnknown = WorksheetFunction.Min(getUnknown, upperBounds(layerIndex, unknownIndex))
   Next layerIndex
   Erase letterIndexes
End Function

Private Function getMu(unknownIndex As Integer) As Integer
   Dim layerIndex As Integer
   Dim sectionIndex As Integer
   Dim lettersASum As Integer
   Dim tempIndexes() As Integer
   Dim letterIndexes() As Integer
   
   letterIndexes = getLetterIndexes(unknownIndex)
   ReDim tempIndexes(numberOfLayers - 1)
   For layerIndex = 0 To numberOfLayers - 1
      tempIndexes(layerIndex) = 0
   Next layerIndex
   
   lettersASum = 0
   For sectionIndex = 0 To letterIndexes(0)
      lettersASum = lettersASum + letters(0)(sectionIndex)
   Next sectionIndex
   
   getMu = (1 - numberOfLayers) * sumOfLetters _
      + (numberOfLayers - 2) * lettersASum _
      + 2 * upperBounds(0, unknownIndex)
   
   For layerIndex = 0 To numberOfLayers - 1
      For sectionIndex = 0 To letterIndexes(layerIndex)
         getMu = getMu + upperBounds(layerIndex, getUnknownIndex(tempIndexes))
         tempIndexes(layerIndex) = tempIndexes(layerIndex) + 1
      Next sectionIndex
      tempIndexes(layerIndex) = tempIndexes(layerIndex) - 1
      getMu = getMu - upperBounds(0, getUnknownIndex(tempIndexes))
   Next layerIndex
   
   Erase letterIndexes
   Erase tempIndexes
End Function

Private Sub fillUnknowns()
   Dim unknownIndex As Integer
   If diminishingUnknownIndex <> -1 Then
      unknowns(diminishingUnknownIndex) = unknowns(diminishingUnknownIndex) - 1
   End If
   For unknownIndex = diminishingUnknownIndex + 1 To numberOfUnknowns - 1
      unknowns(unknownIndex) = getUnknown(unknownIndex)
      lowerBounds(unknownIndex) = getMu(unknownIndex)
      If lowerBounds(unknownIndex) < 0 Then lowerBounds(unknownIndex) = 0
   Next unknownIndex
End Sub

Private Sub fillDiminishingUnknownIndex()
   Dim unknownIndex As Integer
   diminishingUnknownIndex = -1
   For unknownIndex = numberOfUnknowns - 1 To 0 Step -1
      If unknowns(unknownIndex) > lowerBounds(unknownIndex) Then
         diminishingUnknownIndex = unknownIndex
         Exit For
      End If
   Next unknownIndex
End Sub

Private Sub printArray(dgArray() As Integer, dgSize As Integer, dgFactorial As Boolean, rowIndex As Long, columnIndex As Integer)
   Dim i As Integer
   Dim factorial As String
   factorial = ""
   If dgFactorial Then factorial = "!"
   For i = 0 To dgSize - 1
      Cells(rowIndex, columnIndex + i) = dgArray(i) & factorial
   Next i
End Sub

Private Sub printPointersOfDenominator(columnIndex As Integer)
   Dim i As Integer
   Dim j As Integer
   Dim factorGroupIndexes() As Integer
   For i = 0 To numberOfUnknowns - 1
      factorGroupIndexes = getLetterIndexes(i)
      For j = 0 To numberOfLayers - 1
         Cells(j + 1, columnIndex + i) = degrees(j)(factorGroupIndexes(j))
      Next j
   Next i
End Sub

Private Sub groupRepetitionsFromDenominator()
   Dim groupIndex As Integer
   ReDim numeratorRepetitions(numberOfNumeratorDegrees - 1)
   For groupIndex = 0 To numberOfNumeratorDegrees - 1
      numeratorRepetitions(groupIndex) = 0
   Next groupIndex
   For groupIndex = 0 To numberOfUnknowns - 1
      numeratorRepetitions(numeratorConformity(groupIndex)) = numeratorRepetitions(numeratorConformity(groupIndex)) + unknowns(groupIndex)
   Next groupIndex
End Sub

Private Sub fillDegreesOfResult()
   Dim groupIndex As Integer
   Dim j As Integer
   Dim k As Integer
   ReDim resultDegrees(sumOfLetters - 1)
   k = 0
   For groupIndex = 0 To numberOfNumeratorDegrees - 1
      For j = 1 To numeratorRepetitions(groupIndex)
         resultDegrees(k) = numeratorDegrees(groupIndex)
         k = k + 1
      Next j
   Next groupIndex
   'Debug.Print getInfo()
End Sub

' String functions

Public Function getInfo() As String
   Dim debugString As String
   Dim i As Integer, j As Integer
   debugString = "NumberOfLayers = " & numberOfLayers & vbLf & "SumOfLetters = " & sumOfLetters
   For i = 0 To numberOfLayers - 1
      debugString = debugString & vbLf & i & ":"
      For j = 0 To numberOfSections(i) - 1
         debugString = debugString & " " & letters(i)(j)
      Next j
   Next i
   debugString = debugString & vbLf & "NumberOfUnknowns = " & numberOfUnknowns
   getInfo = debugString
End Function

Public Function getUnknownInfo() As String
   Dim i As Integer
   Dim debugString As String
   For i = 0 To numberOfUnknowns - 1
      debugString = debugString & " " & unknowns(i)
   Next i
   debugString = debugString & " | " & diminishingUnknownIndex
   getUnknownInfo = debugString
End Function

Public Function getLetterInfo() As String
   Dim debugString As String
   Dim i As Integer
   Dim j As Integer
   debugString = "Array of letters" & vbLf
   For i = 0 To numberOfLayers - 1
      debugString = debugString & "Factor " & i + 1 & ". Groups: " & numberOfSections(i) & ". Repetitions:"
      For j = 0 To numberOfSections(i) - 1
         debugString = debugString & " " & letters(i)(j)
      Next j
      debugString = debugString & vbLf
   Next i
   getLetterInfo = debugString
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

