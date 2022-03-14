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

Dim NumberOfLayers As Integer
Dim SumOfLetters As Integer
Dim Letters() As Variant
Dim Degrees() As Variant
Dim NumberOfSections() As Integer

Dim Unknowns() As Integer
Dim mNumberOfUnknowns As Integer
Dim UpperBounds() As Integer
Dim LowerBounds() As Integer
Dim DiminishingUnknownIndex As Integer

Dim NumeratorDegrees() As Integer
Dim NumeratorRepetitions() As Integer
Dim ResultDegrees() As Integer
Dim DenominatorDegrees() As Integer

Dim NumberOfNumeratorDegrees As Integer

' Property Get

Public Property Get NumberOfUnknowns()
   NumberOfUnknowns = mNumberOfUnknowns
End Property

' Methods

Public Sub fillArrays(NumberOfFactors As Integer, NumberOfDegrees As Integer)
    Dim UngroupedDegrees() As Integer
    Dim TempArray() As Integer
    Dim LayerIndex As Integer
    Dim DegreeIndex As Integer
    Dim ConformityArray() As Integer
    Dim isFound As Boolean
    Dim SectionIndex As Integer
    
    NumberOfLayers = NumberOfFactors
    SumOfLetters = NumberOfDegrees
    
    ReDim Letters(NumberOfLayers - 1)
    ReDim Degrees(NumberOfLayers - 1)
    ReDim NumberOfSections(NumberOfLayers - 1)
    ReDim UngroupedDegrees(NumberOfLayers - 1, SumOfLetters - 1)
    ReDim ConformityArray(SumOfLetters - 1)
    For LayerIndex = 0 To NumberOfLayers - 1
        For DegreeIndex = 0 To SumOfLetters - 1
            UngroupedDegrees(LayerIndex, DegreeIndex) = Cells(1 + LayerIndex, 5 + DegreeIndex)
        Next DegreeIndex
        ReDim TempArray(SumOfLetters - 1)
        
        Degrees(LayerIndex) = TempArray
        Letters(LayerIndex) = TempArray
        NumberOfSections(LayerIndex) = 0
        For DegreeIndex = 0 To SumOfLetters - 1
            isFound = False
            For SectionIndex = 0 To NumberOfSections(LayerIndex) - 1
                If Degrees(LayerIndex)(SectionIndex) = UngroupedDegrees(LayerIndex, DegreeIndex) Then
                    ConformityArray(DegreeIndex) = SectionIndex
                    isFound = True
                    Exit For
                End If
            Next SectionIndex
            If (Not isFound) Then
                NumberOfSections(LayerIndex) = NumberOfSections(LayerIndex) + 1
                Degrees(LayerIndex)(NumberOfSections(LayerIndex) - 1) = UngroupedDegrees(LayerIndex, DegreeIndex)
                ConformityArray(DegreeIndex) = NumberOfSections(LayerIndex) - 1
            End If
            Letters(LayerIndex)(ConformityArray(DegreeIndex)) = Letters(LayerIndex)(ConformityArray(DegreeIndex)) + 1
        Next DegreeIndex
    Next LayerIndex
    Erase UngroupedDegrees
    Erase TempArray
    Erase ConformityArray
End Sub

Public Sub fillDegreesOfDenominator()
   Dim FactorIndex As Integer
   Dim GroupIndex As Integer
   Dim FactorGroupIndexes() As Integer
   ReDim DenominatorDegrees(mNumberOfUnknowns - 1)
   For GroupIndex = 0 To mNumberOfUnknowns - 1
      FactorGroupIndexes = getLetterIndexes(GroupIndex)
      For FactorIndex = 0 To NumberOfLayers - 1
         DenominatorDegrees(GroupIndex) = DenominatorDegrees(GroupIndex) + Degrees(FactorIndex)(FactorGroupIndexes(FactorIndex))
      Next FactorIndex
   Next GroupIndex
   Erase FactorGroupIndexes
   groupDegrees
   'Result.allocateMemory NumberOfDegrees
End Sub

Public Sub groupDegrees()
    Dim DegreeIndex As Integer
    Dim ConformityArray() As Integer
    Dim isFound As Boolean
    Dim SectionIndex As Integer
    ReDim ConformityArray(mNumberOfUnknowns - 1)
    ReDim NumeratorDegrees(mNumberOfUnknowns - 1)
    NumberOfNumeratorDegrees = 0
    For DegreeIndex = 0 To mNumberOfUnknowns - 1
        isFound = False
        For SectionIndex = 0 To NumberOfNumeratorDegrees - 1
            If NumeratorDegrees(SectionIndex) = DenominatorDegrees(DegreeIndex) Then
                ConformityArray(DegreeIndex) = SectionIndex
                isFound = True
                Exit For
            End If
        Next SectionIndex
        If (Not isFound) Then
            NumberOfNumeratorDegrees = NumberOfNumeratorDegrees + 1
            NumeratorDegrees(NumberOfNumeratorDegrees - 1) = DenominatorDegrees(DegreeIndex)
            ConformityArray(DegreeIndex) = NumberOfNumeratorDegrees - 1
        End If
    Next DegreeIndex
    Erase ConformityArray
End Sub

Public Sub prepareSolution()
   Dim i As Integer
   mNumberOfUnknowns = 1
   For i = 0 To NumberOfLayers - 1
      mNumberOfUnknowns = mNumberOfUnknowns * NumberOfSections(i)
   Next i
   ReDim Unknowns(mNumberOfUnknowns - 1)
   ReDim UpperBounds(NumberOfLayers - 1, mNumberOfUnknowns - 1)
   ReDim LowerBounds(mNumberOfUnknowns - 1)
   DiminishingUnknownIndex = -1
End Sub

Public Function getLetterIndexes(ByVal UnknownIndex As Integer) As Integer()
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

Public Function getUnknownIndex(ByRef LetterIndexes() As Integer) As Integer
   Dim i As Integer
   getUnknownIndex = LetterIndexes(0)
   For i = 1 To NumberOfLayers - 1
      getUnknownIndex = getUnknownIndex * NumberOfSections(i)
      getUnknownIndex = getUnknownIndex + LetterIndexes(i)
   Next i
End Function

Function isFirstIndex(ByVal LayerIndex As Integer, ByVal UnknownIndex As Integer) As Boolean
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

Public Function getUnknown(UnknownIndex As Integer) As Integer
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
      getUnknown = getMinimum(getUnknown, UpperBounds(LayerIndex, UnknownIndex))
   Next LayerIndex
   Erase LetterIndexes
End Function

Public Function getMu(UnknownIndex As Integer) As Integer
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

Public Function getDiminishingUnknownIndex() As Integer
   Dim UnknownIndex As Integer
   getDiminishingUnknownIndex = -1
   For UnknownIndex = mNumberOfUnknowns - 1 To 0 Step -1
      If Unknowns(UnknownIndex) > LowerBounds(UnknownIndex) Then
         getDiminishingUnknownIndex = UnknownIndex
         Exit For
      End If
   Next UnknownIndex
End Function

Public Sub fillUnknowns()
   Dim UnknownIndex As Integer
   If DiminishingUnknownIndex <> -1 Then
      Unknowns(DiminishingUnknownIndex) = Unknowns(DiminishingUnknownIndex) - 1
   End If
   For UnknownIndex = 0 To mNumberOfUnknowns - 1
      If UnknownIndex > DiminishingUnknownIndex Then
         Unknowns(UnknownIndex) = getUnknown(UnknownIndex)
         LowerBounds(UnknownIndex) = getMu(UnknownIndex)
         If LowerBounds(UnknownIndex) < 0 Then LowerBounds(UnknownIndex) = 0
      End If
   Next UnknownIndex
   DiminishingUnknownIndex = getDiminishingUnknownIndex()
End Sub

Public Function getUnknownArray() As Integer()
   getUnknownArray = Unknowns()
End Function

Public Function isDone() As Boolean
   isDone = False
   If DiminishingUnknownIndex = -1 Then isDone = True
End Function

' Print functions

Public Sub printArray(dgArray() As Integer, dgSize As Integer, ByVal dgFactorial As Boolean, ByVal RowIndex As Integer, ByVal ColumnIndex As Integer)
    Dim i As Integer
    Dim Factorial As String
    Factorial = ""
    If dgFactorial Then Factorial = "!"
    For i = 0 To dgSize - 1
        Sheets(1).Cells(RowIndex, ColumnIndex + i) = dgArray(i) & Factorial
    Next i
End Sub

Public Sub printUnknowns()
    Call printArray(NumeratorDegrees, NumberOfNumeratorDegrees, False, 13, 1)
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
   DebugString = DebugString & vbLf & "NumberOfUnknowns = " & mNumberOfUnknowns
   getInfo = DebugString
End Function

Public Function getUnknownInfo() As String
   Dim i As Integer
   Dim DebugString As String
   For i = 0 To mNumberOfUnknowns - 1
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
End Sub
