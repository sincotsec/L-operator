VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Equation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public NumberOfLayers As Integer
Public SumOfLetters As Integer
Dim Letters() As Variant
Dim NumberOfSections() As Integer
Dim Unknowns() As Integer

Public Sub allocateMemory(NumberOfFactors As Integer, NumberOfDegrees As Integer)
   NumberOfLayers = NumberOfFactors
   SumOfLetters = NumberOfDegrees
   ReDim Letters(NumberOfLayers - 1)
   ReDim NumberOfSections(NumberOfLayers - 1)
End Sub

Public Function getInfo() As String
   Dim DebugString As String
   Dim i As Integer, j As Integer
   DebugString = "NumberOfLayers = " & NumberOfLayers _
      & vbLf & "SumOfLetters = " & SumOfLetters
   For i = 0 To NumberOfLayers - 1
      DebugString = DebugString & vbLf & i & ":"
      For j = 0 To NumberOfSections(i) - 1
         DebugString = DebugString & " " & Letters(i)(j)
      Next j
   Next i
   getInfo = DebugString
End Function

Public Sub fillArray(ByRef FactorsArray() As Operator)
   Dim DebugString As String
   Dim i As Integer, j As Integer
   Dim SecondArray() As Integer
   For i = 0 To NumberOfLayers - 1
      NumberOfSections(i) = FactorsArray(i).NumberOfGroups
      ReDim SecondArray(NumberOfSections(i) - 1)
      For j = 0 To NumberOfSections(i) - 1
         SecondArray(j) = FactorsArray(i).Repetition(j)
      Next j
      Letters(i) = SecondArray
   Next i
   Erase SecondArray
End Sub
