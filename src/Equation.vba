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

Public Sub allocateMemory(NumberOfFactors As Integer, NumberOfDegrees As Integer)
   NumberOfLayers = NumberOfFactors
   SumOfLetters = NumberOfDegrees
   ReDim Letters(NumberOfLayers)
   ReDim NumberOfSections(NumberOfLayers)
End Sub

Public Sub fillString(SheetRow As Integer)
   Dim i As Integer
   Dim Degrees() As Integer
   Dim SecondArray() As Integer
   ReDim Degrees(SumOfLetters)
   Sheets(1).Select
   For i = 0 To SumOfLetters - 1
      Degrees(i) = Cells(SheetRow, 3 + i)
   Next i
End Sub
