VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Operator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public NumberOfGroups As Integer
Dim Groups() As Group

Public FirstColumn As Integer
Public LastColumn As Integer
Public Hue As Double

' Property Get

Public Property Get Degree(GroupIndex As Integer) As Integer
   Degree = Groups(GroupIndex).Degree
End Property

Public Property Get Repetition(GroupIndex As Integer) As Integer
   Repetition = Groups(GroupIndex).Repetition
End Property

' Property Let

Public Property Let Degree(GroupIndex As Integer, Value As Integer)
   Groups(GroupIndex).Degree = Value
End Property

Public Property Let Repetition(GroupIndex As Integer, Value As Integer)
   Groups(GroupIndex).Repetition = Value
End Property

' Methods

Public Sub allocateMemory(parNumberOfGroups As Integer)
   NumberOfGroups = parNumberOfGroups
   ReDim Groups(NumberOfGroups)
End Sub

Public Sub fillStringFactor(SheetRow As Integer)
   Dim i As Integer
   For i = 0 To NumberOfGroups - 1
      Groups(i).Degree = Sheets(1).Cells(SheetRow, 3 + i)
      Groups(i).Repetition = 1
   Next i
End Sub

Public Sub groupDegreesFromOperator(OperatorFrom As Operator, ConformityArray() As Integer)
   Dim isFound As Boolean
   Dim i As Integer
   Dim j As Integer
   Dim TemporaryOperator As Operator
   Set TemporaryOperator = New Operator
   ReDim ConformityArray(OperatorFrom.NumberOfGroups)
   TemporaryOperator.allocateMemory OperatorFrom.NumberOfGroups
   NumberOfGroups = 0
   For i = 0 To OperatorFrom.NumberOfGroups - 1
      isFound = False
      For j = 0 To NumberOfGroups - 1
         If TemporaryOperator.Degree(j) = OperatorFrom.Degree(i) Then
            ConformityArray(i) = j
            isFound = True
            Exit For
         End If
      Next j
      If (Not isFound) Then
         NumberOfGroups = NumberOfGroups + 1
         TemporaryOperator.Degree(NumberOfGroups - 1) = OperatorFrom.Degree(i)
         ConformityArray(i) = NumberOfGroups - 1
      End If
   Next i

   ReDim Groups(NumberOfGroups)
   For i = 0 To NumberOfGroups - 1
      Groups(i).Degree = TemporaryOperator.Degree(i)
   Next i
   Set TemporaryOperator = Nothing
End Sub

Public Sub prepareTitle(ByVal RowIndex As Integer)
   With Range(Sheets(2).Cells(RowIndex, FirstColumn), Sheets(2).Cells(RowIndex, LastColumn))
      .Font.Bold = True
      .EntireColumn.Interior.Color = ColorFromHSL(Hue, 100, 60)
   End With
End Sub

Public Sub printItemOfGroup(dgItem As Integer, ByVal RowIndex As Integer)
   Dim i As Integer
   For i = 0 To NumberOfGroups - 1
      Select Case dgItem
         Case dgRepetition
            Sheets(2).Cells(RowIndex, FirstColumn + i) = Groups(i).Repetition
         Case dgDegree
            Sheets(2).Cells(RowIndex, FirstColumn + i) = Groups(i).Degree
      End Select
   Next i
End Sub

Public Function getInfo() As String
   Dim i As Integer
   getInfo = vbLf & "Number of groups: " & NumberOfGroups & vbLf & "Degrees: "
   For i = 0 To NumberOfGroups - 1
      getInfo = getInfo & " " & Groups(i).Degree & "[" & Groups(i).Repetition & "]"
   Next i
End Function

Public Sub setColumns(LastColumnOfPreviousOperator As Integer, parHue As Integer)
   FirstColumn = LastColumnOfPreviousOperator + 1
   LastColumn = LastColumnOfPreviousOperator + NumberOfGroups
   Hue = parHue
   If Hue > 360 Then Hue = Hue - 360
End Sub

' Destructor

Private Sub Class_Terminate()
   Erase Groups
End Sub
