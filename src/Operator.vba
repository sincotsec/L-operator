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

Dim mNumberOfGroups As Integer
Dim Degrees() As Integer
Dim Repetitions() As Integer

Dim mFirstColumn As Integer
Dim mLastColumn As Integer
Dim mHue As Double

' Property Get

Public Property Get NumberOfGroups() As Integer
   NumberOfGroups = mNumberOfGroups
End Property

Public Property Get Degree(GroupIndex As Integer) As Integer
   Degree = Degrees(GroupIndex)
End Property

Public Property Get Repetition(GroupIndex As Integer) As Integer
   Repetition = Repetitions(GroupIndex)
End Property

Public Property Get FirstColumn() As Integer
   FirstColumn = mFirstColumn
End Property

Public Property Get LastColumn() As Integer
   LastColumn = mLastColumn
End Property

' Property Let

Public Property Let Degree(GroupIndex As Integer, Value As Integer)
   Degrees(GroupIndex) = Value
End Property

Public Property Let Repetition(GroupIndex As Integer, Value As Integer)
   Repetitions(GroupIndex) = Value
End Property

' Methods

Public Sub allocateMemory(parNumberOfGroups As Integer)
   mNumberOfGroups = parNumberOfGroups
   ReDim Degrees(mNumberOfGroups - 1)
   ReDim Repetitions(mNumberOfGroups - 1)
End Sub

Public Sub fillStringFactor(SheetRow As Integer)
   Dim i As Integer
   For i = 0 To mNumberOfGroups - 1
      Degrees(i) = Sheets(1).Cells(SheetRow, 3 + i)
      Repetitions(i) = 1
   Next i
End Sub

Public Sub getDegreeString()
   Dim DebugString As String
   Dim i As Integer, j As Integer
   For i = 0 To mNumberOfGroups - 1
      For j = 1 To Repetitions(i)
         DebugString = DebugString & " " & Degrees(i)
      Next j
   Next i
   Debug.Print DebugString
End Sub

Public Sub groupDegreesFromOperator(OperatorFrom As Operator, ConformityArray() As Integer)
   Dim isFound As Boolean
   Dim i As Integer
   Dim j As Integer
   Dim TemporaryOperator As Operator
   Set TemporaryOperator = New Operator
   ReDim ConformityArray(OperatorFrom.NumberOfGroups - 1)
   TemporaryOperator.allocateMemory OperatorFrom.NumberOfGroups
   mNumberOfGroups = 0
   For i = 0 To OperatorFrom.NumberOfGroups - 1
      isFound = False
      For j = 0 To mNumberOfGroups - 1
         If TemporaryOperator.Degree(j) = OperatorFrom.Degree(i) Then
            ConformityArray(i) = j
            isFound = True
            Exit For
         End If
      Next j
      If (Not isFound) Then
         mNumberOfGroups = mNumberOfGroups + 1
         TemporaryOperator.Degree(mNumberOfGroups - 1) = OperatorFrom.Degree(i)
         ConformityArray(i) = mNumberOfGroups - 1
      End If
   Next i
   ReDim Degrees(mNumberOfGroups - 1)
   ReDim Repetitions(mNumberOfGroups - 1)
   For i = 0 To mNumberOfGroups - 1
      Degrees(i) = TemporaryOperator.Degree(i)
   Next i
   Set TemporaryOperator = Nothing
End Sub

Public Sub groupRepetitionsFromOperator(OperatorFrom As Operator, ConformityArray() As Integer)
   Dim GroupIndex As Integer
   For GroupIndex = 0 To mNumberOfGroups - 1
      Repetitions(GroupIndex) = 0
   Next GroupIndex
   For GroupIndex = 0 To OperatorFrom.NumberOfGroups - 1
      Repetitions(ConformityArray(GroupIndex)) = Repetitions(ConformityArray(GroupIndex)) + OperatorFrom.Repetition(GroupIndex)
   Next GroupIndex
End Sub

Public Sub prepareTitle(ByVal RowIndex As Integer)
   With Range(Sheets(2).Cells(RowIndex, FirstColumn), Sheets(2).Cells(RowIndex, LastColumn))
      .Font.Bold = True
      .EntireColumn.Interior.Color = ColorFromHSL(mHue, 100, 60)
   End With
End Sub

Public Sub printItemOfGroup(dgItem As Integer, ByVal RowIndex As Integer)
   Dim i As Integer
   For i = 0 To mNumberOfGroups - 1
      Select Case dgItem
         Case dgRepetition
            Sheets(2).Cells(RowIndex, FirstColumn + i) = Repetitions(i)
         Case dgDegree
            Sheets(2).Cells(RowIndex, FirstColumn + i) = Degrees(i)
      End Select
   Next i
End Sub

Public Function getInfo() As String
   Dim i As Integer
   getInfo = vbLf & "Number of groups: " & mNumberOfGroups & vbLf & "Degrees: "
   For i = 0 To mNumberOfGroups - 1
      getInfo = getInfo & " " & Degrees(i) & "[" & Repetitions(i) & "]"
   Next i
End Function

Public Sub setColumns(LastColumnOfPreviousOperator As Integer, parHue As Integer)
   mFirstColumn = LastColumnOfPreviousOperator + 1
   mLastColumn = LastColumnOfPreviousOperator + mNumberOfGroups
   mHue = parHue
End Sub

' Destructor

Private Sub Class_Terminate()
   Erase Degrees
   Erase Repetitions
End Sub
