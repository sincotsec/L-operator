Attribute VB_Name = "Common"
Option Explicit

Enum dgItem
   dgRepetition = 0
   dgDegree = 1
End Enum

Function getMaximum(Number1 As Integer, Number2 As Integer) As Integer
   getMaximum = Number2
   If Number1 >= Number2 Then getMaximum = Number1
End Function

Function getMinimum(Number1 As Integer, Number2 As Integer) As Integer
   getMinimum = Number2
   If Number1 <= Number2 Then getMinimum = Number1
End Function

