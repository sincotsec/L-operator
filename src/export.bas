Attribute VB_Name = "export"
Option Explicit

Sub exportProject()
   Dim i As Integer
   Dim FSO As Scripting.FileSystemObject
   Dim fileName As String
   Dim components As Object
   Const repoPath = "C:\sincotsec\L-operator\src"
   Set FSO = New Scripting.FileSystemObject
   If FSO.FolderExists(repoPath) Then FSO.DeleteFolder repoPath
   FSO.CreateFolder repoPath
   Set FSO = Nothing
   Set components = ThisWorkbook.VBProject.VBComponents
   For i = 1 To components.Count
      components(i).export repoPath & "\" & components(i).Name & ".bas"
   Next i
   Set components = Nothing
End Sub
