Attribute VB_Name = "export"
Option Explicit

Sub exportProject()
   Dim i As Integer
   Dim FSO As Scripting.FileSystemObject
   Dim FileName As String
   Dim Components As Object
   Const RepoPath = "C:\sincotsec\L-operator\src"
   Set FSO = New Scripting.FileSystemObject
   If FSO.FolderExists(RepoPath) Then FSO.DeleteFolder RepoPath
   FSO.CreateFolder RepoPath
   Set FSO = Nothing
   Set Components = ThisWorkbook.VBProject.VBComponents
   For i = 1 To Components.Count
      Components(i).export RepoPath & "\" & Components(i).Name & ".bas"
   Next i
   Set Components = Nothing
End Sub
