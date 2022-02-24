Attribute VB_Name = "Export"
Option Explicit

Sub exportProject()
    Dim i As Integer
    Dim FSO As Scripting.FileSystemObject
    Dim Project As Object
    Set Project = ThisWorkbook.VBProject
    Dim FileName As String
    Const RepoPath = "C:\sincotsec\L-operator\src"
    Set FSO = New Scripting.FileSystemObject
    If FSO.FolderExists(RepoPath) Then FSO.DeleteFolder RepoPath
    FSO.CreateFolder RepoPath
    For i = 1 To Project.VBComponents.Count
        Project.VBComponents(i).Export RepoPath & "\" & Project.VBComponents(i).Name & ".vba"
    Next i
    Set FSO = Nothing
    Set Project = Nothing
End Sub
