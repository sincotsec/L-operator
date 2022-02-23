Attribute VB_Name = "Export"
Option Explicit

Sub exportProject()
    Dim i As Integer
    Dim FSO As Scripting.FileSystemObject
    Dim FileName As String
    Const RepoPath = "C:\sincotsec\L-operator\src"
    Set FSO = New Scripting.FileSystemObject
    If FSO.FolderExists(RepoPath) Then FSO.DeleteFolder RepoPath
    FSO.CreateFolder RepoPath
    With ThisWorkbook.VBProject
        For i = 1 To .VBComponents.Count
            .VBComponents(i).Export RepoPath & "\" & .VBComponents(i).Name & ".vba"
        Next i
    End With
    Set FSO = Nothing
End Sub
