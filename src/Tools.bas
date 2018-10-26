Attribute VB_Name = "Tools"
' http://stackoverflow.com/questions/915317/does-vba-have-dictionary-structure
Public Function ColContains(col As Collection, key As Variant) As Boolean
    On Error Resume Next
    col (key) ' Just try it. If it fails, Err.Number will be nonzero.
    Contains = (Err.Number = 0)
    Err.Clear
End Function

Public Sub SplitTagServer(ByVal tagPath As String, ByRef tagServer As String, ByRef tagName As String)

    For L = Len(tagPath) To 1 Step -1
        value = Mid(tagPath, L, 1)
        'Check for split
        If Mid(tagPath, L, 1) = "\" Then
        'split between server id and Tag Name found
            tagName = Mid(tagPath, L + 1, Len(tagPath))
            tagServer = Mid(tagPath, 3, L - 3)
            Exit For
        ElseIf L <= 1 Then
            tagName = tagPath
            tagServer = ""

        End If
    Next
End Sub


'Check if the selected directory is a folder
'Returns true if it's a folder, false otherwise.
Public Function FolderExists(tempdirectory As String) As Boolean
    On Error GoTo ErrHandler
    If Dir(tempdirectory & "\", vbDirectory) = vbNullString Then
        FolderExists = False
    Else
        FolderExists = True
    End If
    Exit Function
  
ErrHandler:
    FolderExists = False
End Function
