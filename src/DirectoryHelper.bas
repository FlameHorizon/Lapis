Attribute VB_Name = "DirectoryHelper"
Option Explicit


Private Const ModuleName As String = "Directory"

Public Const DirectorySeparatorChar As String = "\"


' Creates a whole folder structure using given folder path.
' Note: Only DirectorySeparatorChar character used to separate directory levels in a path string
' that reflects a hierarchical file system organization is supported.
' Throws ArgumentException when path is empty string.
' Throws ArgumentException when path does not contains supported directory separator.
Public Sub CreateDirectory(ByVal Path As String)

    Const MethodName = "CreateDirectory"
    
    If Path = vbNullString Then
        Exception.ArgumentException "Path", _
                                    "Path can't be an empty string. " _
                                  & ModuleName & "." & MethodName
    End If
    
    If StringHelper.Contains(Path, DirectorySeparatorChar) = False Then
        Exception.ArgumentException "Path", _
                                    "Path does have valid directory separator. " _
                                  & ModuleName & "." & MethodName
    End If
    
    
    ' To avoid problems with creating new folders,
    ' backslash from the end of folder path must
    ' be removed.
    Dim NormalizedPath As String
    NormalizedPath = RemoveTrailingBackslash(Path)

    Dim FolderPath As String
    Dim PathSection As Variant
    For Each PathSection In GetPathSections(NormalizedPath)
        FolderPath = FolderPath & PathSection & "\"

        If Not Tools.Fso.FolderExists(FolderPath) Then
            Tools.Fso.CreateFolder FolderPath
        End If
    Next PathSection
    
End Sub


Public Sub RemoveAllFilesFromFolder(ByVal Path As String)
    Kill Path & "\*.*"
End Sub


Private Function RemoveTrailingBackslash(ByVal Path As String) As String

    If Path Like "*\" Then
        RemoveTrailingBackslash = VBA.Left$(Path, Len(Path) - 1)
    Else
        RemoveTrailingBackslash = Path
    End If

End Function


' Splits a path to a file into individual folder segemetns.
' Example: "C:\abc1\abc1" will be splited into array with elements "C:", "abc1" and "abc2".
Private Function GetPathSections(ByVal Path As String) As String()
    GetPathSections = Split(Path, "\")
End Function



