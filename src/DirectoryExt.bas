Attribute VB_Name = "DirectoryExt"
Option Explicit
'@Folder("Helper")

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
        Errors.OnArgumentError "Path", _
                                    "Path can't be an empty string. " _
                                    & ModuleName & "." & MethodName
    End If
    
    If StringExt.Contains(Path, DirectorySeparatorChar) = False Then
        Errors.OnArgumentError "Path", _
                                    "Path does have valid directory separator. " _
                                    & ModuleName & "." & MethodName
    End If
    
    
    ' To avoid problems with creating new folders,
    ' backslash from the end of folder path must
    ' be removed if, it exists.
    Dim NormalizedPath As String
    NormalizedPath = IIf(NeedsNormalization(Path), NormalizePath(Path), Path)
    CreateDirectoryBySections GetPathSections(NormalizedPath)
    
End Sub


Public Sub RemoveAllFilesFromFolder(ByVal Path As String)
    Kill Path & "\*.*"
End Sub


Private Function NeedsNormalization(ByVal Path As String) As Boolean
    NeedsNormalization = Path Like "*\"
End Function


' Path is normalized with the last character is backslash and it's removed.
' Remove last character from the right side of the string.
Private Function NormalizePath(ByVal Path As String) As String
    NormalizePath = VBA.Left$(Path, Len(Path) - 1)
End Function


' Splits a path to a file into individual folder segemetns.
' Example: "C:\abc1\abc1" will be splited into array with elements "C:", "abc1" and "abc2".
Private Function GetPathSections(ByVal Path As String) As String()
    GetPathSections = Split(Path, "\")
End Function


Private Sub CreateDirectoryBySections(ByVal PathSections As Variant)
    
    Dim FolderPath As String
    Dim PathSection As Variant
    For Each PathSection In PathSections
        FolderPath = BuildFolderPath(FolderPath, PathSection)
        
        If Not Tools.Fso.FolderExists(FolderPath) Then
            Tools.Fso.CreateFolder FolderPath
        End If
    Next PathSection

End Sub


Private Function BuildFolderPath(ByVal Path1 As String, ByVal Path2 As String) As String
    BuildFolderPath = Path1 & Path2 & "\"
End Function




