Attribute VB_Name = "DirectoryExt"
Option Explicit
'@Folder("Helper")

Private Const ModuleName As String = "DirectoryExt"

Public Const DirectorySeparatorChar As String = "\"

' Specifies whether to search the current directory,
' or the current directory and all subdirectories.
Public Enum SearchOption
    
    TopDirectoryOnly
    AllDirectories

End Enum


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


' Splits a path to a file into individual folder segments.
' Example: "C:\abc1\abc1" will be splitted into array with elements "C:", "abc1" and "abc2".
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


' Removes all of files that meet specified criteria from the selected folder.
Public Sub RemoveFromFolder(ByVal Path As String, Optional ByRef SearchPattern As String = "*.*", _
                            Optional ByRef SearchOpt As SearchOption = TopDirectoryOnly)
    
    Const MethodName = "RemoveFromFolder"
    
    If Path = vbNullString Then
        Errors.OnArgumentError "Path", ModuleName & "." & MethodName
    End If
    
    If SearchPattern = vbNullString Then
        Errors.OnArgumentError "SearchPattern", ModuleName & "." & MethodName
    End If
    
    If Not Tools.Fso.FolderExists(Path) Then
        Errors.OnDirectoryNotFound "Path", ModuleName & "." & MethodName
    End If
    
    Dim FoldersToDelete As Collection ' <String>
    Set FoldersToDelete = GetFiles(Path, SearchPattern, SearchOpt)
    
    Dim FolderPath As Variant ' <String>
    For Each FolderPath In FoldersToDelete
        Tools.Fso.DeleteFolder FolderPath
    Next FolderPath
    
    Set FoldersToDelete = Nothing
 
End Sub


' Returns paths of files that met specified search pattern.
' Param:
' Path: The absolute path to the directory to search. This string is not case-sensitive.
'
' SearchPattern: The search string to match against the names of files in path. This parameter can
'                contain a combination of valid literal path and wildcard (* and ?) characters,
'                but it doesn't support regular expressions.
' SearchOpt: Defines if user wants to also search in all subdirectiories
Public Function GetFiles(ByVal Path As String, Optional ByVal SearchPattern As String = "*", _
                         Optional ByRef SearchOpt As SearchOption = TopDirectoryOnly) As Collection

    Const MethodName = "GetFiles"

    If Path = vbNullString Then
        Errors.OnArgumentError "Path", ModuleName & "." & MethodName
    End If

    If Not Tools.Fso.FolderExists(Path) Then
        Errors.OnDirectoryNotFound "Path", ModuleName & "." & MethodName
    End If

    If SearchPattern = vbNullString Then
        Errors.OnArgumentError "SearchPattern", ModuleName & "." & MethodName
    End If

    Dim Output As New Collection

    If SearchOpt = SearchOption.AllDirectories Then
        Set Output = GetFilesRecursive(Path, SearchPattern)
    Else
        Set Output = GetFilesTopDirectory(Path, SearchPattern)
    End If

    Set GetFiles = Output
    Set Output = Nothing

End Function


' Method returns a collection of file paths start starting from StartFrom folder.
' Param:
' StartFrom : Folder path from which search will began.
' SearchPattern : File name filter. In order to add file to the output
'                 file name has to match SearchPattern.
Private Function GetFilesRecursive(ByVal StartFrom As String, ByVal SearchPattern As String) As Collection
    
    Dim SubFld As Scripting.Folder
    Dim Fld As Scripting.Folder
    Dim FilePath As Variant '<String>
    Dim Output As New Collection
    Dim FolderStack As New Stack
    
    FolderStack.Init1 100
    FolderStack.Push StartFrom
    
    Do While FolderStack.Count > 0
        Set Fld = Tools.Fso.GetFolder(FolderStack.Pop)
        
        ' Push folders on top of the stack from current directory.
        For Each SubFld In Fld.SubFolders
            FolderStack.Push SubFld
        Next SubFld
        
        ' Add to the collection, every file from current
        ' directory which matches SearchPattern.
        For Each FilePath In GetFilesTopDirectory(Fld.Path, SearchPattern)
            Output.Add FilePath
        Next FilePath
        
    Loop
    
    Set GetFilesRecursive = Output
    Set Output = Nothing
    Set FolderStack = Nothing
    Set Fld = Nothing
    
End Function


' Returns a collection of file paths from the selected folder. File name has to match the SearchPattern.
Private Function GetFilesTopDirectory(ByVal Path As String, ByVal SearchPattern As String) As Collection '<String>
    
    Dim File As Scripting.File
    Dim Output As New Collection
    
    For Each File In Tools.Fso.GetFolder(Path).Files
        If File.Name Like SearchPattern Then
            Output.Add File.Path
        End If
    Next File

    Set GetFilesTopDirectory = Output
    Set Output = Nothing
    Set File = Nothing

End Function


Private Function RemoveTrailingBackslash(ByVal Path As String) As String
    
    If Path Like "*\" Then
        RemoveTrailingBackslash = VBA.Left$(Path, Len(Path) - 1)
    Else
        RemoveTrailingBackslash = Path
    End If
    
End Function


' Returns number of text lines based on the specified Stream.
' Throws an ArgumentNullException when Stream is set to nothing.
Public Function LinesCount(ByVal Stream As ADODB.Stream) As Long

    Const MethodName = "LinesCount"
    
    If Stream Is Nothing Then
        Errors.OnArgumentNull "Stream", ModuleName & "." & MethodName
    End If
    
    If IsStreamClosed(Stream) Then
        Errors.OnInvalidOperation "Stream", "Stream is closed. " & ModuleName & "." & MethodName
    End If
    
    Dim Ln As Long
    Stream.Position = 0
    Do While Stream.EOS <> True
        Stream.SkipLine
        Ln = Ln + 1
    Loop
    
    LinesCount = Ln
    
End Function


' Determines if TextStream is closed.
' There is no property of TextStream object (like object.Closed) to know directly if Stream is close
' or open. To know TextStream state, method attempt to read cursor position. If it fails
' (throws an error), that would mean Stream is not readable (closed).
' Throws an ArgumentNullException when Stream is set to nothing.
Public Function IsStreamClosed(ByRef Stream As ADODB.Stream) As Boolean
    
    Const MethodName = "IsStreamClosed"
    
    If Stream Is Nothing Then
        Errors.OnArgumentNull "Stream", ModuleName & "." & MethodName
    End If
    
    On Error Resume Next
    '@Ignore VariableNotUsed
    Dim Pos As Long: Pos = Stream.Position
    
    If Err.Number = 91 Then
        IsStreamClosed = True
        On Error GoTo 0
        
    ElseIf Err.Number = 0 Then
        IsStreamClosed = False
        
    Else
        ' Other, unexpected error occured. This error has to be moved upward.
        Errors.OnUnhandledError Err.Number & ", " & Err.Description & " " & ModuleName & "." & MethodName
    End If
    
End Function



