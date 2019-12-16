Attribute VB_Name = "Path"
'@Folder("Lapis")
Option Explicit


' List of illegal characters which can't be used in the path.
Public Property Get IllegalCharactersInPath() As String()
    IllegalCharactersInPath = Split("< > |")
End Property


' List of illegal characters which can't be used in the filename.
Public Property Get IllegalCharactersInFilename() As String()
    IllegalCharactersInFilename = Split("~ # % & * : < > ? { | } / \ [ ]")
End Property


' Returns a value of indicating if given path has any illegal characters
' defined in IllegalCharactersInPath property.
Public Function HasIllegalCharacterInPath(ByVal Path As String) As Boolean
    HasIllegalCharacterInPath = (StringExt.IndexOfAny(Path, IllegalCharactersInPath) > 0)
End Function


' Returns a value of indicating if given filename has any illegal characters
' deinfed in IllegalCharactersInFilename property.
Public Function HasIllegalCharacterInFilename(ByVal FileName As String) As Boolean
    HasIllegalCharacterInFilename = (StringExt.IndexOfAny(FileName, IllegalCharactersInFilename) > 0)
End Function


