Attribute VB_Name = "Path"
Option Explicit
'@Folder("Lapis")


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
    HasIllegalCharacterInPath = (StringH.IndexOfAny(Path, IllegalCharactersInPath) > 0)
End Function


' Returns a value of indicating if given filename has any illegal characters
' deinfed in IllegalCharactersInFilename property.
Public Function HasIllegalCharacterInFilename(ByVal FileName As String) As Boolean
    HasIllegalCharacterInFilename = (StringH.IndexOfAny(FileName, IllegalCharactersInFilename) > 0)
End Function


