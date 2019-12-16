Attribute VB_Name = "Factory"
'@Folder("Lapis")
Option Explicit


Public Function GetStringEqualityComparer() As StringEqualityComparer
    Set GetStringEqualityComparer = New StringEqualityComparer
End Function


Public Function GetLongEqualityComparer() As LongEqualityComparer
    Set GetLongEqualityComparer = New LongEqualityComparer
End Function


