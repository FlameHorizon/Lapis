Attribute VB_Name = "Factory"
'@Folder("Lapis")
Option Explicit


Public Function GetStringEqualityComparer() As StringEqualityComparer
    Set GetStringEqualityComparer = New StringEqualityComparer
End Function


Public Function GetLongEqualityComparer() As LongEqualityComparer
    Set GetLongEqualityComparer = New LongEqualityComparer
End Function


Public Function GetLongComparer() As LongComparer
    Set GetLongComparer = New LongComparer
End Function


Public Function GetValueTypeToStringConverter() As ValueTypeToStringConverter
    Set GetValueTypeToStringConverter = New ValueTypeToStringConverter
End Function


