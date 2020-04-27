Attribute VB_Name = "StringEqualityComparers"
'@Folder("Collection")
Option Explicit

' Collection of static method which supplements IEqualityComaprer interface.


' Returns a Defualt comparer which in this case is string comparer.
Public Property Get OrdinalIgnoreCase() As OrdinalIgnoreCaseStringEquality
    Set OrdinalIgnoreCase = New OrdinalIgnoreCaseStringEquality
End Property
