Attribute VB_Name = "StringComparer"
'@Folder("Lapis")
Option Explicit

' Collection of static method which supplements Comaprer class.


' Returns a Defualt comparer which in this case is string comparer.
Public Property Get OrdinalIgnoreCase() As OrdinalIgnoreCaseStringComparer
    Set OrdinalIgnoreCase = New OrdinalIgnoreCaseStringComparer
End Property


