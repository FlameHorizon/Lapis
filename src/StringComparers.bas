Attribute VB_Name = "StringComparers"
'@Folder("Collection")
Option Explicit

' Collection of static method which supplements Comaprer class.


' Returns a Defualt comparer which in this case is string comparer.
Public Property Get OrdinalIgnoreCase() As OrdinalIgnoreCaseStringComparer
    Set OrdinalIgnoreCase = New OrdinalIgnoreCaseStringComparer
End Property


