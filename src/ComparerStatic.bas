Attribute VB_Name = "ComparerStatic"
Option Explicit
'@Folder("Lapis")

' Collection of static method which supplements Comaprer class.


' Returns a Defualt comparer which in this case is string comparer.
Public Property Get Default() As Comparer
    Set Default = New Comparer
End Property


