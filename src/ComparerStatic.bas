Attribute VB_Name = "ComparerStatic"
'@Folder("Lapis")
Option Explicit

' Collection of static method which supplements Comaprer class.


' Returns a Defualt comparer which in this case is string comparer.
Public Property Get Default() As Comparer
    Set Default = New Comparer
End Property


