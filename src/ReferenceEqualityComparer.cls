VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ReferenceEqualityComparer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Collection")
Option Explicit

Implements Lapis.IEqualityComparer


Public Function Equals(ByVal x As Variant, ByVal y As Variant) As Boolean
    
    If VBA.IsObject(x) = False Or VBA.IsObject(y) = False Then
        Equals = False
        Exit Function
    End If
    
    If x Is Nothing And y Is Nothing Then
        Equals = True
        Exit Function
    End If
    
    If TypeOf x Is IEquatable Then
        Equals = CompareUsingEquatable(x, y)
        Exit Function
    End If
    
    Equals = x Is y
    
End Function


Private Function CompareUsingEquatable(ByVal Value As IEquatable, ByVal y As Variant) As Long
    CompareUsingEquatable = Value.Equals(y)
End Function


Private Function IEqualityComparer_Equals(ByVal x As Variant, ByVal y As Variant) As Boolean
    IEqualityComparer_Equals = Me.Equals(x, y)
End Function

