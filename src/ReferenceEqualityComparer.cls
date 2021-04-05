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
    Else
        Equals = x Is y
    End If
    
End Function


Private Function IEqualityComparer_Equals(ByVal x As Variant, ByVal y As Variant) As Boolean
    IEqualityComparer_Equals = Me.Equals(x, y)
End Function

