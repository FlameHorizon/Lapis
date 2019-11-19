VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LongEqualityComparer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'@Folder("Lapis")

Implements IVBAEqualityComparer


Public Function Equals(ByVal Obj1 As Variant, ByVal Obj2 As Variant) As Boolean
    
    If (Not IsNumeric(Obj1) Or Not IsNumeric(Obj2)) Then
        Equals = False
        Exit Function
    End If
    
    Equals = (Obj1 = Obj2)
    
End Function


Private Function IVBAEqualityComparer_Equals(ByVal Obj1 As Variant, ByVal Obj2 As Variant) As Boolean
    IVBAEqualityComparer_Equals = Equals(Obj1, Obj2)
End Function


