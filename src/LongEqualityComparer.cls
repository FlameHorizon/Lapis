VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LongEqualityComparer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Lapis")
Option Explicit

Implements IVBAEqualityComparer


Public Function Equals(ByVal x As Variant, ByVal y As Variant) As Boolean
    
    If (Not IsNumeric(x) Or Not IsNumeric(y)) Then
        Equals = False
        Exit Function
    End If
    
    Equals = (x = y)
    
End Function


Private Function IVBAEqualityComparer_Equals(ByVal x As Variant, ByVal y As Variant) As Boolean
    IVBAEqualityComparer_Equals = Equals(x, y)
End Function
