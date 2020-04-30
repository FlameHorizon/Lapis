VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OrdinalIgnoreCaseStringEquality"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Collection")
Option Explicit

Implements IEqualityComparer

Private Const ClassName As String = "OrdinalIgnoreCaseStringEquality"


Public Function Equals(ByRef x As Variant, ByRef y As Variant) As Boolean
    
    If System.IsNothing(x) Or System.IsNothing(y) Then
        Equals = False
        Exit Function
    End If
    
    If VarType(x) <> VbVarType.vbString Or VarType(y) <> VbVarType.vbString Then
        Equals = False
        Exit Function
    End If
    
    ' Now, we are sure we are working with strings.
    If VBA.Len(x) <> VBA.Len(y) Then
        Equals = False
        Exit Function
    End If

    Equals = VBA.StrComp(x, y, VbCompareMethod.vbTextCompare) = 0

End Function


Private Function IEqualityComparer_Equals(ByVal x As Variant, ByVal y As Variant) As Boolean
    IEqualityComparer_Equals = Me.Equals(x, y)
End Function


