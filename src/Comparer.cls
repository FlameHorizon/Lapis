VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Comparer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'@Exposed
'@Folder("Lapis")

Implements IComparer

Private Const ClassName As String = "Comparer"


Public Function Compare(ByRef x As Variant, ByRef y As Variant) As Long
    
    If IsNull(x) Or IsEmpty(x) Then
        Compare = -1
        Exit Function
    End If
    
    If IsNull(y) Or IsEmpty(y) Then
        Compare = 1
        Exit Function
    End If
    
    If VarType(x) = vbString And VarType(y) = vbString Then
        Compare = VBA.StrComp(x, y, vbTextCompare)
    Else
        Exception.ArgumentException vbNullString, ClassName & ".Compare"
    End If

End Function


Private Function IComparer_Compare(ByRef x As Variant, ByRef y As Variant) As Long
    IComparer_Compare = Compare(x, y)
End Function


