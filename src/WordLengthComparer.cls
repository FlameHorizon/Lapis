VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WordLengthComparer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Exposed
'@Folder("Collection")
Option Explicit

Implements IComparer

Private Const ClassName As String = "WordLengthComparer"


Public Function Compare(ByRef x As Variant, ByRef y As Variant) As Long
    
    If IsNull(x) Or IsEmpty(x) Then
        Compare = -1
        Exit Function
    End If
    
    If IsNull(y) Or IsEmpty(y) Then
        Compare = 1
        Exit Function
    End If
    
    If VarType(x) = vbString And IsNumeric(y) Then
        If Len(x) > y Then
            Compare = 1
        ElseIf Len(x) = y Then
            Compare = 0
        Else
            Compare = -1
        End If
    Else
        Errors.OnArgumentError vbNullString, ClassName & ".Compare"
    End If

End Function


Private Function IComparer_Compare(ByRef x As Variant, ByRef y As Variant) As Long
    IComparer_Compare = Compare(x, y)
End Function

