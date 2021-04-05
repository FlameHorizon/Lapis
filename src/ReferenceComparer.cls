VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ReferenceComparer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Collection")
Option Explicit

Private Const ModuleName As String = "ReferenceComparer"

Implements Lapis.IComparer


Public Function Compare(ByRef x As Variant, ByRef y As Variant) As Long

    Const MethodName = "Compare"

    If VBA.IsObject(x) = False Or VBA.IsObject(y) = False Then
        Lapis.Errors.OnArgumentError vbNullString, _
                                     "Atleast of side of comparison is not an object. " _
                                     & ModuleName & "." & MethodName
    End If
    
    If x Is Nothing And y Is Nothing Then
        Compare = 0
        Exit Function
    End If
    
    If y Is Nothing Then
        Compare = 1
        Exit Function
    End If
    
    If x Is Nothing Then
        Compare = -1
        Exit Function
    End If

    If TypeOf x Is IComparable Then
        Compare = CompareUsingComparable(x, y)
    Else
        Lapis.Errors.OnArgumentError vbNullString, _
                                     "Default comparer wasn't found for x argument. " _
                                     & ModuleName & "." & MethodName
    End If

End Function


Private Function CompareUsingComparable(ByVal Value As IComparable, ByVal y As Variant) As Long
    CompareUsingComparable = Value.CompareTo(y)
End Function


Private Function IComparer_Compare(ByRef x As Variant, ByRef y As Variant) As Long
    IComparer_Compare = Me.Compare(x, y)
End Function



