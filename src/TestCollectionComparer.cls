VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TestCollectionComparer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Tests.TestClasses")
Option Explicit

' This test class should be used only in the testing environment.
' Do never use this class in production code!

Implements IComparer


' Compares two instances of collection using their sizes.
Public Function Compare(ByVal x As Collection, ByVal y As Collection) As Long
    
    If x Is Nothing And y Is Nothing Then
        Compare = 0
        Exit Function
    End If
    
    If x Is Nothing Then
        Compare = -1
        Exit Function
    End If
    
    If y Is Nothing Then
        Compare = 1
        Exit Function
    End If

    Compare = IIf(x.Count < y.Count, -1, IIf(x.Count = y.Count, 0, 1))
    
End Function


Private Function IComparer_Compare(ByRef x As Variant, ByRef y As Variant) As Long
    IComparer_Compare = Compare(x, y)
End Function


