VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LongComparer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Collection")
Option Explicit

Implements IComparer


' Compares two objects together.
' Method should return value less than zero if x is less than y.
' Zero if x is equal to y, or a value greater than zero if z is greater than
Public Function Compare(ByVal x As Variant, ByVal y As Variant) As Long
    
    ' For this comparer, nothing values will be acceptable and they
    ' will be sorted lower than any other value.
    If System.IsNothing(x) And System.IsNothing(y) Then
        Compare = 0
        Exit Function
    End If
    
    If System.IsNothing(x) Then
        Compare = -1
        Exit Function
    End If
    
    If System.IsNothing(y) Then
        Compare = 1
        Exit Function
    End If

    Compare = IIf(x < y, -1, IIf(x = y, 0, 1))
    
End Function


Private Function IComparer_Compare(ByRef x As Variant, ByRef y As Variant) As Long
    IComparer_Compare = Compare(x, y)
End Function


