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
Public Function Compare(ByVal x As Long, ByVal y As Long) As Long
    Compare = IIf(x < y, -1, IIf(x = y, 0, 1))
End Function


Private Function IComparer_Compare(ByRef x As Variant, ByRef y As Variant) As Long
    IComparer_Compare = Compare(x, y)
End Function

