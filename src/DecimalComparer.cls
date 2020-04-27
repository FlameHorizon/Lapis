VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DecimalComparer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IComparer


' Compares two objects together.
' Method should return value less than zero if x is less than y.
' Zero if x is equal to y, or a value greater than zero if z is greater than y.
Public Function Compare(ByRef x As Variant, ByRef y As Variant) As Long

    If VBA.Round(x, 15) < VBA.Round(y, 15) Then
        Compare = -1
    
    ElseIf VBA.Round(x, 15) = VBA.Round(y, 15) Then
        Compare = 0
        
    Else
        Compare = 1
        
    End If

End Function


Private Function IComparer_Compare(ByRef x As Variant, ByRef y As Variant) As Long
    IComparer_Compare = Me.Compare(x, y)
End Function



