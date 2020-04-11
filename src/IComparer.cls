VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IComparer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'@Exposed
'@Folder("Collection")
Option Explicit


' Compares two objects together.
' Method should return value less than zero if x is less than y.
' Zero if x is equal to y, or a value greater than zero if z is greater than y.
Public Function Compare(ByRef x As Variant, ByRef y As Variant) As Long
End Function


