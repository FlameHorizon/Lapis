VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LongLongComparer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Collection")
Option Explicit

Implements IComparer

Public Function Compare(ByRef x As Variant, ByRef y As Variant) As Long
    Compare = IIf(x < y, -1, IIf(x = y, 0, 1))
End Function


Private Function IComparer_Compare(ByRef x As Variant, ByRef y As Variant) As Long
    IComparer_Compare = Me.Compare(x, y)
End Function

