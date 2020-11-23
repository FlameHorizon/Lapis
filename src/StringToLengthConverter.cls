VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StringToLengthConverter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Lapis")
Option Explicit

Implements Lapis.IConverter


Public Function Convert(ByVal Item As Variant) As Variant
    Convert = VBA.Len(Item)
End Function


Private Function IConverter_Convert(ByVal Item As Variant) As Variant
    IConverter_Convert = Me.Convert(Item)
End Function
