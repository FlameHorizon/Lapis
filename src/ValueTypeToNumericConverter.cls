VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ValueTypeToNumericConverter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Lapis")
Option Explicit

Implements Lapis.IConverter

Private pPropertyName As String


Public Function Convert(ByVal Item As Variant) As Variant
    Convert = CDec(Item)
End Function


Private Function IConverter_Convert(ByVal Item As Variant) As Variant
    IConverter_Convert = Me.Convert(Item)
End Function


Public Property Get PropertyName() As String
    PropertyName = pPropertyName
End Property


Public Property Let PropertyName(ByVal Value As String)
    pPropertyName = Value
End Property
