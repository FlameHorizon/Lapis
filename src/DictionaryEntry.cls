VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DictionaryEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'@Folder("Collection")

Private pKey As Variant
Private pValue As Variant


Public Sub Init1(ByVal Key As Variant, ByVal Value As Variant)

    pKey = Key
    pValue = Value

End Sub


Public Property Get EntryValue() As Variant
    EntryValue = pValue
End Property


Public Property Let EntryValue(ByVal Value As Variant)
    pValue = Value
End Property


Public Property Get Key() As Variant
    Key = pKey
End Property


Public Property Let Key(ByVal Value As Variant)
    pKey = Value
End Property


