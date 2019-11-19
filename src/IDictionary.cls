VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IDictionary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'@Folder("Collection")
'@Exposed


Public Function GetItem(ByVal Key As Variant) As Variant
End Function


Public Sub SetItem(ByVal Key As Variant, ByVal Item As Variant)
End Sub


Public Property Get Keys() As IVBACollection
End Property


Public Property Get Values() As IVBACollection
End Property


Public Function Contains(ByRef Key As Variant) As Boolean
End Function


Public Sub Add(ByRef Key As Variant, ByVal Value As Variant)
End Sub


Public Sub Clear()
End Sub


Public Property Get IsReadOnly() As Boolean
End Property


Public Property Get IsFixedSize() As Boolean
End Property


Public Sub Remove(ByRef Key As Variant)
End Sub


