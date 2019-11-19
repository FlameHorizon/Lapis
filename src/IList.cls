VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'@Exposed
Option Explicit
'@Folder("Collection")


Public Function GetItem(ByVal Index As Long) As Variant
End Function


Public Sub SetItem(ByVal Index As Long, ByVal Value As Variant)
End Sub


Public Function Add(ByVal Value As Variant) As Long
End Function


Public Function Contains(ByRef Value As Variant) As Boolean
End Function


Public Sub Clear()
End Sub


Public Property Get IsReadOnly() As Boolean
End Property


Public Property Get IsFixedSize() As Boolean
End Property


Public Function IndexOf(ByRef Value As Variant) As Long
End Function


Public Sub Insert(ByVal Index As Long, ByVal Value As Variant)
End Sub


Public Sub Remove(ByVal Value As Variant)
End Sub


Public Sub RemoveAt(ByVal Index As Long)
End Sub


Public Property Get Count() As Long
End Property


Public Sub CopyTo(ByRef Arr() As Variant, ByVal Index As Long)
End Sub


