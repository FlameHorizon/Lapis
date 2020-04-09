VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PropertyToStringConverter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IToString


Private pPropertyName As String


Public Property Get PropertyName() As String
    PropertyName = pPropertyName
End Property


Public Property Let PropertyName(ByVal Value As String)
    pPropertyName = Value
End Property


Public Function ToString(ByVal Obj As Variant) As String
    ToString = VBA.CallByName(Obj, PropertyName, VbCallType.VbGet)
End Function


Private Function IToString_ToString(ByVal Obj As Variant) As String
    IToString_ToString = ToString(Obj)
End Function
