VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ValueTypeToStringConverter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IToString


Public Function ToString(ByVal Obj As Variant) As String
    ToString = VBA.CStr(Obj)
End Function


Private Function IToString_ToString(ByVal Obj As Variant) As String
    IToString_ToString = ToString(Obj)
End Function


