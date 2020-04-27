VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WorkbookToStringConverter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Tests.TestClasses")
Option Explicit

Implements IToString


Public Function ToString(ByVal Obj As Workbook) As String
    ToString = Obj.Name
End Function


Private Function IToString_ToString(ByVal Obj As Variant) As String
    IToString_ToString = ToString(Obj)
End Function


