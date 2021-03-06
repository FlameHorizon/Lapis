VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StringEqualityComparer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Collection")
Option Explicit

Implements IEqualityComparer


Public Function Equals(ByVal x As String, ByVal y As String) As Boolean
    
    If (x = vbNullString And y = vbNullString) Then
        Equals = True
        
    ElseIf (x = vbNullString Or y = vbNullString) Then
        Equals = False
        
    ElseIf (StrComp(x, y, vbTextCompare)) = 0 Then
        Equals = True
        
    Else
        Equals = False
    End If
    
End Function


Private Function IEqualityComparer_Equals(ByVal x As Variant, ByVal y As Variant) As Boolean
    IEqualityComparer_Equals = Equals(x, y)
End Function


