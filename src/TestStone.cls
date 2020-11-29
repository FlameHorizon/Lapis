VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TestStone"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Tests.TestClasses")
Option Explicit

Implements IComparable

Public Weight As Single
Public Age As Single


Public Function CompareTo(ByVal x As TestStone) As Long

    If IsObject(x) = False Then
        Lapis.Errors.OnArgumentError "x", "TestStone.CompareTo"
    End If
    
    If x Is Nothing Then
        CompareTo = 1
        Exit Function
    End If
    
    Dim Value As TestStone
    Set Value = x
    
    If Me.Age = Value.Age And Me.Weight = Value.Weight Then
        CompareTo = 0
    
    ElseIf Me.Weight < Value.Weight Then
        CompareTo = -1
        
    Else
        CompareTo = 1
    End If

End Function


Private Function IComparable_CompareTo(ByRef x As Variant) As Long
    IComparable_CompareTo = Me.CompareTo(x)
End Function


