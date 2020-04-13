Attribute VB_Name = "System"
'@Folder("Lapis")
Option Explicit

Public Const LongMinValue As Long = -2147483648#
Public Const LongMaxValue As Long = 2147483647#

Public Function IsNothing(ByVal Item As Variant) As Boolean

    If IsObject(Item) = False Then
        IsNothing = False
        Exit Function
    End If
    
    IsNothing = (Item Is Nothing)

End Function
