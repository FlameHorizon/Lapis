Attribute VB_Name = "ComparersTests"
'@Folder("Tests")
Option Explicit

Private Const ModuleName As String = "ComparersTest"


Public Sub Start()
    
    LongComparerReturnsOneWhenNumberIsComparedToNothingTest

End Sub


Private Sub LongComparerReturnsOneWhenNumberIsComparedToNothingTest()

    On Error GoTo ErrHandler
    Const MethodName = "LongComparerReturnsOneWhenNumberIsComparedToNothingTest"

    ' Arrange
    Dim Comparer As New LongComparer
    
    ' Act
    Dim Actual As Long
    Actual = Comparer.Compare(1, Nothing)
    
    ' Assert
    ExUnit.AreEqual 1, Actual, GetSig(MethodName)

    Exit Sub
ErrHandler:
    Lapis.ExUnit.TestFailRunTime GetSig(MethodName)

End Sub


Private Function GetSig(ByVal MethodName As String) As String
    GetSig = ModuleName & "." & MethodName
End Function
