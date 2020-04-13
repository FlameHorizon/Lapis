Attribute VB_Name = "PredicateTests"
'@Folder("Tests")
Option Explicit

Private Const ModuleName As String = "PredicateTests"


Public Sub Start()

    PredicateUsesStringComparerWhenComparerIsNotDefinedTest

End Sub


Private Sub PredicateUsesStringComparerWhenComparerIsNotDefinedTest()

    On Error GoTo ErrHandler
    Const MethodName = "PredicateUsesStringComparerWhenComparerIsNotDefinedTest"

    ' Arrange
    Dim Predicate As New Predicate
    Predicate.Operator = ComparisonOperator.EqualTo
    Predicate.ComparisonValue = "a"
    
    ' Act
    Dim Actual As Boolean
    Actual = Predicate.Eval("b")
    
    ' Assert
    ExUnit.IsFalse Actual, GetSig(MethodName)

    Exit Sub
ErrHandler:
    Lapis.ExUnit.TestFailRunTime GetSig(MethodName)

End Sub


Private Function GetSig(ByVal MethodName As String) As String
    GetSig = ModuleName & "." & MethodName
End Function

