Attribute VB_Name = "PredicateTests"
'@Folder("Tests")
Option Explicit

Private Const ModuleName As String = "PredicateTests"


Public Sub Start()

    EvalUsesStringComparerWhenComparerIsNotDefinedTest
    EvalReturnsValueWhenOperatorIsGreaterThanTest
    EvalReturnsValueWhenOperatorIsEqualTest
    EvalReturnsValueWhenOperatorIsLessThenTest
    EvalReturnsValueWhenOperatorIsGreaterThanOrEqualToTest
    EvalReturnsValueWhenOperatorIsLessThanOrEqualToTest
    EvalReturnsValueWhenOperatorIsNotEqualToTest

End Sub


Private Sub EvalUsesStringComparerWhenComparerIsNotDefinedTest()

    On Error GoTo ErrHandler
    Const MethodName = "EvalUsesStringComparerWhenComparerIsNotDefinedTest"

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


Private Sub EvalReturnsValueWhenOperatorIsGreaterThanTest()

    On Error GoTo ErrHandler
    Const MethodName = "EvalReturnsValueWhenOperatorIsGreaterThanTest"

    ' Arrange
    Dim Predicate As New Predicate
    Predicate.Operator = ComparisonOperator.GreaterThan
    Predicate.ComparisonValue = 0 ' RHS
    Predicate.Comparer = New LongComparer
    
    ' Act
    Dim Actual As Boolean
    Actual = Predicate.Eval(1) ' LHS
    
    ' Assert
    ExUnit.IsTrue Actual, GetSig(MethodName)

    Exit Sub
ErrHandler:
    Lapis.ExUnit.TestFailRunTime GetSig(MethodName)

End Sub


Private Sub EvalReturnsValueWhenOperatorIsEqualTest()

    On Error GoTo ErrHandler
    Const MethodName = "EvalReturnsValueWhenOperatorIsEqualTest"

    ' Arrange
    Dim Predicate As New Predicate
    Predicate.ComparisonValue = 1
    Predicate.Operator = ComparisonOperator.EqualTo
    Predicate.Comparer = New LongComparer
    
    ' Act
    Dim Actual As Boolean
    Actual = Predicate.Eval(1)
    
    ' Assert
    ExUnit.IsTrue Actual, GetSig(MethodName)
    
    Exit Sub
ErrHandler:
    Lapis.ExUnit.TestFailRunTime GetSig(MethodName)

End Sub


Private Sub EvalReturnsValueWhenOperatorIsLessThenTest()

    On Error GoTo ErrHandler
    Const MethodName = "EvalReturnsValueWhenOperatorIsLessThenTest"

    ' Arrange
    Dim Predicate As New Predicate
    Predicate.ComparisonValue = 1
    Predicate.Operator = ComparisonOperator.LessThan
    Predicate.Comparer = New LongComparer
    
    ' Act
    Dim Actual As Boolean
    Actual = Predicate.Eval(0)
    
    ' Assert
    ExUnit.IsTrue Actual, GetSig(MethodName)

    Exit Sub
ErrHandler:
    Lapis.ExUnit.TestFailRunTime GetSig(MethodName)

End Sub


Private Sub EvalReturnsValueWhenOperatorIsGreaterThanOrEqualToTest()

    On Error GoTo ErrHandler
    Const MethodName = "EvalReturnsValueWhenOperatorIsGreaterThanOrEqualToTest"
    
    ' Arrange
    Dim Predicate As New Predicate
    Predicate.ComparisonValue = 1
    Predicate.Operator = ComparisonOperator.GreaterThanOrEqualTo
    Predicate.Comparer = New LongComparer

    ' Act & Assert
    ExUnit.IsFalse Predicate.Eval(0), GetSig(MethodName)
    ExUnit.IsTrue Predicate.Eval(1), GetSig(MethodName)
    ExUnit.IsTrue Predicate.Eval(2), GetSig(MethodName)

    Exit Sub
ErrHandler:
    Lapis.ExUnit.TestFailRunTime GetSig(MethodName)

End Sub


Private Sub EvalReturnsValueWhenOperatorIsLessThanOrEqualToTest()

    On Error GoTo ErrHandler
    Const MethodName = "EvalReturnsValueWhenOperatorIsLessThanOrEqualToTest"

    ' Arrange
    Dim Predicate As New Predicate
    Predicate.ComparisonValue = 1
    Predicate.Operator = ComparisonOperator.LessThanOrEqualTo
    Predicate.Comparer = New LongComparer

    ' Act & Assert
    ExUnit.IsTrue Predicate.Eval(-1), GetSig(MethodName)
    ExUnit.IsTrue Predicate.Eval(0), GetSig(MethodName)

    Exit Sub
ErrHandler:
    Lapis.ExUnit.TestFailRunTime GetSig(MethodName)

End Sub


Private Sub EvalReturnsValueWhenOperatorIsNotEqualToTest()

    On Error GoTo ErrHandler
    Const MethodName = "EvalReturnsValueWhenOperatorIsNotEqualToTest"

    ' Arrange
    Dim Predicate As New Predicate
    Predicate.ComparisonValue = 1
    Predicate.Operator = ComparisonOperator.NotEqualTo
    Predicate.Comparer = New LongComparer

    ' Act & Assert
    ExUnit.IsTrue Predicate.Eval(0), GetSig(MethodName)
    ExUnit.IsFalse Predicate.Eval(1), GetSig(MethodName)
    ExUnit.IsTrue Predicate.Eval(2), GetSig(MethodName)

    Exit Sub
ErrHandler:
    Lapis.ExUnit.TestFailRunTime GetSig(MethodName)

End Sub
