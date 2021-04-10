Attribute VB_Name = "Collection2Tests"
'@Folder("Tests")
Option Explicit

Private Const ModuleName As String = "Collection2Tests"


Public Sub Start()

    AllReturnsFalseWhenSourceDoesNotSatisfyConditionTest
    AllReturnsTrueWhenSourceDoesSatisfyConditionTest
    AllReturnsArgumentNullErrorWhenSourceIsNothingTest
    AllReturnsArgumentNullErrorWhenPreidcateIsNothingTest
    AllReturnsTrueWhenSourceIsEmptyTest
    AllReturnsFalseWhenSourceSatisfyConditionAndHasNothingTest
    AllReturnsFalseWhenSourceObjectsSatisfyConditionAndHasNothingTest
    
    WhereTest
    WhereReturnsArgumentNullErrorWhenSourceIsNothingTest
    WhereReturnsArgumentNullErrorWhenPredicateIsNothingTest
    WhereReturnsOutputWhenFilteringObjectsTest
    WhereReturnsArgumentNullErrorWhenElementOfSourceIsNothingTest
    
    SomeReturnsTrueWhenPredicateIsNothingAndSourceIsNotEmptyTest
    SomeReturnsFalseWhenPredicateIsNothingAndSourceIsEmptyTest
    SomeReturnsTrueWhenAtleastOneItemSatisfyConditionTest
    SomeReturnsFalseWhenNoItemSatisfyConditionTest
    SomeReturnsFalseWhenSourceIsEmptyTest
    SomeReturnsArgumentNullErrorWhenSourceIsNothingTest
    
    SumReturnsValueWhenSourceContainsIntegersTest
    SumReturnsValueWhenSourceContainsReferencedTypesTest
    SumReturnsValueWhenSourceIsEmptyTest
    SumReturnsValueWhenSourceContainsIntegersAndNothingTest
    SumReturnsArgumentNullErrorWhenSourceIsNothingTest
    
    AverageReturnsValueWhenSourceContainsIntegersTest
    AverageReturnsArgumentNullErrorWhenSourceIsEmptyTest
    AverageReturnsValueWhenSourceIsEmptyTest
    AverageReturnsValueWhenSourceContainsIntegersAndNothingTest
    AverageReturnsValueWhenSourceContainsOnlyNothingTest
    AverageReturnsValueWhenSourceContainsReferecenTypesTest
    
    CountTest
    CountReturnsArgumentNullErrorWhenSourceIsNothingTest
    
    ConvertTest
    ConvertReturnsArgumentNullWhenSourceIsNothingTest
    ConvertReturnsArgumentNullWhenSelectorIsNothingTest
    
    FirstReturnsValueWhenSourceDoesSatisfyConditionTest
    FirstReturnsArgumentNullErrorWhenSourceIsNothingTest
    FirstReturnsInvalidOperationErrorWhenSourceDoesNotSatisfyConditionTest
    FirstReturnsInvalidOperationErrorWhenSourceIsEmptyTest
    
    LastReturnsValueWhenSourceDoesSatisfyConditionTest
    LastReturnsArgumentNullErrorWhenSourceIsNothingTest
    LastReturnsInvalidOperationErrorWhenSourceDoesNotSatisfyConditionTest
    LastReturnsInvalidOperationErrorWhenSourceIsEmptyTest
    
End Sub


Private Function GetSig(ByVal MethodName As String) As String
    GetSig = ModuleName & "." & MethodName
End Function


Private Sub AllReturnsFalseWhenSourceDoesNotSatisfyConditionTest()

    On Error GoTo ErrHandler
    Const MethodName = "AllReturnsFalseWhenSourceDoesNotSatisfyConditionTest"

    ' Arrange
    Dim Expr As ICallable
    Set Expr = Lambda.Create("$1 = 1")
    
    ' Act
    Dim Actual As Boolean
    Actual = CollectionExt2.All(CollectionExt.Make(1, 2, 3), Expr)
    
    ' Assert
    ExUnit.IsFalse Actual, GetSig(MethodName)

    Exit Sub
ErrHandler:
    Lapis.ExUnit.TestFailRunTime GetSig(MethodName)

End Sub


Private Sub AllReturnsTrueWhenSourceDoesSatisfyConditionTest()

    On Error GoTo ErrHandler
    Const MethodName = "AllReturnsTrueWhenSourceDoesSatisfyConditionTest"

    ' Arrange
    Dim Expr As ICallable
    Set Expr = Lambda.Create("$1 = 1")
    
    ' Act
    Dim Actual As Boolean
    Actual = CollectionExt2.All(CollectionExt.Make(1, 1, 1), Expr)
    
    ' Assert
    ExUnit.IsTrue Actual, GetSig(MethodName)

    Exit Sub
ErrHandler:
    Lapis.ExUnit.TestFailRunTime GetSig(MethodName)

End Sub


Private Sub AllReturnsArgumentNullErrorWhenSourceIsNothingTest()

    On Error GoTo ErrHandler
    Const ExpectedError As Long = ErrorCode.ArgumentNull
    Const MethodName = "AllReturnsArgumentNullErrorWhenSourceIsNothingTest"
    
    ' Act
    CollectionExt2.All Nothing, Lambda.Create("$1 = 1")
    
    ' Assert
ErrHandler:
    Lapis.ExUnit.IsException ExpectedError, Err.Number, GetSig(MethodName)

End Sub


Private Sub AllReturnsArgumentNullErrorWhenPreidcateIsNothingTest()

    On Error GoTo ErrHandler
    Const ExpectedError As Long = ErrorCode.ArgumentNull
    Const MethodName = "AllReturnsArgumentNullErrorWhenPreidcateIsNothingTest"

    ' Act
    CollectionExt2.All New Collection, Nothing
    
    ' Assert
ErrHandler:
    Lapis.ExUnit.IsException ExpectedError, Err.Number, GetSig(MethodName)

End Sub


Private Sub AllReturnsTrueWhenSourceIsEmptyTest()

    On Error GoTo ErrHandler
    Const MethodName = "AllReturnsTrueWhenSourceIsEmptyTest"

    ' Arrange
    Dim Expr As ICallable
    Set Expr = Lambda.Create("$1 = 1")
    
    ' Act
    Dim Actual As Boolean
    Actual = CollectionExt2.All(New Collection, Expr)
    
    ' Assert
    ExUnit.IsTrue Actual, GetSig(MethodName)

    Exit Sub
ErrHandler:
    Lapis.ExUnit.TestFailRunTime GetSig(MethodName)

End Sub


Private Sub AllReturnsFalseWhenSourceSatisfyConditionAndHasNothingTest()

    On Error GoTo ErrHandler
    Const MethodName = "AllReturnsFalseWhenSourceSatisfyConditionAndHasNothingTest"

    ' Arrange
    Dim Expr As ICallable
    Set Expr = Lambda.Create("$1 = 1")
    
    ' Act
    Dim Actual As Boolean
    Actual = CollectionExt2.All(CollectionExt.Make(1, Nothing, 1), Expr)
    
    ' Assert
    ExUnit.IsFalse Actual, GetSig(MethodName)

    Exit Sub
ErrHandler:
    Lapis.ExUnit.TestFailRunTime GetSig(MethodName)

End Sub


Private Sub AllReturnsFalseWhenSourceObjectsSatisfyConditionAndHasNothingTest()

    On Error GoTo ErrHandler
    Const MethodName = "AllReturnsFalseWhenSourceObjectsSatisfyConditionAndHasNothingTest"

    ' Act
    Dim Actual As Boolean
    Actual = CollectionExt2.All(CollectionExt.Make(New TestStone, Nothing), _
                                Lambda.Create("$1.Age = 0 and $1.Weight = 0"))
    
    ' Assert
    ExUnit.IsFalse Actual, GetSig(MethodName)

    Exit Sub
ErrHandler:
    ExUnit.TestFailRunTime GetSig(MethodName)

End Sub


Private Sub WhereTest()

    On Error GoTo ErrHandler
    Const MethodName = "WhereTest"

    ' Act
    Dim Actual As Collection
    Set Actual = CollectionExt2.Where(CollectionExt.Make(1, 2, 3), Lambda.Create("$1 = 2"))
    
    ' Assert
    ExUnit.AreEqual 1, Actual.Count, GetSig(MethodName)

    Exit Sub
ErrHandler:
    Lapis.ExUnit.TestFailRunTime GetSig(MethodName)

End Sub


Private Sub WhereReturnsArgumentNullErrorWhenSourceIsNothingTest()

    On Error GoTo ErrHandler
    Const ExpectedError As Long = ErrorCode.ArgumentNull
    Const MethodName = "WhereReturnsArgumentNullErrorWhenSourceIsNothingTest"

    ' Act
    CollectionExt2.Where Nothing, Lambda.Create("$1 = 2")

    ' Assert
ErrHandler:
    Lapis.ExUnit.IsException ExpectedError, Err.Number, GetSig(MethodName)

End Sub


Private Sub WhereReturnsArgumentNullErrorWhenPredicateIsNothingTest()

    On Error GoTo ErrHandler
    Const ExpectedError As Long = ErrorCode.ArgumentNull
    Const MethodName = "WhereReturnsArgumentNullErrorWhenPredicateIsNothingTest"

    ' Act
    CollectionExt2.Where New Collection, Nothing
    
    ' Assert
ErrHandler:
    Lapis.ExUnit.IsException ExpectedError, Err.Number, GetSig(MethodName)

End Sub


Private Sub WhereReturnsOutputWhenFilteringObjectsTest()

    On Error GoTo ErrHandler
    Const MethodName = "WhereReturnsOutputWhenFilteringObjectsTest"

    ' Arrange
    Dim Stn1 As New TestStone
    Stn1.Age = 110
    Stn1.Weight = 100
    
    Dim Stn2 As New TestStone
    Stn2.Age = 105
    Stn2.Weight = 200
    
    ' Act
    Dim Actual As Collection
    Set Actual = CollectionExt2.Where(CollectionExt.Make(Stn1, Stn2), _
                                      Lambda.Create("$1.Age > 100 and $1.Weight >= 200"))
    
    ' Assert
    ExUnit.AreEqual 1, Actual.Count, GetSig(MethodName)

    Exit Sub
ErrHandler:
    ExUnit.TestFailRunTime GetSig(MethodName)

End Sub


Private Sub WhereReturnsArgumentNullErrorWhenElementOfSourceIsNothingTest()

    On Error GoTo ErrHandler
    Const ExpectedError As Long = ErrorCode.ArgumentNull
    Const MethodName = "WhereReturnsArgumentNullErrorWhenElementOfSourceIsNothingTest"

    ' Arrange
    Dim Stn1 As New TestStone
    
    ' Act
    Dim Actual As Collection
    Set Actual = CollectionExt2.Where(CollectionExt.Make(Stn1, Nothing), _
                                      Lambda.Create("$1.Age > 100 and $1.Weight >= 200"))
    
    ' Assert
    ExUnit.AreEqual 1, Actual.Count, GetSig(MethodName)

ErrHandler:
    ExUnit.IsException ExpectedError, Err.Number, GetSig(MethodName)

End Sub


Private Sub SomeReturnsTrueWhenPredicateIsNothingAndSourceIsNotEmptyTest()

    On Error GoTo ErrHandler
    Const MethodName = "SomeReturnsTrueWhenPredicateIsNothingAndSourceIsNotEmptyTest"

    ' Act & Assert
    ExUnit.IsTrue CollectionExt2.Some(CollectionExt.Make(1, 2, 3)), GetSig(MethodName)

    Exit Sub
ErrHandler:
    Lapis.ExUnit.TestFailRunTime GetSig(MethodName)

End Sub


Private Sub SomeReturnsFalseWhenPredicateIsNothingAndSourceIsEmptyTest()

    On Error GoTo ErrHandler
    Const MethodName = "SomeReturnsFalseWhenPredicateIsNothingAndSourceIsEmptyTest"

    ' Act & Assert
    ExUnit.IsFalse CollectionExt2.Some(New Collection), GetSig(MethodName)

    Exit Sub
ErrHandler:
    Lapis.ExUnit.TestFailRunTime GetSig(MethodName)

End Sub


Private Sub SomeReturnsTrueWhenAtleastOneItemSatisfyConditionTest()

    On Error GoTo ErrHandler
    Const MethodName = "SomeReturnsTrueWhenAtleastOneItemSatisfyConditionTest"
    
    ' Act
    Dim Actual As Boolean
    Actual = CollectionExt2.Some(CollectionExt.Make(1, 2, 3), Lambda.Create("$1 = 2"))

    ' Assert
    ExUnit.IsTrue Actual, GetSig(MethodName)

    Exit Sub
ErrHandler:
    Lapis.ExUnit.TestFailRunTime GetSig(MethodName)

End Sub


Private Sub SomeReturnsFalseWhenNoItemSatisfyConditionTest()

    On Error GoTo ErrHandler
    Const MethodName = "SomeReturnsFalseWhenNoItemSatisfyConditionTest"

    ' Act
    Dim Actual As Boolean
    Actual = CollectionExt2.Some(CollectionExt.Make(1, 2, 3), Lambda.Create("$1 = 0"))

    ' Assert
    ExUnit.IsFalse Actual, GetSig(MethodName)

    Exit Sub
ErrHandler:
    Lapis.ExUnit.TestFailRunTime GetSig(MethodName)

End Sub


Private Sub SomeReturnsFalseWhenSourceIsEmptyTest()

    On Error GoTo ErrHandler
    Const MethodName = "SomeReturnsFalseWhenSourceIsEmptyTest"

    ' Act & Assert
    ExUnit.IsFalse CollectionExt2.Some(New Collection, Lambda.Create("$1 = 0")), GetSig(MethodName)

    Exit Sub
ErrHandler:
    Lapis.ExUnit.TestFailRunTime GetSig(MethodName)

End Sub


Private Sub SomeReturnsArgumentNullErrorWhenSourceIsNothingTest()

    On Error GoTo ErrHandler
    Const ExpectedError As Long = ErrorCode.ArgumentNull
    Const MethodName = "SomeReturnsArgumentNullErrorWhenSourceIsNothingTest"

    ' Act
    CollectionExt2.Some Nothing

    ' Assert
ErrHandler:
    Lapis.ExUnit.IsException ExpectedError, Err.Number, GetSig(MethodName)

End Sub


Private Sub SumReturnsValueWhenSourceContainsIntegersTest()

    On Error GoTo ErrHandler
    Const MethodName = "SumReturnsValueWhenSourceContainsIntegersTest"
    
    ' Act
    Dim Actual As Variant
    Actual = CollectionExt2.Sum(CollectionExt.Make(1, 2, 3))

    ' Assert
    ExUnit.AreEqual 6, Actual, GetSig(MethodName)
    
    Exit Sub
ErrHandler:
    Lapis.ExUnit.TestFailRunTime GetSig(MethodName)

End Sub


Private Sub SumReturnsValueWhenSourceContainsReferencedTypesTest()

    On Error GoTo ErrHandler
    Const MethodName = "SumReturnsValueWhenSourceContainsReferencedTypesTest"

    ' Act
    Dim Source As Collection
    Set Source = CollectionExt.Make(CollectionExt.Make(1), CollectionExt.Make(1))

    ' Assert
    ExUnit.AreEqual 2, CollectionExt2.Sum(Source, Lambda.Create("$1.Count")), GetSig(MethodName)

    Exit Sub
ErrHandler:
    Lapis.ExUnit.TestFailRunTime GetSig(MethodName)

End Sub


Private Sub SumReturnsValueWhenSourceIsEmptyTest()

    On Error GoTo ErrHandler
    Const MethodName = "SumReturnsValueWhenSourceIsEmptyTest"

    ' Act
    Dim Actual As Variant
    Actual = CollectionExt2.Sum(CollectionExt.Make(Nothing))
    Actual = CollectionExt2.Sum(CollectionExt.Make(Nothing), Lambda.Create("$1.Count"))

    ' Assert
    ExUnit.AreEqual 0, Actual, GetSig(MethodName)

    Exit Sub
ErrHandler:
    Lapis.ExUnit.TestFailRunTime GetSig(MethodName)

End Sub


Private Sub SumReturnsValueWhenSourceContainsIntegersAndNothingTest()

    On Error GoTo ErrHandler
    Const MethodName = "SumReturnsValueWhenSourceContainsIntegersAndNothingTest"

    ' Act
    Dim Actual As Variant
    Actual = CollectionExt2.Sum(CollectionExt.Make(1, Nothing, 2))

    ' Assert
    ExUnit.AreEqual 3, Actual, GetSig(MethodName)

    Exit Sub
ErrHandler:
    Lapis.ExUnit.TestFailRunTime GetSig(MethodName)

End Sub


Private Sub SumReturnsArgumentNullErrorWhenSourceIsNothingTest()

    On Error GoTo ErrHandler
    Const ExpectedError As Long = ErrorCode.ArgumentNull
    Const MethodName = "SumReturnsArgumentNullErrorWhenSourceIsNothingTest"

    ' Act
    CollectionExt2.Sum Nothing
    
    ' Assert
ErrHandler:
    Lapis.ExUnit.IsException ExpectedError, Err.Number, GetSig(MethodName)

End Sub


Private Sub AverageReturnsValueWhenSourceContainsIntegersTest()

    On Error GoTo ErrHandler
    Const MethodName = "AverageReturnsValueWhenSourceContainsIntegersTest"

    ' Act
    Dim Actual As Variant
    Actual = CollectionExt2.Average(CollectionExt.Make(1, 2, 3))
    
    ' Assert
    ExUnit.AreEqual 2, Actual, GetSig(MethodName)
    

    Exit Sub
ErrHandler:
    Lapis.ExUnit.TestFailRunTime GetSig(MethodName)

End Sub


Private Sub AverageReturnsArgumentNullErrorWhenSourceIsEmptyTest()

    On Error GoTo ErrHandler
    Const ExpectedError As Long = ErrorCode.ArgumentNull
    Const MethodName = "AverageReturnsArgumentNullErrorWhenSourceIsEmptyTest"

    ' Act
    CollectionExt2.Average Nothing

    ' Assert
ErrHandler:
    Lapis.ExUnit.IsException ExpectedError, Err.Number, GetSig(MethodName)

End Sub


Private Sub AverageReturnsValueWhenSourceIsEmptyTest()

    On Error GoTo ErrHandler
    Const MethodName = "AverageReturnsValueWhenSourceIsEmptyTest"

    ' Act
    Dim Actual As Variant
    Actual = CollectionExt2.Average(New Collection)
    
    ' Assert
    ExUnit.AreEqual 0, Actual, GetSig(MethodName)

    Exit Sub
ErrHandler:
    Lapis.ExUnit.TestFailRunTime GetSig(MethodName)

End Sub


Private Sub AverageReturnsValueWhenSourceContainsIntegersAndNothingTest()

    On Error GoTo ErrHandler
    Const MethodName = "AverageReturnsValueWhenSourceContainsIntegersAndNothingTest"

    ' Arrange
    Dim Source As New Collection
    Source.Add 1
    Source.Add Nothing
    Source.Add 2

    ' Act
    Dim Actual As Variant
    Actual = CollectionExt2.Average(Source)
    
    ' Assert
    ExUnit.AreEqual Round(1.5, 2), Round(Actual, 2), GetSig(MethodName)

    Exit Sub
ErrHandler:
    Lapis.ExUnit.TestFailRunTime GetSig(MethodName)

End Sub


Private Sub AverageReturnsValueWhenSourceContainsOnlyNothingTest()

    On Error GoTo ErrHandler
    Const MethodName = "AverageReturnsValueWhenSourceContainsOnlyNothingTest"

    ' Act
    Dim Actual As Variant
    Actual = CollectionExt2.Average(CollectionExt.Make(Nothing, Nothing))

    ' Assert
    ExUnit.AreEqual 0, Actual, GetSig(MethodName)

    Exit Sub
ErrHandler:
    Lapis.ExUnit.TestFailRunTime GetSig(MethodName)

End Sub


Private Sub AverageReturnsValueWhenSourceContainsReferecenTypesTest()

    On Error GoTo ErrHandler
    Const MethodName = "AverageReturnsValueWhenSourceContainsReferecenTypesTest"

    ' Arrange
    Dim Stn1 As New TestStone: Stn1.Age = 10
    Dim Stn2 As New TestStone: Stn2.Age = 5
    
    ' Act
    Dim Actual As Variant
    Actual = CollectionExt2.Average(CollectionExt.Make(Stn1, Stn2), Lambda.Create("$1.Age"))

    ' Assert
    ExUnit.AreEqual Round(7.5, 2), Round(Actual, 2), GetSig(MethodName)

    Exit Sub
ErrHandler:
    ExUnit.TestFailRunTime GetSig(MethodName)

End Sub


Private Sub CountTest()

    On Error GoTo ErrHandler
    Const MethodName = "CountTest"

    Dim Source As Collection
    Set Source = CollectionExt.Make(1, 2, 3)
    
    ' Act & Assert
    ExUnit.AreEqual 3, CollectionExt2.Count(Source), GetSig(MethodName)
    ExUnit.AreEqual 2, CollectionExt2.Count(Source, Lambda.Create("$1 >= 2")), GetSig(MethodName)

    Exit Sub
ErrHandler:
    Lapis.ExUnit.TestFailRunTime GetSig(MethodName)

End Sub


Private Sub CountReturnsArgumentNullErrorWhenSourceIsNothingTest()

    On Error GoTo ErrHandler
    Const ExpectedError As Long = ErrorCode.ArgumentNull
    Const MethodName = "CountReturnsArgumentNullErrorWhenSourceIsNothingTest"
    
    ' Act
    CollectionExt2.Count Nothing

    ' Assert
ErrHandler:
    Lapis.ExUnit.IsException ExpectedError, Err.Number, GetSig(MethodName)

End Sub



Private Sub ConvertTest()

    On Error GoTo ErrHandler
    Const MethodName = "ConvertTest"
    
    ' Arrange
    Dim Expr As ICallable: Set Expr = Lambda.Create("len($1)")

    ' Act
    Dim Actual As Collection
    Set Actual = CollectionExt2.Convert(CollectionExt.Make("a", "ab", "abc"), Expr)
    
    ' Assert
    Lapis.ExUnit.AreEqual 1, Actual.Item(1), GetSig(MethodName)
    Lapis.ExUnit.AreEqual 2, Actual.Item(2), GetSig(MethodName)
    Lapis.ExUnit.AreEqual 3, Actual.Item(3), GetSig(MethodName)

    Exit Sub
ErrHandler:
    Lapis.ExUnit.TestFailRunTime GetSig(MethodName)

End Sub


Private Sub ConvertReturnsArgumentNullWhenSourceIsNothingTest()

    On Error GoTo ErrHandler
    Const ExpectedError As Long = ErrorCode.ArgumentNull
    Const MethodName = "ConvertReturnsArgumentNullWhenSourceIsNothingTest"

    ' Act
    Lapis.CollectionExt2.Convert Nothing, Lambda.Create("len($1)")

    ' Assert
ErrHandler:
    ExUnit.IsException ExpectedError, Err.Number, GetSig(MethodName)

End Sub


Private Sub ConvertReturnsArgumentNullWhenSelectorIsNothingTest()

    On Error GoTo ErrHandler
    Const ExpectedError As Long = ErrorCode.ArgumentNull
    Const MethodName = "ConvertReturnsArgumentNullWhenSelectorIsNothingTest"

    ' Act
    Lapis.CollectionExt2.Convert New Collection, Nothing

    ' Assert
ErrHandler:
    ExUnit.IsException ExpectedError, Err.Number, GetSig(MethodName)

End Sub


Private Sub FirstReturnsValueWhenSourceDoesSatisfyConditionTest()

    On Error GoTo ErrHandler
    Const MethodName = "FirstReturnsValueWhenSourceDoesSatisfyConditionTest"

    ' Value types
    Dim Predicate As ICallable: Set Predicate = Lambda.Create("$1 = 3")
    Dim Source As Collection: Set Source = CollectionExt.Make(1, 2, 3)

    ExUnit.AreEqual 1, CollectionExt2.First(Source), GetSig(MethodName)
    ExUnit.AreEqual 3, CollectionExt2.First(Source, Predicate), GetSig(MethodName)
    
    ' Reference Types
    Dim Stn1 As New TestStone: Stn1.Age = 1: Stn1.Weight = 10
    Dim Stn2 As New TestStone: Stn2.Age = 2: Stn2.Weight = 20
    Dim Stn3 As New TestStone: Stn3.Age = 1: Stn3.Weight = 30

    Set Source = CollectionExt.Make(Stn1, Stn2, Stn3)
    Set Predicate = Lambda.Create("$1.Age = 1")

    ' Act & Assert
    ExUnit.AreEqual 10, CollectionExt2.First(Source).Weight, GetSig(MethodName)
    ExUnit.AreEqual 10, CollectionExt2.First(Source, Lambda.Create("$1.Age = 1")).Weight, GetSig(MethodName)

    Exit Sub
ErrHandler:
    Lapis.ExUnit.TestFailRunTime GetSig(MethodName)

End Sub


Private Sub FirstReturnsArgumentNullErrorWhenSourceIsNothingTest()

    On Error GoTo ErrHandler
    Const ExpectedError As Long = ErrorCode.ArgumentNull
    Const MethodName = "FirstReturnsArgumentNullErrorWhenSourceIsNothingTest"

    ' Act
    CollectionExt2.First Nothing

    ' Assert
ErrHandler:
    Lapis.ExUnit.IsException ExpectedError, Err.Number, GetSig(MethodName)

End Sub


Private Sub FirstReturnsInvalidOperationErrorWhenSourceDoesNotSatisfyConditionTest()

    On Error GoTo ErrHandler
    Const ExpectedError As Long = ErrorCode.InvalidOperation
    Const MethodName = "FirstReturnsInvalidOperationErrorWhenSourceDoesNotSatisfyConditionTest"

    ' Act
    CollectionExt2.First CollectionExt.Make(1, 2, 3), Lambda.Create("$1 = 5")

    ' Assert
ErrHandler:
    Lapis.ExUnit.IsException ExpectedError, Err.Number, GetSig(MethodName)

End Sub


Private Sub FirstReturnsInvalidOperationErrorWhenSourceIsEmptyTest()

    On Error GoTo ErrHandler
    Const ExpectedError As Long = ErrorCode.InvalidOperation
    Const MethodName = "FirstReturnsInvalidOperationErrorWhenSourceIsEmptyTest"

    ' Act
    CollectionExt2.First New Collection

ErrHandler:
    Lapis.ExUnit.IsException ExpectedError, Err.Number, GetSig(MethodName)

End Sub



Private Sub LastReturnsValueWhenSourceDoesSatisfyConditionTest()

    On Error GoTo ErrHandler
    Const MethodName = "LastReturnsValueWhenSourceDoesSatisfyConditionTest"

    ' Arrange
    Dim Stn1 As New TestStone: Stn1.Age = 1: Stn1.Weight = 10
    Dim Stn2 As New TestStone: Stn2.Age = 2: Stn2.Weight = 20
    Dim Stn3 As New TestStone: Stn3.Age = 1: Stn3.Weight = 30

    Dim Source As Collection: Set Source = CollectionExt.Make(Stn1, Stn2, Stn3)

    ' Act & Assert
    ExUnit.AreEqual 30, CollectionExt2.Last(Source).Weight, GetSig(MethodName)
    ExUnit.AreEqual 30, CollectionExt2.Last(Source, Lambda.Create("$1.Age = 1")).Weight, GetSig(MethodName)
    
    Exit Sub
ErrHandler:
    Lapis.ExUnit.TestFailRunTime GetSig(MethodName)

End Sub


Private Sub LastReturnsArgumentNullErrorWhenSourceIsNothingTest()

    On Error GoTo ErrHandler
    Const ExpectedError As Long = ErrorCode.ArgumentNull
    Const MethodName = "LastReturnsArgumentNullErrorWhenSourceIsNothingTest"

    ' Act
    CollectionExt2.Last Nothing

    ' Assert
ErrHandler:
    Lapis.ExUnit.IsException ExpectedError, Err.Number, GetSig(MethodName)

End Sub


Private Sub LastReturnsInvalidOperationErrorWhenSourceDoesNotSatisfyConditionTest()

    On Error GoTo ErrHandler
    Const ExpectedError As Long = ErrorCode.InvalidOperation
    Const MethodName = "LastReturnsInvalidOperationErrorWhenSourceDoesNotSatisfyConditionTest"

    ' Act
    CollectionExt2.Last CollectionExt.Make(1, 2, 3), Lambda.Create("$1 = 5")

    ' Assert
ErrHandler:
    Lapis.ExUnit.IsException ExpectedError, Err.Number, GetSig(MethodName)

End Sub


Private Sub LastReturnsInvalidOperationErrorWhenSourceIsEmptyTest()

    On Error GoTo ErrHandler
    Const ExpectedError As Long = ErrorCode.InvalidOperation
    Const MethodName = "LastReturnsInvalidOperationErrorWhenSourceIsEmptyTest"

    ' Act
    CollectionExt2.Last New Collection

    ' Assert
ErrHandler:
    Lapis.ExUnit.IsException ExpectedError, Err.Number, GetSig(MethodName)

End Sub

