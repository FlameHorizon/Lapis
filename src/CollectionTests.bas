Attribute VB_Name = "CollectionTests"
'@Folder("Tests")
Option Explicit

Private Const ModuleName As String = "CollectionTests"


Public Sub Start()
    
    DistinctTest
    DistinctReturnsItemWhenSingleItemGivenTest
    DistinctReturnsEmptyCollectionWhenSourceIsEmptyTest
    DistinctReturnsArgumentNullErrorWhenSourceIsNothingTest
    
    ExceptTest
    ExceptReturnsEmptyCollectionWhenFirstAndSecondAreEmptyTest
    ExceptReturnsEmptyCollectionWhenFirstIsEmptyTest
    ExceptReturnsArgumentNullErrorIfFirstIsNothingTest
    ExceptReturnsArgumentNullErrorIfSecondIsNothingTest
    
    IntersectTest
    IntersectReturnsEmptyCollectionWhenFirstAndSecondAreEmptyTest
    IntersectReturnsEmptyCollectionWhenFirstIsEmpty
    IntersectReturnsArgumentNullErrorIfFirstIsNothingTest
    IntersectReturnsArgumentNullErrorIfSecondIsNothingTest
    
    MinWhenSourceContainsValueTypesTest
    MinWhenSourceContainsReferencedTypesTest
    MinReturnsArgumentNullErrorWhenSourceIsNothingTest
    MinReturnsLowestValueWhenNothingIsPresentTest
    MinReturnsNothingWhenOnlyNothingIsPresentTest
    
    MaxWhenSourceContainsValueTypesTest
    MaxWhenSourceContainsReferencedTypesTest
    MaxReturnsArgumentNullErrorWhenSourceIsNothingTest
    MaxReturnsArgumentNullErrorWhenComparerIsNothingTest
    MaxReturnsHighestValueWhenNothingIsPresentTest
    MaxReturnsNothingWhenOnlyNothingIsPresentTest
    
    RangeTest
    RangeReturnsArgumentOutOfRangeErrorWhenCountIsLessThanZeroTest
    RangeReturnsArgumentOutOfRangeWhenStartAndCountExceedsLimitTest
    RangeReturnsEmptyCollectionWhenCountIsZeroTest
    
    RepeatValueTypesTest
    RepeatReferencedTypesTest
    RepeatReturnsArgumentOutOfRangeErrorWhenCountIsLessThanZeroTest
    RepeatReturnsEmptyCollectionWhenCountIsZeroTest
    
    ReverseTest
    ReverseReturnsArgumentNullErrorWhenSourceIsNothingTest
    ReverseReturnsEmptyCollectionWhenSourceIsEmptyTest
    
    TakeTest
    TakeReturnsArgumentNullErrorWhenSourceIsNothingTest
    TakeReturnsEmptyCollectionWhenCountIsZeroTest
    TakeReturnsEmptyCollectionWhenCountIsNegativeTest
    TakeReturnsEmptyCollectionWhenSourceIsEmptyTest
    
    SkipTest
    SkipReturnsArgumentNullErrorWhenSourceIsNothingTest
    SkipReturnsEntireSourceWhenCountIsZeroOrNegativeTest
    SkipReturnsEmptyCollectionWhenCountIsGreaterThanNumberOfItemsInSourceTest
    
    SequenceEqualReturnsTrueWhenFirstAndSecondContainsEqualElementsTest
    SequenceEqualReturnsFalseWhenFirstAndSecondContainsNotEqualElementsTest
    SequenceEqualRetrunsFalseWhenFirstAndSecondAreDiffrentSizesTest
    SequenceEqualReturnsArgumentNullErrorWhenFirstIsNothingTest
    SequenceEqualReturnsArgumentNullErrorWhenSecondIsNothingTest
    SequenceEqualReturnsArgumentNullErrorWhenComparerIsNothingTest

    IndexOfTest
    IndexOfReturnsValueWhenItemNotExistsInSourceTest
    IndexOfReturnsMinusOneWhenItemCantBeComparedTest
    IndexOfReturnsValueWhenItemImplementsIComparableTest
    
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
    
    SelectOneReturnsValueWhenSourceDoesSatisfyConditionTest
    SelectOneReturnsArgumentNullErrorWhenSourceIsNothingTest
    SelectOneReturnsInvalidOperationErrorWhenSourceDoesNotSatisfyConditionTest
    SelectOneReturnsInvalidOperationErrorWhenSourceIsEmptyTest
    SelectOneReturnsInvalidOperationErrorWhenSourceDoesSatisfyConditionMultipleTimesTest

End Sub


Private Sub DistinctTest()

    On Error GoTo ErrHandler
    Const MethodName = "DistinctTest"

    ' Act
    Dim Actual As Collection
    Set Actual = CollectionExt.Distinct(CollectionExt.Make(1, 2, 3, 2))

    ' Assert
    ExUnit.AreEqual 3, Actual.Count, GetSig(MethodName)
    ExUnit.AreEqual 1, Actual.Item(1), GetSig(MethodName)
    ExUnit.AreEqual 2, Actual.Item(2), GetSig(MethodName)
    ExUnit.AreEqual 3, Actual.Item(3), GetSig(MethodName)

    Exit Sub
ErrHandler:
    Lapis.ExUnit.TestFailRunTime GetSig(MethodName)

End Sub


Private Function GetSig(ByVal MethodName As String) As String
    GetSig = ModuleName & "." & MethodName
End Function


Private Sub DistinctReturnsItemWhenSingleItemGivenTest()

    On Error GoTo ErrHandler
    Const MethodName = "DistinctReturnsItemWhenSingleItemGivenTest"

    ' Act
    Dim Actual As Collection
    Set Actual = CollectionExt.Distinct(CollectionExt.Make(1))
    
    ' Assert
    ExUnit.AreEqual 1, Actual.Item(1), GetSig(MethodName)

    Exit Sub
ErrHandler:
    Lapis.ExUnit.TestFailRunTime GetSig(MethodName)

End Sub


Private Sub DistinctReturnsEmptyCollectionWhenSourceIsEmptyTest()

    On Error GoTo ErrHandler
    Const MethodName = "DistinctReturnsEmptyCollectionWhenSourceIsEmptyTest"

    ' Act
    Dim Actual As Collection
    Set Actual = CollectionExt.Distinct(New Collection)
    
    ' Assert
    ExUnit.AreEqual 0, Actual.Count, GetSig(MethodName)

    Exit Sub
ErrHandler:
    Lapis.ExUnit.TestFailRunTime GetSig(MethodName)

End Sub


Private Sub DistinctReturnsArgumentNullErrorWhenSourceIsNothingTest()

    On Error GoTo ErrHandler
    Const ExpectedError As Long = ErrorCode.ArgumentNull
    Const MethodName = "DistinctReturnsArgumentNullErrorWhenSourceIsNothingTest"
    
    ' Act
    CollectionExt.Distinct Nothing
    
    ' Assert
ErrHandler:
    Lapis.ExUnit.IsException ExpectedError, Err.Number, GetSig(MethodName)

End Sub


Private Function MakeStone(ByVal Weight As Single, ByVal Age As Single) As TestStone
    
    Dim Output As New TestStone
    With Output
        .Weight = Weight
        .Age = Age
    End With
    Set MakeStone = Output
    
End Function


Private Sub ExceptTest()

    On Error GoTo ErrHandler
    Const MethodName = "ExceptTest"
    
    ' Act
    Dim Actual As Collection
    Set Actual = CollectionExt.Except(CollectionExt.Make(1, 2, 3), CollectionExt.Make(1, 3))
    
    ' Assert
    ExUnit.AreEqual 1, Actual.Count, GetSig(MethodName)
    ExUnit.AreEqual 2, Actual.Item(1), GetSig(MethodName)

    Exit Sub
ErrHandler:
    Lapis.ExUnit.TestFailRunTime GetSig(MethodName)

End Sub


Private Sub ExceptReturnsEmptyCollectionWhenFirstAndSecondAreEmptyTest()

    On Error GoTo ErrHandler
    Const MethodName = "ExceptReturnsEmptyCollectionWhenFirstAndSecondAreEmptyTest"

    ' Act
    Dim Actual As Collection
    Set Actual = CollectionExt.Except(New Collection, New Collection, New LongEqualityComparer)

    ' Assert
    ExUnit.AreEqual 0, Actual.Count, GetSig(MethodName)

    Exit Sub
ErrHandler:
    Lapis.ExUnit.TestFailRunTime GetSig(MethodName)

End Sub


Private Sub ExceptReturnsEmptyCollectionWhenFirstIsEmptyTest()

    On Error GoTo ErrHandler
    Const MethodName = "ExceptReturnsEmptyCollectionWhenFirstIsEmptyTest"

    ' Act
    Dim Actual As Collection
    Set Actual = CollectionExt.Except(New Collection, CollectionExt.Make(1, 2, 3), New LongEqualityComparer)
    
    ' Assert
    ExUnit.AreEqual 0, Actual.Count, GetSig(MethodName)
    
    Exit Sub
ErrHandler:
    Lapis.ExUnit.TestFailRunTime GetSig(MethodName)

End Sub


Private Sub ExceptReturnsArgumentNullErrorIfFirstIsNothingTest()

    On Error GoTo ErrHandler
    Const ExpectedError As Long = ErrorCode.ArgumentNull
    Const MethodName = "ExceptReturnsArgumentNullErrorIfFirstIsNothingTest"

    ' Act
    CollectionExt.Except Nothing, New Collection, New LongEqualityComparer
    
    ' Assert
ErrHandler:
    Lapis.ExUnit.IsException ExpectedError, Err.Number, GetSig(MethodName)

End Sub


Private Sub ExceptReturnsArgumentNullErrorIfSecondIsNothingTest()

    On Error GoTo ErrHandler
    Const ExpectedError As Long = ErrorCode.ArgumentNull
    Const MethodName = "ExceptReturnsArgumentNullErrorIfSecondIsNothingTest"
    
    ' Act
    CollectionExt.Except New Collection, Nothing, New LongEqualityComparer
    
    ' Assert
ErrHandler:
    Lapis.ExUnit.IsException ExpectedError, Err.Number, GetSig(MethodName)

End Sub


Private Sub IntersectTest()

    On Error GoTo ErrHandler
    Const MethodName = "IntersectTest"
    
    ' Act
    Dim Actual As Collection
    Set Actual = CollectionExt.Intersect(CollectionExt.Make(1, 2, 3), _
                                         CollectionExt.Make(2, 3))
    
    ' Assert
    ExUnit.AreEqual 2, Actual.Count, GetSig(MethodName)
    ExUnit.AreEqual 2, Actual.Item(1), GetSig(MethodName)
    ExUnit.AreEqual 3, Actual.Item(2), GetSig(MethodName)

    Exit Sub
ErrHandler:
    Lapis.ExUnit.TestFailRunTime GetSig(MethodName)
    
End Sub


Private Sub IntersectReturnsEmptyCollectionWhenFirstAndSecondAreEmptyTest()

    On Error GoTo ErrHandler
    Const MethodName = "IntersectReturnsEmptyCollectionWhenFirstAndSecondAreEmptyTest"

    ' Act
    Dim Actual As Collection
    Set Actual = CollectionExt.Intersect(New Collection, New Collection, New LongEqualityComparer)

    ' Assert
    ExUnit.AreEqual 0, Actual.Count, GetSig(MethodName)

    Exit Sub
ErrHandler:
    Lapis.ExUnit.TestFailRunTime GetSig(MethodName)

End Sub


Private Sub IntersectReturnsArgumentNullErrorIfFirstIsNothingTest()

    On Error GoTo ErrHandler
    Const ExpectedError As Long = ErrorCode.ArgumentNull
    Const MethodName = "IntersectReturnsArgumentNullErrorIfFirstIsNothingTest"

    ' Act
    CollectionExt.Intersect Nothing, New Collection, New LongEqualityComparer
    
    ' Assert
ErrHandler:
    Lapis.ExUnit.IsException ExpectedError, Err.Number, GetSig(MethodName)

End Sub


Private Sub IntersectReturnsArgumentNullErrorIfSecondIsNothingTest()

    On Error GoTo ErrHandler
    Const ExpectedError As Long = ErrorCode.ArgumentNull
    Const MethodName = "IntersectReturnsArgumentNullErrorIfSecondIsNothingTest"
    
    ' Act
    CollectionExt.Intersect New Collection, Nothing, New LongEqualityComparer
    
    ' Assert
ErrHandler:
    Lapis.ExUnit.IsException ExpectedError, Err.Number, GetSig(MethodName)

End Sub


Private Sub IntersectReturnsEmptyCollectionWhenFirstIsEmpty()

    On Error GoTo ErrHandler
    Const MethodName = "IntersectReturnsEmptyCollectionWhenFirstIsEmpty"

    ' Act
    Dim Actual As Collection
    Set Actual = CollectionExt.Intersect(New Collection, CollectionExt.Make(2, 3), New LongEqualityComparer)
    
    ' Assert
    ExUnit.AreEqual 0, Actual.Count, GetSig(MethodName)

    Exit Sub
ErrHandler:
    Lapis.ExUnit.TestFailRunTime GetSig(MethodName)

End Sub


Private Sub MinWhenSourceContainsValueTypesTest()

    On Error GoTo ErrHandler
    Const MethodName = "MinWhenSourceContainsValueTypesTest"

    ' Assert
    ExUnit.AreEqual 1, CollectionExt.Min(CollectionExt.Make(3, 2, 1)), GetSig(MethodName)
    ExUnit.AreEqual "a", CollectionExt.Min(CollectionExt.Make("s", "b", "a")), GetSig(MethodName)
    
    Exit Sub
ErrHandler:
    Lapis.ExUnit.TestFailRunTime GetSig(MethodName)

End Sub


Private Sub MinWhenSourceContainsReferencedTypesTest()

    On Error GoTo ErrHandler
    Const MethodName = "MinWhenSourceContainsReferencedTypesTest"

    ' Arrange
    Dim Source As New Collection
    Source.Add CollectionExt.Make(1, 2)
    Source.Add CollectionExt.Make(1)

    ' Act
    Dim Actual As Object
    Set Actual = CollectionExt.Min(Source, Lambda.Create("$1.Count"))
    
    ' Assert
    ExUnit.AreEqual 1, Actual.Count, GetSig(MethodName)
    ExUnit.AreEqual 1, Actual.Item(1), GetSig(MethodName)
    
    Exit Sub
ErrHandler:
    Lapis.ExUnit.TestFailRunTime GetSig(MethodName)

End Sub


Private Sub MinReturnsArgumentNullErrorWhenSourceIsNothingTest()

    On Error GoTo ErrHandler
    Const ExpectedError As Long = ErrorCode.ArgumentNull
    Const MethodName = "MinReturnsArgumentNullErrorWhenSourceIsNothingTest"

    ' Act
    CollectionExt.Min Nothing
    
    ' Assert
ErrHandler:
    Lapis.ExUnit.IsException ExpectedError, Err.Number, GetSig(MethodName)

End Sub


Private Sub MinReturnsLowestValueWhenNothingIsPresentTest()

    On Error GoTo ErrHandler
    Const MethodName = "MinReturnsLowestValueWhenNothingIsPresentTest"

    ' Arrange
    Dim Source As Collection
    Set Source = CollectionExt.Make(New Collection, Nothing, Nothing, CollectionExt.Make(1))
    
    ' Act
    Dim Actual As Object
    Set Actual = CollectionExt.Min(Source, Lambda.Create("$1.Count"))

    ' Assert
    ExUnit.AreEqual 0, Actual.Count, GetSig(MethodName)
    
    Exit Sub
ErrHandler:
    Lapis.ExUnit.TestFailRunTime GetSig(MethodName)

End Sub


Private Sub MinReturnsNothingWhenOnlyNothingIsPresentTest()

    On Error GoTo ErrHandler
    Const MethodName = "MinReturnsNothingWhenOnlyNothingIsPresentTest"

    ' Act
    Dim Actual As Object
    Set Actual = CollectionExt.Min(CollectionExt.Make(Nothing))
    
    ' Assert
    ExUnit.AreSame Nothing, Actual, GetSig(MethodName)
    
    Exit Sub
ErrHandler:
    Lapis.ExUnit.TestFailRunTime GetSig(MethodName)

End Sub


Private Sub MaxWhenSourceContainsValueTypesTest()

    On Error GoTo ErrHandler
    Const MethodName = "MaxWhenSourceContainsValueTypesTest"

    ' Act
    Dim Actual As Variant
    Actual = CollectionExt.Max(CollectionExt.Make(3, 2, 1), New LongComparer)
    
    ' Assert
    ExUnit.AreEqual 3, Actual, GetSig(MethodName)

    Exit Sub
ErrHandler:
    Lapis.ExUnit.TestFailRunTime GetSig(MethodName)

End Sub


Private Sub MaxWhenSourceContainsReferencedTypesTest()

    On Error GoTo ErrHandler
    Const MethodName = "MaxWhenSourceContainsReferencedTypesTest"

    ' Arrange
    Dim Source As New Collection
    Source.Add CollectionExt.Make(1, 2)
    Source.Add CollectionExt.Make(1)

    ' Act
    Dim Actual As Object
    Set Actual = CollectionExt.Max(Source, New TestCollectionComparer)
    
    ' Assert
    ExUnit.AreEqual 2, Actual.Count, GetSig(MethodName)
    ExUnit.AreEqual 1, Actual.Item(1), GetSig(MethodName)
    ExUnit.AreEqual 2, Actual.Item(2), GetSig(MethodName)
    
    Exit Sub
ErrHandler:
    Lapis.ExUnit.TestFailRunTime GetSig(MethodName)

End Sub


Private Sub MaxReturnsArgumentNullErrorWhenSourceIsNothingTest()

    On Error GoTo ErrHandler
    Const ExpectedError As Long = ErrorCode.ArgumentNull
    Const MethodName = "MaxReturnsArgumentNullErrorWhenSourceIsNothingTest"

    ' Act
    CollectionExt.Max Nothing, New LongComparer
    
    ' Assert
ErrHandler:
    Lapis.ExUnit.IsException ExpectedError, Err.Number, GetSig(MethodName)

End Sub


Private Sub MaxReturnsArgumentNullErrorWhenComparerIsNothingTest()

    On Error GoTo ErrHandler
    Const ExpectedError As Long = ErrorCode.ArgumentNull
    Const MethodName = "MaxReturnsArgumentNullErrorWhenComparerIsNothingTest"

    ' Act
    CollectionExt.Max New Collection, Nothing
    
    ' Assert
ErrHandler:
    Lapis.ExUnit.IsException ExpectedError, Err.Number, GetSig(MethodName)

End Sub


Private Sub MaxReturnsHighestValueWhenNothingIsPresentTest()

    On Error GoTo ErrHandler
    Const MethodName = "MaxReturnsHighestValueWhenNothingIsPresentTest"

    ' Arrange
    Dim Source As Collection
    Set Source = CollectionExt.Make(New Collection, Nothing, CollectionExt.Make(1))
    
    ' Act
    Dim Actual As Object
    Set Actual = CollectionExt.Max(Source, New TestCollectionComparer)

    ' Assert
    ExUnit.AreEqual 1, Actual.Count, GetSig(MethodName)
    
    Exit Sub
ErrHandler:
    Lapis.ExUnit.TestFailRunTime GetSig(MethodName)

End Sub


Private Sub MaxReturnsNothingWhenOnlyNothingIsPresentTest()

    On Error GoTo ErrHandler
    Const MethodName = "MaxReturnsNothingWhenOnlyNothingIsPresentTest"

    ' Act
    Dim Actual As Object
    Set Actual = CollectionExt.Max(CollectionExt.Make(Nothing), New TestCollectionComparer)
    
    ' Assert
    ExUnit.AreSame Nothing, Actual, GetSig(MethodName)
    
    Exit Sub
ErrHandler:
    Lapis.ExUnit.TestFailRunTime GetSig(MethodName)

End Sub


Private Sub RangeTest()

    On Error GoTo ErrHandler
    Const MethodName = "RangeTest"
    
    ' Act
    Dim Actual As Collection
    Set Actual = CollectionExt.Range(0, 3)
    
    ' Assert
    ExUnit.AreEqual 3, Actual.Count, GetSig(MethodName)
    ExUnit.AreEqual 0, Actual.Item(1), GetSig(MethodName)
    ExUnit.AreEqual 1, Actual.Item(2), GetSig(MethodName)
    ExUnit.AreEqual 2, Actual.Item(3), GetSig(MethodName)

    Exit Sub
ErrHandler:
    Lapis.ExUnit.TestFailRunTime GetSig(MethodName)

End Sub


Private Sub RangeReturnsArgumentOutOfRangeErrorWhenCountIsLessThanZeroTest()

    On Error GoTo ErrHandler
    Const ExpectedError As Long = ErrorCode.ArgumentOutOfRange
    Const MethodName = "RangeReturnsArgumentOutOfRangeErrorWhenCountIsLessThanZeroTest"

    ' Act
    CollectionExt.Range 0, -2
        
    ' Assert
ErrHandler:
    Lapis.ExUnit.IsException ExpectedError, Err.Number, GetSig(MethodName)

End Sub


Private Sub RangeReturnsArgumentOutOfRangeWhenStartAndCountExceedsLimitTest()

    On Error GoTo ErrHandler
    Const ExpectedError As Long = ErrorCode.ArgumentOutOfRange
    Const MethodName = "RangeReturnsArgumentOutOfRangeWhenStartAndCountExceedsLimitTest"

    ' Act
    CollectionExt.Range System.LongMaxValue - 1, 2
    
    ' Assert
ErrHandler:
    Lapis.ExUnit.IsException ExpectedError, Err.Number, GetSig(MethodName)

End Sub


Private Sub RangeReturnsEmptyCollectionWhenCountIsZeroTest()

    On Error GoTo ErrHandler
    Const MethodName = "RangeReturnsEmptyCollectionWhenCountIsZeroTest"

    ' Act
    Dim Actual As Collection
    Set Actual = CollectionExt.Range(1, 0)
    
    ' Assert
    ExUnit.AreEqual 0, Actual.Count, GetSig(MethodName)
    
    Exit Sub
ErrHandler:
    Lapis.ExUnit.TestFailRunTime GetSig(MethodName)

End Sub


Private Sub RepeatValueTypesTest()

    On Error GoTo ErrHandler
    Const MethodName = "RepeatValueTypesTest"

    ' Act
    Dim Actual As Collection
    Set Actual = CollectionExt.Repeat("a", 3)
    
    ' Assert
    ExUnit.AreEqual 3, Actual.Count, GetSig(MethodName)
    ExUnit.AreEqual "a", Actual.Item(1), GetSig(MethodName)
    ExUnit.AreEqual "a", Actual.Item(2), GetSig(MethodName)
    ExUnit.AreEqual "a", Actual.Item(3), GetSig(MethodName)

    Exit Sub
ErrHandler:
    Lapis.ExUnit.TestFailRunTime GetSig(MethodName)

End Sub


Private Sub RepeatReferencedTypesTest()

    On Error GoTo ErrHandler
    Const MethodName = "RepeatReferencedTypesTest"

    ' Act
    Dim Actual As Collection
    Set Actual = CollectionExt.Repeat(ThisWorkbook, 2)
    
    ' Assert
    ExUnit.AreEqual 2, Actual.Count, GetSig(MethodName)

    Exit Sub
ErrHandler:
    Lapis.ExUnit.TestFailRunTime GetSig(MethodName)

End Sub


Private Sub RepeatReturnsArgumentOutOfRangeErrorWhenCountIsLessThanZeroTest()

    On Error GoTo ErrHandler
    Const ExpectedError As Long = ErrorCode.ArgumentOutOfRange
    Const MethodName = "RepeatReturnsArgumentOutOfRangeErrorWhenCountIsLessThanZeroTest"

    ' Act
    CollectionExt.Repeat "a", -1
    
    ' Assert
ErrHandler:
    Lapis.ExUnit.IsException ExpectedError, Err.Number, GetSig(MethodName)

End Sub


Private Sub RepeatReturnsEmptyCollectionWhenCountIsZeroTest()

    On Error GoTo ErrHandler
    Const MethodName = "RepeatReturnsEmptyCollectionWhenCountIsZeroTest"

    ' Act
    Dim Actual As Collection
    Set Actual = CollectionExt.Repeat("a", 0)
    
    ' Assert
    ExUnit.AreEqual 0, Actual.Count, GetSig(MethodName)

    Exit Sub
ErrHandler:
    Lapis.ExUnit.TestFailRunTime GetSig(MethodName)

End Sub


Private Sub ReverseTest()

    On Error GoTo ErrHandler
    Const MethodName = "ReverseTest"

    ' Act
    Dim Actual As Collection
    Set Actual = CollectionExt.Reverse(CollectionExt.Make(1, 2, 3))
    
    ' Assert
    ExUnit.AreEqual 3, Actual.Count, GetSig(MethodName)
    ExUnit.AreEqual 3, Actual.Item(1), GetSig(MethodName)
    ExUnit.AreEqual 2, Actual.Item(2), GetSig(MethodName)
    ExUnit.AreEqual 1, Actual.Item(3), GetSig(MethodName)

    Exit Sub
ErrHandler:
    Lapis.ExUnit.TestFailRunTime GetSig(MethodName)

End Sub


Private Sub ReverseReturnsArgumentNullErrorWhenSourceIsNothingTest()

    On Error GoTo ErrHandler
    Const ExpectedError As Long = ErrorCode.ArgumentNull
    Const MethodName = "ReverseReturnsArgumentNullErrorWhenSourceIsNothingTest"

    ' Act
    CollectionExt.Reverse Nothing
    
    ' Assert
ErrHandler:
    Lapis.ExUnit.IsException ExpectedError, Err.Number, GetSig(MethodName)

End Sub


Private Sub ReverseReturnsEmptyCollectionWhenSourceIsEmptyTest()

    On Error GoTo ErrHandler
    Const MethodName = "ReverseReturnsEmptyCollectionWhenSourceIsEmptyTest"

    ' Act
    Dim Actual As Collection
    Set Actual = CollectionExt.Reverse(New Collection)
    
    ' Assert
    ExUnit.AreEqual 0, Actual.Count, GetSig(MethodName)
    
    Exit Sub
ErrHandler:
    Lapis.ExUnit.TestFailRunTime GetSig(MethodName)

End Sub


Private Sub TakeTest()

    On Error GoTo ErrHandler
    Const MethodName = "TakeTest"

    ' Act
    Dim Actual As Collection
    Set Actual = CollectionExt.Take(CollectionExt.Make(1, 2, 3), 2)

    ' Assert
    ExUnit.AreEqual 2, Actual.Count, GetSig(MethodName)
    ExUnit.AreEqual 1, Actual.Item(1), GetSig(MethodName)
    ExUnit.AreEqual 2, Actual.Item(2), GetSig(MethodName)

    Exit Sub
ErrHandler:
    Lapis.ExUnit.TestFailRunTime GetSig(MethodName)

End Sub


Private Sub TakeReturnsArgumentNullErrorWhenSourceIsNothingTest()

    On Error GoTo ErrHandler
    Const ExpectedError As Long = ErrorCode.ArgumentNull
    Const MethodName = "TakeReturnsArgumentNullErrorWhenSourceIsNothingTest"

    ' Act
    CollectionExt.Take Nothing, 1
    
    ' Assert
ErrHandler:
    Lapis.ExUnit.IsException ExpectedError, Err.Number, GetSig(MethodName)

End Sub


Private Sub TakeReturnsEmptyCollectionWhenCountIsZeroTest()

    On Error GoTo ErrHandler
    Const MethodName = "TakeReturnsEmptyCollectionWhenCountIsZeroTest"
    
    ' Act
    Dim Actual As Collection
    Set Actual = CollectionExt.Take(CollectionExt.Make(1, 2, 3), 0)
    
    ' Assert
    ExUnit.AreEqual 0, Actual.Count, GetSig(MethodName)
    
    Exit Sub
ErrHandler:
    Lapis.ExUnit.TestFailRunTime GetSig(MethodName)

End Sub


Private Sub TakeReturnsEmptyCollectionWhenCountIsNegativeTest()

    On Error GoTo ErrHandler
    Const MethodName = "TakeReturnsEmptyCollectionWhenCountIsNegativeTest"

    ' Act
    Dim Actual As Collection
    Set Actual = CollectionExt.Take(CollectionExt.Make(1, 2, 3), -10)
    
    ' Assert
    ExUnit.AreEqual 0, Actual.Count, GetSig(MethodName)

    Exit Sub
ErrHandler:
    Lapis.ExUnit.TestFailRunTime GetSig(MethodName)

End Sub


Private Sub TakeReturnsEmptyCollectionWhenSourceIsEmptyTest()

    On Error GoTo ErrHandler
    Const MethodName = "TakeReturnsEmptyCollectionWhenSourceIsEmptyTest"

    ' Act
    Dim Actual As Collection
    Set Actual = CollectionExt.Take(New Collection, 2)
    
    ' Assert
    ExUnit.AreEqual 0, Actual.Count, GetSig(MethodName)
    
    Exit Sub
ErrHandler:
    Lapis.ExUnit.TestFailRunTime GetSig(MethodName)

End Sub


Private Sub SkipTest()

    On Error GoTo ErrHandler
    Const MethodName = "SkipTest"
    
    ' Act
    Dim Actual As Collection
    Set Actual = CollectionExt.Skip(CollectionExt.Make(1, 2, 3), 2)

    ' Assert
    ExUnit.AreEqual 1, Actual.Count, GetSig(MethodName)
    ExUnit.AreEqual 3, Actual.Item(1), GetSig(MethodName)

    Exit Sub
ErrHandler:
    Lapis.ExUnit.TestFailRunTime GetSig(MethodName)

End Sub


Private Sub SkipReturnsArgumentNullErrorWhenSourceIsNothingTest()

    On Error GoTo ErrHandler
    Const ExpectedError As Long = ErrorCode.ArgumentNull
    Const MethodName = "SkipReturnsArgumentNullErrorWhenSourceIsNothingTest"

    ' Act
    CollectionExt.Skip Nothing, 1
    
    ' Assert
ErrHandler:
    Lapis.ExUnit.IsException ExpectedError, Err.Number, GetSig(MethodName)

End Sub


Private Sub SkipReturnsEntireSourceWhenCountIsZeroOrNegativeTest()

    On Error GoTo ErrHandler
    Const MethodName = "SkipReturnsEntireSourceWhenCountIsZeroOrNegativeTest"
    
    ' Arrange
    Dim Source As Collection
    Set Source = CollectionExt.Make(1, 2, 3)
    
    ' Act & Assert
    ExUnit.AreEqual 3, CollectionExt.Skip(Source, 0).Count, GetSig(MethodName)
    ExUnit.AreEqual 3, CollectionExt.Skip(Source, -1).Count, GetSig(MethodName)

    Exit Sub
ErrHandler:
    Lapis.ExUnit.TestFailRunTime GetSig(MethodName)

End Sub


Private Sub SkipReturnsEmptyCollectionWhenCountIsGreaterThanNumberOfItemsInSourceTest()

    On Error GoTo ErrHandler
    Const MethodName = "SkipReturnsEmptyCollectionWhenCountIsGreaterThanNumberOfItemsInSourceTest"
    
    ' Act & Assert
    ExUnit.AreEqual 0, CollectionExt.Skip(CollectionExt.Make(1, 2, 3), 5).Count, GetSig(MethodName)

    Exit Sub
ErrHandler:
    Lapis.ExUnit.TestFailRunTime GetSig(MethodName)

End Sub


Private Sub SequenceEqualReturnsTrueWhenFirstAndSecondContainsEqualElementsTest()

    On Error GoTo ErrHandler
    Const MethodName = "SequenceEqualReturnsTrueWhenFirstAndSecondContainsEqualElementsTest"

    ' Arrange
    Dim First As Collection
    Set First = CollectionExt.Make(1, 2, 3)
    
    Dim Second As Collection
    Set Second = CollectionExt.Make(1, 2, 3)

    ' Act & Assert
    ExUnit.IsTrue CollectionExt.SequenceEqual(First, Second, New LongEqualityComparer), GetSig(MethodName)

    Exit Sub
ErrHandler:
    Lapis.ExUnit.TestFailRunTime GetSig(MethodName)

End Sub


Private Sub SequenceEqualReturnsFalseWhenFirstAndSecondContainsNotEqualElementsTest()

    On Error GoTo ErrHandler
    Const MethodName = "SequenceEqualReturnsFalseWhenFirstAndSecondContainsNotEqualElementsTest"

    ' Arrange
    Dim First As Collection
    Set First = CollectionExt.Make(1, 2, 3)
    
    Dim Second As Collection
    Set Second = CollectionExt.Make(1, 2, 4)

    ' Act & Assert
    ExUnit.IsFalse CollectionExt.SequenceEqual(First, Second, New LongEqualityComparer), GetSig(MethodName)

    Exit Sub
ErrHandler:
    Lapis.ExUnit.TestFailRunTime GetSig(MethodName)

End Sub


Private Sub SequenceEqualRetrunsFalseWhenFirstAndSecondAreDiffrentSizesTest()

    On Error GoTo ErrHandler
    Const MethodName = "SequenceEqualRetrunsFalseWhenFirstAndSecondAreDiffrentSizesTest"

    ' Arrange
    Dim First As Collection
    Set First = CollectionExt.Make(1, 2, 3)
    
    Dim Second As Collection
    Set Second = CollectionExt.Make(1, 2)

    ' Act & Assert
    ExUnit.IsFalse CollectionExt.SequenceEqual(First, Second, New LongEqualityComparer), GetSig(MethodName)

    Exit Sub
ErrHandler:
    Lapis.ExUnit.TestFailRunTime GetSig(MethodName)

End Sub


Private Sub SequenceEqualReturnsArgumentNullErrorWhenFirstIsNothingTest()

    On Error GoTo ErrHandler
    Const ExpectedError As Long = ErrorCode.ArgumentNull
    Const MethodName = "SequenceEqualReturnsArgumentNullErrorWhenFirstIsNothingTest"

    ' Act
    CollectionExt.SequenceEqual Nothing, CollectionExt.Make(1, 2), New LongEqualityComparer
    
    ' Assert
ErrHandler:
    Lapis.ExUnit.IsException ExpectedError, Err.Number, GetSig(MethodName)

End Sub


Private Sub SequenceEqualReturnsArgumentNullErrorWhenSecondIsNothingTest()

    On Error GoTo ErrHandler
    Const ExpectedError As Long = ErrorCode.ArgumentNull
    Const MethodName = "SequenceEqualReturnsArgumentNullErrorWhenSecondIsNothingTest"

    ' Act
    CollectionExt.SequenceEqual CollectionExt.Make(1, 2), Nothing, New LongEqualityComparer
    
    ' Assert
ErrHandler:
    Lapis.ExUnit.IsException ExpectedError, Err.Number, GetSig(MethodName)

End Sub


Private Sub SequenceEqualReturnsArgumentNullErrorWhenComparerIsNothingTest()

    On Error GoTo ErrHandler
    Const ExpectedError As Long = ErrorCode.ArgumentNull
    Const MethodName = "SequenceEqualReturnsArgumentNullErrorWhenComparerIsNothingTest"

    ' Act
    CollectionExt.SequenceEqual CollectionExt.Make(1, 2), CollectionExt.Make(1, 2), Nothing
    
    ' Assert
ErrHandler:
    Lapis.ExUnit.IsException ExpectedError, Err.Number, GetSig(MethodName)

End Sub


Private Sub IndexOfTest()

    On Error GoTo ErrHandler
    Const MethodName = "IndexOfTest"

    ' Act & Assert
    Lapis.ExUnit.AreEqual 2, CollectionExt.IndexOf(CollectionExt.Make(1, 2, 3, 4), 2), GetSig(MethodName)
    Lapis.ExUnit.AreEqual 1, CollectionExt.IndexOf(CollectionExt.Make(1, 2, 3, 4), 1), GetSig(MethodName)
    Lapis.ExUnit.AreEqual 4, CollectionExt.IndexOf(CollectionExt.Make(1, 2, 3, 4), 4), GetSig(MethodName)

    Exit Sub
ErrHandler:
    Lapis.ExUnit.TestFailRunTime GetSig(MethodName)

End Sub


Private Sub IndexOfReturnsValueWhenItemNotExistsInSourceTest()

    On Error GoTo ErrHandler
    Const MethodName = "IndexOfReturnsValueWhenItemNotExistsInSourceTest"

    ' Act & Assert
    Lapis.ExUnit.AreEqual -1, CollectionExt.IndexOf(CollectionExt.Make(1, 2, 3, 4), 10), GetSig(MethodName)
    
    Exit Sub
ErrHandler:
    Lapis.ExUnit.TestFailRunTime GetSig(MethodName)

End Sub


Private Sub IndexOfReturnsValueWhenItemNotDifferentTypeThanSourceTest()

    On Error GoTo ErrHandler
    Const MethodName = "IndexOfReturnsValueWhenItemNotDifferentTypeThanSourceTest"
    
    ' Act & Assert
    Lapis.ExUnit.AreEqual -1, CollectionExt.IndexOf(CollectionExt.Make(1, 2, 3, 4), "a"), GetSig(MethodName)
    
    Exit Sub
ErrHandler:
    Lapis.ExUnit.TestFailRunTime GetSig(MethodName)

End Sub


Private Sub IndexOfReturnsMinusOneWhenItemCantBeComparedTest()

    On Error GoTo ErrHandler
    Const MethodName = "IndexOfReturnsMinusOneWhenItemCantBeComparedTest"
    
    ' Assert
    Lapis.ExUnit.AreEqual -1, CollectionExt.IndexOf(CollectionExt.Make(1, 2, 3, 4), ThisWorkbook), GetSig(MethodName)
    
    Exit Sub
ErrHandler:
    Lapis.ExUnit.TestFailRunTime GetSig(MethodName)

End Sub


Private Sub IndexOfReturnsValueWhenItemImplementsIComparableTest()

    On Error GoTo ErrHandler
    Const MethodName = "IndexOfReturnsValueWhenItemImplementsIComparableTest"

    ' Act
    Dim Actual As Long
    Actual = CollectionExt.IndexOf(CollectionExt.Make(MakeStone(1, 2), MakeStone(2, 3)), MakeStone(1, 2))
    
    ' Assert
    Lapis.ExUnit.AreEqual 1, Actual, GetSig(MethodName)
    
    Exit Sub
ErrHandler:
    Lapis.ExUnit.TestFailRunTime GetSig(MethodName)

End Sub


Private Sub AllReturnsFalseWhenSourceDoesNotSatisfyConditionTest()

    On Error GoTo ErrHandler
    Const MethodName = "AllReturnsFalseWhenSourceDoesNotSatisfyConditionTest"

    ' Arrange
    Dim Expr As ICallable
    Set Expr = Lambda.Create("$1 = 1")
    
    ' Act
    Dim Actual As Boolean
    Actual = CollectionExt.All(CollectionExt.Make(1, 2, 3), Expr)
    
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
    Actual = CollectionExt.All(CollectionExt.Make(1, 1, 1), Expr)
    
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
    CollectionExt.All Nothing, Lambda.Create("$1 = 1")
    
    ' Assert
ErrHandler:
    Lapis.ExUnit.IsException ExpectedError, Err.Number, GetSig(MethodName)

End Sub


Private Sub AllReturnsArgumentNullErrorWhenPreidcateIsNothingTest()

    On Error GoTo ErrHandler
    Const ExpectedError As Long = ErrorCode.ArgumentNull
    Const MethodName = "AllReturnsArgumentNullErrorWhenPreidcateIsNothingTest"

    ' Act
    CollectionExt.All New Collection, Nothing
    
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
    Actual = CollectionExt.All(New Collection, Expr)
    
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
    Actual = CollectionExt.All(CollectionExt.Make(1, Nothing, 1), Expr)
    
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
    Actual = CollectionExt.All(CollectionExt.Make(New TestStone, Nothing), _
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
    Set Actual = CollectionExt.Where(CollectionExt.Make(1, 2, 3), Lambda.Create("$1 = 2"))
    
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
    CollectionExt.Where Nothing, Lambda.Create("$1 = 2")

    ' Assert
ErrHandler:
    Lapis.ExUnit.IsException ExpectedError, Err.Number, GetSig(MethodName)

End Sub


Private Sub WhereReturnsArgumentNullErrorWhenPredicateIsNothingTest()

    On Error GoTo ErrHandler
    Const ExpectedError As Long = ErrorCode.ArgumentNull
    Const MethodName = "WhereReturnsArgumentNullErrorWhenPredicateIsNothingTest"

    ' Act
    CollectionExt.Where New Collection, Nothing
    
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
    Set Actual = CollectionExt.Where(CollectionExt.Make(Stn1, Stn2), _
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
    Set Actual = CollectionExt.Where(CollectionExt.Make(Stn1, Nothing), _
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
    ExUnit.IsTrue CollectionExt.Some(CollectionExt.Make(1, 2, 3)), GetSig(MethodName)

    Exit Sub
ErrHandler:
    Lapis.ExUnit.TestFailRunTime GetSig(MethodName)

End Sub


Private Sub SomeReturnsFalseWhenPredicateIsNothingAndSourceIsEmptyTest()

    On Error GoTo ErrHandler
    Const MethodName = "SomeReturnsFalseWhenPredicateIsNothingAndSourceIsEmptyTest"

    ' Act & Assert
    ExUnit.IsFalse CollectionExt.Some(New Collection), GetSig(MethodName)

    Exit Sub
ErrHandler:
    Lapis.ExUnit.TestFailRunTime GetSig(MethodName)

End Sub


Private Sub SomeReturnsTrueWhenAtleastOneItemSatisfyConditionTest()

    On Error GoTo ErrHandler
    Const MethodName = "SomeReturnsTrueWhenAtleastOneItemSatisfyConditionTest"
    
    ' Act
    Dim Actual As Boolean
    Actual = CollectionExt.Some(CollectionExt.Make(1, 2, 3), Lambda.Create("$1 = 2"))

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
    Actual = CollectionExt.Some(CollectionExt.Make(1, 2, 3), Lambda.Create("$1 = 0"))

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
    ExUnit.IsFalse CollectionExt.Some(New Collection, Lambda.Create("$1 = 0")), GetSig(MethodName)

    Exit Sub
ErrHandler:
    Lapis.ExUnit.TestFailRunTime GetSig(MethodName)

End Sub


Private Sub SomeReturnsArgumentNullErrorWhenSourceIsNothingTest()

    On Error GoTo ErrHandler
    Const ExpectedError As Long = ErrorCode.ArgumentNull
    Const MethodName = "SomeReturnsArgumentNullErrorWhenSourceIsNothingTest"

    ' Act
    CollectionExt.Some Nothing

    ' Assert
ErrHandler:
    Lapis.ExUnit.IsException ExpectedError, Err.Number, GetSig(MethodName)

End Sub


Private Sub SumReturnsValueWhenSourceContainsIntegersTest()

    On Error GoTo ErrHandler
    Const MethodName = "SumReturnsValueWhenSourceContainsIntegersTest"
    
    ' Act
    Dim Actual As Variant
    Actual = CollectionExt.Sum(CollectionExt.Make(1, 2, 3))

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
    ExUnit.AreEqual 2, CollectionExt.Sum(Source, Lambda.Create("$1.Count")), GetSig(MethodName)

    Exit Sub
ErrHandler:
    Lapis.ExUnit.TestFailRunTime GetSig(MethodName)

End Sub


Private Sub SumReturnsValueWhenSourceIsEmptyTest()

    On Error GoTo ErrHandler
    Const MethodName = "SumReturnsValueWhenSourceIsEmptyTest"

    ' Act
    Dim Actual As Variant
    Actual = CollectionExt.Sum(CollectionExt.Make(Nothing))
    Actual = CollectionExt.Sum(CollectionExt.Make(Nothing), Lambda.Create("$1.Count"))

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
    Actual = CollectionExt.Sum(CollectionExt.Make(1, Nothing, 2))

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
    CollectionExt.Sum Nothing
    
    ' Assert
ErrHandler:
    Lapis.ExUnit.IsException ExpectedError, Err.Number, GetSig(MethodName)

End Sub


Private Sub AverageReturnsValueWhenSourceContainsIntegersTest()

    On Error GoTo ErrHandler
    Const MethodName = "AverageReturnsValueWhenSourceContainsIntegersTest"

    ' Act
    Dim Actual As Variant
    Actual = CollectionExt.Average(CollectionExt.Make(1, 2, 3))
    
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
    CollectionExt.Average Nothing

    ' Assert
ErrHandler:
    Lapis.ExUnit.IsException ExpectedError, Err.Number, GetSig(MethodName)

End Sub


Private Sub AverageReturnsValueWhenSourceIsEmptyTest()

    On Error GoTo ErrHandler
    Const MethodName = "AverageReturnsValueWhenSourceIsEmptyTest"

    ' Act
    Dim Actual As Variant
    Actual = CollectionExt.Average(New Collection)
    
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
    Actual = CollectionExt.Average(Source)
    
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
    Actual = CollectionExt.Average(CollectionExt.Make(Nothing, Nothing))

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
    Actual = CollectionExt.Average(CollectionExt.Make(Stn1, Stn2), Lambda.Create("$1.Age"))

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
    ExUnit.AreEqual 3, CollectionExt.Count(Source), GetSig(MethodName)
    ExUnit.AreEqual 2, CollectionExt.Count(Source, Lambda.Create("$1 >= 2")), GetSig(MethodName)

    Exit Sub
ErrHandler:
    Lapis.ExUnit.TestFailRunTime GetSig(MethodName)

End Sub


Private Sub CountReturnsArgumentNullErrorWhenSourceIsNothingTest()

    On Error GoTo ErrHandler
    Const ExpectedError As Long = ErrorCode.ArgumentNull
    Const MethodName = "CountReturnsArgumentNullErrorWhenSourceIsNothingTest"
    
    ' Act
    CollectionExt.Count Nothing

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
    Set Actual = CollectionExt.Convert(CollectionExt.Make("a", "ab", "abc"), Expr)
    
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
    Lapis.CollectionExt.Convert Nothing, Lambda.Create("len($1)")

    ' Assert
ErrHandler:
    ExUnit.IsException ExpectedError, Err.Number, GetSig(MethodName)

End Sub


Private Sub ConvertReturnsArgumentNullWhenSelectorIsNothingTest()

    On Error GoTo ErrHandler
    Const ExpectedError As Long = ErrorCode.ArgumentNull
    Const MethodName = "ConvertReturnsArgumentNullWhenSelectorIsNothingTest"

    ' Act
    Lapis.CollectionExt.Convert New Collection, Nothing

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

    ExUnit.AreEqual 1, CollectionExt.First(Source), GetSig(MethodName)
    ExUnit.AreEqual 3, CollectionExt.First(Source, Predicate), GetSig(MethodName)
    
    ' Reference Types
    Dim Stn1 As New TestStone: Stn1.Age = 1: Stn1.Weight = 10
    Dim Stn2 As New TestStone: Stn2.Age = 2: Stn2.Weight = 20
    Dim Stn3 As New TestStone: Stn3.Age = 1: Stn3.Weight = 30

    Set Source = CollectionExt.Make(Stn1, Stn2, Stn3)
    Set Predicate = Lambda.Create("$1.Age = 1")

    ' Act & Assert
    ExUnit.AreEqual 10, CollectionExt.First(Source).Weight, GetSig(MethodName)
    ExUnit.AreEqual 10, CollectionExt.First(Source, Lambda.Create("$1.Age = 1")).Weight, GetSig(MethodName)

    Exit Sub
ErrHandler:
    Lapis.ExUnit.TestFailRunTime GetSig(MethodName)

End Sub


Private Sub FirstReturnsArgumentNullErrorWhenSourceIsNothingTest()

    On Error GoTo ErrHandler
    Const ExpectedError As Long = ErrorCode.ArgumentNull
    Const MethodName = "FirstReturnsArgumentNullErrorWhenSourceIsNothingTest"

    ' Act
    CollectionExt.First Nothing

    ' Assert
ErrHandler:
    Lapis.ExUnit.IsException ExpectedError, Err.Number, GetSig(MethodName)

End Sub


Private Sub FirstReturnsInvalidOperationErrorWhenSourceDoesNotSatisfyConditionTest()

    On Error GoTo ErrHandler
    Const ExpectedError As Long = ErrorCode.InvalidOperation
    Const MethodName = "FirstReturnsInvalidOperationErrorWhenSourceDoesNotSatisfyConditionTest"

    ' Act
    CollectionExt.First CollectionExt.Make(1, 2, 3), Lambda.Create("$1 = 5")

    ' Assert
ErrHandler:
    Lapis.ExUnit.IsException ExpectedError, Err.Number, GetSig(MethodName)

End Sub


Private Sub FirstReturnsInvalidOperationErrorWhenSourceIsEmptyTest()

    On Error GoTo ErrHandler
    Const ExpectedError As Long = ErrorCode.InvalidOperation
    Const MethodName = "FirstReturnsInvalidOperationErrorWhenSourceIsEmptyTest"

    ' Act
    CollectionExt.First New Collection

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
    ExUnit.AreEqual 30, CollectionExt.Last(Source).Weight, GetSig(MethodName)
    ExUnit.AreEqual 30, CollectionExt.Last(Source, Lambda.Create("$1.Age = 1")).Weight, GetSig(MethodName)
    
    Exit Sub
ErrHandler:
    Lapis.ExUnit.TestFailRunTime GetSig(MethodName)

End Sub


Private Sub LastReturnsArgumentNullErrorWhenSourceIsNothingTest()

    On Error GoTo ErrHandler
    Const ExpectedError As Long = ErrorCode.ArgumentNull
    Const MethodName = "LastReturnsArgumentNullErrorWhenSourceIsNothingTest"

    ' Act
    CollectionExt.Last Nothing

    ' Assert
ErrHandler:
    Lapis.ExUnit.IsException ExpectedError, Err.Number, GetSig(MethodName)

End Sub


Private Sub LastReturnsInvalidOperationErrorWhenSourceDoesNotSatisfyConditionTest()

    On Error GoTo ErrHandler
    Const ExpectedError As Long = ErrorCode.InvalidOperation
    Const MethodName = "LastReturnsInvalidOperationErrorWhenSourceDoesNotSatisfyConditionTest"

    ' Act
    CollectionExt.Last CollectionExt.Make(1, 2, 3), Lambda.Create("$1 = 5")

    ' Assert
ErrHandler:
    Lapis.ExUnit.IsException ExpectedError, Err.Number, GetSig(MethodName)

End Sub


Private Sub LastReturnsInvalidOperationErrorWhenSourceIsEmptyTest()

    On Error GoTo ErrHandler
    Const ExpectedError As Long = ErrorCode.InvalidOperation
    Const MethodName = "LastReturnsInvalidOperationErrorWhenSourceIsEmptyTest"

    ' Act
    CollectionExt.Last New Collection

    ' Assert
ErrHandler:
    Lapis.ExUnit.IsException ExpectedError, Err.Number, GetSig(MethodName)

End Sub



Private Sub SelectOneReturnsValueWhenSourceDoesSatisfyConditionTest()

    On Error GoTo ErrHandler
    Const MethodName = "SelectOneReturnsValueWhenSourceDoesSatisfyConditionTest"
    
    ' Act & Assert
    ExUnit.AreEqual 1, CollectionExt.SelectOne(CollectionExt.Make(1)), GetSig(MethodName)
    ExUnit.AreEqual 2, CollectionExt.SelectOne(CollectionExt.Make(1, 2, 3), Lambda.Create("$1 = 2")), GetSig(MethodName)

    Exit Sub
ErrHandler:
    Lapis.ExUnit.TestFailRunTime GetSig(MethodName)

End Sub


Private Sub SelectOneReturnsArgumentNullErrorWhenSourceIsNothingTest()

    On Error GoTo ErrHandler
    Const ExpectedError As Long = ErrorCode.ArgumentNull
    Const MethodName = "SelectOneReturnsArgumentNullErrorWhenSourceIsNothingTest"

    ' Act
    CollectionExt.SelectOne Nothing

    ' Assert
ErrHandler:
    Lapis.ExUnit.IsException ExpectedError, Err.Number, GetSig(MethodName)

End Sub


Private Sub SelectOneReturnsInvalidOperationErrorWhenSourceDoesNotSatisfyConditionTest()

    On Error GoTo ErrHandler
    Const ExpectedError As Long = ErrorCode.InvalidOperation
    Const MethodName = "SelectOneReturnsInvalidOperationErrorWhenSourceDoesNotSatisfyConditionTest"

    ' Act
    CollectionExt.SelectOne CollectionExt.Make(1, 2, 3), Lambda.Create("$1 = 5")

    ' Assert
ErrHandler:
    Lapis.ExUnit.IsException ExpectedError, Err.Number, GetSig(MethodName)

End Sub


Private Sub SelectOneReturnsInvalidOperationErrorWhenSourceIsEmptyTest()

    On Error GoTo ErrHandler
    Const ExpectedError As Long = ErrorCode.InvalidOperation
    Const MethodName = "SelectOneReturnsInvalidOperationErrorWhenSourceIsEmptyTest"

    ' Act
    CollectionExt.SelectOne New Collection

ErrHandler:
    Lapis.ExUnit.IsException ExpectedError, Err.Number, GetSig(MethodName)

End Sub


Private Sub SelectOneReturnsInvalidOperationErrorWhenSourceDoesSatisfyConditionMultipleTimesTest()

    On Error GoTo ErrHandler
    Const ExpectedError As Long = ErrorCode.InvalidOperation
    Const MethodName = "SelectOneReturnsInvalidOperationErrorWhenSourceDoesSatisfyConditionMultipleTimesTest"

    ' Act
    ExUnit.AreEqual 3, CollectionExt.SelectOne(CollectionExt.Make(1, 2, 3), Lambda.Create("$1 >= 2")), GetSig(MethodName)
    
    ' Assert
ErrHandler:
    Lapis.ExUnit.IsException ExpectedError, Err.Number, GetSig(MethodName)

End Sub

