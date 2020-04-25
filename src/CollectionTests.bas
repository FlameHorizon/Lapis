Attribute VB_Name = "CollectionTests"
'@Folder("Tests")
Option Explicit

Private Const ModuleName As String = "CollectionTests"


Public Sub Start()

    ToStringValueTypesTest
    ToStringReturnsArgumentNullErrorWhenSourceIsNothingTest
    ToStringReturnsArgumentNullErrorWhenConverterIsNothingTest
    ToStringReturnsEmptyStringWhenSourceIsEmptyTest
    ToStringReturnsStringWithDefinedDelimeterTest
    ToStringReturnsStringWhenDelimterIsNullStringTest
    ToStringReturnsStringWhenItemsAreNumbersTest
    ToStringReturnsStringWhenItemsAreMixedValueTypesTest
    ToStringReturnsStringWhenItemsContainsObjectsTest
    ToStringReturnsInvalidOperationWhenConverterCantConvertNonObjectTest
    ToStringReturnsInvalidOperationWhenConverterCantConvertDiffObjectTest
    ToStringReturnsInvalidOperationWhenConverterCantConvertNothingTest
    
    ToStringByPropertyTest
    ToStringByPropertyReturnsArgumentOutOfRangeErrorWhenPropertyDoesNotExistsTest
    
    DistinctTest
    DistinctReturnsItemWhenSingleItemGivenTest
    DistinctReturnsEmptyCollectionWhenSourceIsEmptyTest
    DistinctReturnsArgumentNullErrorWhenSourceIsNothingTest
    DistinctReturnsArgumentNullErrorWhenComparerIsNothingTest
    
    ExceptTest
    ExceptReturnsEmptyCollectionWhenFirstAndSecondAreEmptyTest
    ExceptReturnsEmptyCollectionWhenFirstIsEmptyTest
    ExceptReturnsArgumentNullErrorIfFirstIsNothingTest
    ExceptReturnsArgumentNullErrorIfSecondIsNothingTest
    ExceptReturnsArgumentNullErrorIfComparerIsNothingTest
    
    IntersectTest
    IntersectReturnsEmptyCollectionWhenFirstAndSecondAreEmptyTest
    IntersectReturnsEmptyCollectionWhenFirstIsEmpty
    IntersectReturnsArgumentNullErrorIfFirstIsNothingTest
    IntersectReturnsArgumentNullErrorIfSecondIsNothingTest
    IntersectReturnsArgumentNullErrorIfComparerIsNothingTest
    
    MinWhenSourceContainsValueTypesTest
    MinWhenSourceContainsReferencedTypesTest
    MinReturnsArgumentNullErrorWhenSourceIsNothingTest
    MinReturnsArgumentNullErrorWhenComparerIsNothingTest
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
    
    SumReturnsValueWhenSourceContainsIntegersTest
    SumReturnsValueWhenSourceContainsReferencedTypesTest
    SumReturnsValueWhenSourceIsEmptyTest
    SumReturnsValueWhenSourceContainsIntegersAndNothingTest
    SumReturnsArgumentNullErrorWhenSourceIsNothingTest
    SumReturnsArgumentNullErrorWhenSelectorIsNothingTest
    
    AverageReturnsValueWhenSourceContainsIntegersTest
    AverageReturnsValueWhenSourceContainsReferencedTypesTest
    AverageReturnsArgumentNullErrorWhenSourceIsEmptyTest
    AverageReturnsArgumentNullErrorWhenSelectorIsEmptyTest
    AverageReturnsValueWhenSourceIsEmptyTest
    AverageReturnsValueWhenSourceContainsIntegersAndNothingTest
    AverageReturnsValueWhenSourceContainsOnlyNothingTest
    
    TakeTest
    TakeReturnsArgumentNullErrorWhenSourceIsNothingTest
    TakeReturnsEmptyCollectionWhenCountIsZeroTest
    TakeReturnsEmptyCollectionWhenCountIsNegativeTest
    TakeReturnsEmptyCollectionWhenSourceIsEmptyTest
    
    AllReturnsFalseWhenSourceDoesNotSatisfyConditionTest
    AllReturnsTrueWhenSourceDoesSatisfyConditionTest
    AllReturnsArgumentNullErrorWhenSourceIsNothingTest
    AllReturnsArgumentNullErrorWhenPreidcateIsNothingTest
    AllReturnsTrueWhenSourceIsEmptyTest
    AllReturnsFalseWhenSourceSatisfyConditionAndHasNothingTest
    
    SomeReturnsTrueWhenPredicateIsNothingAndSourceIsNotEmptyTest
    SomeReturnsFalseWhenPredicateIsNothingAndSourceIsEmptyTest
    SomeReturnsTrueWhenAtleastOneItemSatisfyConditionTest
    SomeReturnsFalseWhenNoItemSatisfyConditionTest
    SomeReturnsFalseWhenSourceIsEmptyTest
    SomeReturnsArgumentNullErrorWhenSourceIsNothingTest
    
    SkipTest
    SkipReturnsArgumentNullErrorWhenSourceIsNothingTest
    SkipReturnsEntireSourceWhenCountIsZeroOrNegativeTest
    SkinReturnsEmptyCollectionWhenCountIsGreaterThanNumberOfItemsInSourceTest
    
    SequenceEqualReturnsTrueWhenFirstAndSecondContainsEqualElementsTest
    SequenceEqualReturnsFalseWhenFirstAndSecondContainsNotEqualElementsTest
    SequenceEqualRetrunsFalseWhenFirstAndSecondAreDiffrentSizesTest
    SequenceEqualReturnsArgumentNullErrorWhenFirstIsNothingTest
    SequenceEqualReturnsArgumentNullErrorWhenSecondIsNothingTest
    SequenceEqualReturnsArgumentNullErrorWhenComparerIsNothingTest
    
    FirstReturnsValueWhenSourceDoesSatisfyConditionTest
    FirstReturnsArgumentNullErrorWhenSourceIsNothingTest
    FirstReturnsArgumentNullErrorWhenPredicateIsNothingTest
    FirstReturnsInvalidOperationErrorWhenSourceDoesNotSatisfyConditionTest
    FirstReturnsInvalidOperationErrorWhenSourceIsEmptyTest
    
    LastReturnsValueWhenSourceDoesSatisfyConditionTest
    LastReturnsArgumentNullErrorWhenSourceIsNothingTest
    LastReturnsArgumentNullErrorWhenPredicateIsNothingTest
    LastReturnsInvalidOperationErrorWhenSourceDoesNotSatisfyConditionTest
    LastReturnsInvalidOperationErrorWhenSourceIsEmptyTest
    
    SelectOneReturnsValueWhenSourceDoesSatisfyConditionTest
    SelectOneReturnsArgumentNullErrorWhenSourceIsNothingTest
    SelectOneReturnsArgumentNullErrorWhenPredicateIsNothingTest
    SelectOneReturnsInvalidOperationErrorWhenSourceDoesNotSatisfyConditionTest
    SelectOneReturnsInvalidOperationErrorWhenSourceIsEmptyTest
    SelectOneReturnsInvalidOperationErrorWhenSourceDoesSatisfyConditionMultipleTimesTest
    
    CountTest
    CountReturnsArgumentNullErrorWhenSourceIsNothingTest
    CountReturnsArgumentNullErrorWhenPredicateIsNothingTest
    
    WhereTest
    WhereReturnsArgumentNullErrorWhenSourceIsNothingTest
    WhereReturnsArgumentNullErrorWhenPredicateIsNothingTest
    
End Sub


Private Sub ToStringValueTypesTest()

    On Error GoTo ErrHandler
    Const MethodName = "ToStringValueTypesTest"
    
    ' Act
    Dim Actual As String
    Actual = CollectionExt.ToString(CollectionExt.Make("abc", "cba"), New ValueTypeToStringConverter)

    ' Assert
    Lapis.ExUnit.AreEqual "abc,cba", Actual, GetSig(MethodName)

    Exit Sub
ErrHandler:
    Lapis.ExUnit.TestFailRunTime GetSig(MethodName)

End Sub


Private Function GetSig(ByVal MethodName As String) As String
    GetSig = ModuleName & "." & MethodName
End Function


Private Sub ToStringReturnsArgumentNullErrorWhenSourceIsNothingTest()

    On Error GoTo ErrHandler
    Const ExpectedError As Long = ErrorCode.ArgumentNull
    Const MethodName = "ToStringReturnsArgumentNullErrorWhenSourceIsNothingTest"

    ' Act
    CollectionExt.ToString Nothing, New ValueTypeToStringConverter
    
    ' Assert
ErrHandler:
    Lapis.ExUnit.IsException ExpectedError, Err.Number, GetSig(MethodName)

End Sub


Private Sub ToStringReturnsArgumentNullErrorWhenConverterIsNothingTest()

    On Error GoTo ErrHandler
    Const ExpectedError As Long = ErrorCode.ArgumentNull
    Const MethodName = "ToStringReturnsArgumentNullErrorWhenConverterIsNothingTest"

    ' Act
    CollectionExt.ToString New Collection, Nothing
    
    ' Assert
ErrHandler:
    Lapis.ExUnit.IsException ExpectedError, Err.Number, GetSig(MethodName)

End Sub


Private Sub ToStringReturnsEmptyStringWhenSourceIsEmptyTest()

    On Error GoTo ErrHandler
    Const MethodName = "ToStringReturnsEmptyStringWhenSourceIsEmptyTest"
    
    ' Act
    Dim Actual As String
    Actual = CollectionExt.ToString(New Collection, New ValueTypeToStringConverter)

    ' Assert
    ExUnit.AreEqual vbNullString, Actual, GetSig(MethodName)

    Exit Sub
ErrHandler:
    Lapis.ExUnit.TestFailRunTime GetSig(MethodName)

End Sub


Private Sub ToStringReturnsStringWithDefinedDelimeterTest()

    On Error GoTo ErrHandler
    Const MethodName = "ToStringReturnsStringWithDefinedDelimeterTest"

    ' Act
    Dim Actual As String
    Actual = CollectionExt.ToString(CollectionExt.Make("abc", "cba"), New ValueTypeToStringConverter, "-")
    
    ' Assert
    ExUnit.AreEqual "abc-cba", Actual, GetSig(MethodName)

    Exit Sub
ErrHandler:
    Lapis.ExUnit.TestFailRunTime GetSig(MethodName)

End Sub


Private Sub ToStringReturnsStringWhenDelimterIsNullStringTest()

    On Error GoTo ErrHandler
    Const MethodName = "ToStringReturnsStringWhenDelimterIsNullStringTest"

    ' Act
    Dim Actual As String
    Actual = CollectionExt.ToString(CollectionExt.Make("abc", "cba"), New ValueTypeToStringConverter, vbNullString)
    
    ' Assert
    ExUnit.AreEqual "abccba", Actual, GetSig(MethodName)


    Exit Sub
ErrHandler:
    Lapis.ExUnit.TestFailRunTime GetSig(MethodName)

End Sub


Private Sub ToStringReturnsStringWhenItemsAreNumbersTest()

    On Error GoTo ErrHandler
    Const MethodName = "ToStringReturnsStringWhenItemsAreNumbersTest"

    ' Act
    Dim Actual As String
    Actual = CollectionExt.ToString(CollectionExt.Make(1, 3, 4), New ValueTypeToStringConverter)

    ' Assert
    ExUnit.AreEqual "1,3,4", Actual, GetSig(MethodName)

    Exit Sub
ErrHandler:
    Lapis.ExUnit.TestFailRunTime GetSig(MethodName)

End Sub


Private Sub ToStringReturnsStringWhenItemsAreMixedValueTypesTest()

    On Error GoTo ErrHandler
    Const MethodName = "ToStringReturnsStringWhenItemsAreMixedValueTypesTest"

    ' Arrange
    Dim Source As Collection
    Set Source = CollectionExt.Make("a", 1, 2.34, #2/2/2020#, &H5CC)
    
    ' Act
    Dim Actual As String
    Actual = CollectionExt.ToString(Source, New ValueTypeToStringConverter)

    ' Assert
    ExUnit.AreEqual "a,1,2.34,2/2/2020,1484", Actual, GetSig(MethodName)

    Exit Sub
ErrHandler:
    Lapis.ExUnit.TestFailRunTime GetSig(MethodName)

End Sub


Private Sub ToStringReturnsStringWhenItemsContainsObjectsTest()

    On Error GoTo ErrHandler
    Const MethodName = "ToStringReturnsStringWhenItemsContainsObjectsTest"

    ' Act
    Dim Actual As String
    Actual = CollectionExt.ToString(CollectionExt.Make(ThisWorkbook, ThisWorkbook), New WorkbookToStringConverter)

    ' Assert
    ExUnit.AreEqual "Lapis.xlam,Lapis.xlam", Actual, GetSig(MethodName)

    Exit Sub
ErrHandler:
    Lapis.ExUnit.TestFailRunTime GetSig(MethodName)

End Sub


Private Sub ToStringReturnsInvalidOperationWhenConverterCantConvertNonObjectTest()

    On Error GoTo ErrHandler
    Const ExpectedError As Long = ErrorCode.InvalidOperation
    Const MethodName = "ToStringReturnsInvalidOperationWhenConverterCantConvertNonObjectTest"

    ' Act
    Dim Actual As String
    Actual = CollectionExt.ToString(CollectionExt.Make("a"), New WorkbookToStringConverter)

    ' Assert
ErrHandler:
    Lapis.ExUnit.IsException ExpectedError, Err.Number, GetSig(MethodName)

End Sub


Private Sub ToStringReturnsInvalidOperationWhenConverterCantConvertDiffObjectTest()

    On Error GoTo ErrHandler
    Const ExpectedError As Long = ErrorCode.InvalidOperation
    Const MethodName = "ToStringReturnsInvalidOperationWhenConverterCantConvertDiffObjectTest"

    ' Act
    Dim Actual As String
    Actual = CollectionExt.ToString(CollectionExt.Make(New Collection), New WorkbookToStringConverter)

    ' Assert
ErrHandler:
    Lapis.ExUnit.IsException ExpectedError, Err.Number, GetSig(MethodName)

End Sub


Private Sub ToStringReturnsInvalidOperationWhenConverterCantConvertNothingTest()

    On Error GoTo ErrHandler
    Const ExpectedError As Long = ErrorCode.InvalidOperation
    Const MethodName = "ToStringReturnsInvalidOperationWhenConverterCantConvertNothingTest"
    
    ' Act
    Dim Actual As String
    Actual = CollectionExt.ToString(CollectionExt.Make(Nothing), New WorkbookToStringConverter)

    ' Assert
ErrHandler:
    Lapis.ExUnit.IsException ExpectedError, Err.Number, GetSig(MethodName)

End Sub


Private Sub ToStringByPropertyTest()

    On Error GoTo ErrHandler
    Const MethodName = "ToStringByPropertyTest"

    ' Act
    Dim Actual As String
    Actual = CollectionExt.ToStringByProperty(CollectionExt.Make(ThisWorkbook, ThisWorkbook), "Name")

    ' Assert
    ExUnit.AreEqual "Lapis.xlam,Lapis.xlam", Actual, GetSig(MethodName)

    Exit Sub
ErrHandler:
    Lapis.ExUnit.TestFailRunTime GetSig(MethodName)

End Sub


Private Sub ToStringByPropertyReturnsArgumentOutOfRangeErrorWhenPropertyDoesNotExistsTest()

    On Error GoTo ErrHandler
    Const ExpectedError As Long = ErrorCode.ArgumentOutOfRange
    Const MethodName = "ToStringByPropertyReturnsArgumentOutOfRangeErrorWhenPropertyDoesNotExistsTest"

    ' Act
    Dim Actual As String
    Actual = CollectionExt.ToStringByProperty(CollectionExt.Make(ThisWorkbook, ThisWorkbook), "NotExistingProperty")

    ' Assert
ErrHandler:
    Lapis.ExUnit.IsException ExpectedError, Err.Number, GetSig(MethodName)

End Sub


Private Sub DistinctTest()

    On Error GoTo ErrHandler
    Const MethodName = "DistinctTest"

    ' Act
    Dim Actual As Collection
    Set Actual = CollectionExt.Distinct(CollectionExt.Make(1, 2, 3, 2), New LongEqualityComparer)

    ' Assert
    ExUnit.AreEqual 3, Actual.Count, GetSig(MethodName)
    ExUnit.AreEqual 1, Actual.Item(1), GetSig(MethodName)
    ExUnit.AreEqual 2, Actual.Item(2), GetSig(MethodName)
    ExUnit.AreEqual 3, Actual.Item(3), GetSig(MethodName)

    Exit Sub
ErrHandler:
    Lapis.ExUnit.TestFailRunTime GetSig(MethodName)

End Sub


Private Sub DistinctReturnsItemWhenSingleItemGivenTest()

    On Error GoTo ErrHandler
    Const MethodName = "DistinctReturnsItemWhenSingleItemGivenTest"

    ' Act
    Dim Actual As Collection
    Set Actual = CollectionExt.Distinct(CollectionExt.Make(1), New LongEqualityComparer)
    
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
    Set Actual = CollectionExt.Distinct(New Collection, New LongEqualityComparer)
    
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
    CollectionExt.Distinct Nothing, New LongEqualityComparer
    
    ' Assert
ErrHandler:
    Lapis.ExUnit.IsException ExpectedError, Err.Number, GetSig(MethodName)

End Sub


Private Sub DistinctReturnsArgumentNullErrorWhenComparerIsNothingTest()

    On Error GoTo ErrHandler
    Const ExpectedError As Long = ErrorCode.ArgumentNull
    Const MethodName = "DistinctReturnsArgumentNullErrorWhenComparerIsNothingTest"
    
    ' Act
    CollectionExt.Distinct New Collection, Nothing
    
    ' Assert
ErrHandler:
    Lapis.ExUnit.IsException ExpectedError, Err.Number, GetSig(MethodName)

End Sub


Private Sub ExceptTest()

    On Error GoTo ErrHandler
    Const MethodName = "ExceptTest"
    
    ' Act
    Dim Actual As Collection
    Set Actual = CollectionExt.Except(CollectionExt.Make(1, 2, 3), CollectionExt.Make(1, 3), New LongEqualityComparer)
    
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


Private Sub ExceptReturnsArgumentNullErrorIfComparerIsNothingTest()

    On Error GoTo ErrHandler
    Const ExpectedError As Long = ErrorCode.ArgumentNull
    Const MethodName = "ExceptReturnsArgumentNullErrorIfComparerIsNothingTest"

    ' Act
    CollectionExt.Except New Collection, New Collection, Nothing
    
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
                                         CollectionExt.Make(2, 3), _
                                         New LongEqualityComparer)
    
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


Private Sub IntersectReturnsArgumentNullErrorIfComparerIsNothingTest()

    On Error GoTo ErrHandler
    Const ExpectedError As Long = ErrorCode.ArgumentNull
    Const MethodName = "IntersectReturnsArgumentNullErrorIfComparerIsNothingTest"

    ' Act
    CollectionExt.Intersect New Collection, New Collection, Nothing
    
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

    ' Act
    Dim Actual As Variant
    Actual = CollectionExt.Min(CollectionExt.Make(3, 2, 1), New LongComparer)
    
    ' Assert
    ExUnit.AreEqual 1, Actual, GetSig(MethodName)
    
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
    Set Actual = CollectionExt.Min(Source, New TestCollectionComparer)
    
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
    CollectionExt.Min Nothing, New LongComparer
    
    ' Assert
ErrHandler:
    Lapis.ExUnit.IsException ExpectedError, Err.Number, GetSig(MethodName)

End Sub


Private Sub MinReturnsArgumentNullErrorWhenComparerIsNothingTest()

    On Error GoTo ErrHandler
    Const ExpectedError As Long = ErrorCode.ArgumentNull
    Const MethodName = "MinReturnsArgumentNullErrorWhenComparerIsNothingTest"

    ' Act
    CollectionExt.Min New Collection, Nothing
    
    ' Assert
ErrHandler:
    Lapis.ExUnit.IsException ExpectedError, Err.Number, GetSig(MethodName)

End Sub


Private Sub MinReturnsLowestValueWhenNothingIsPresentTest()

    On Error GoTo ErrHandler
    Const MethodName = "MinReturnsLowestValueWhenNothingIsPresentTest"

    ' Arrange
    Dim Source As Collection
    Set Source = CollectionExt.Make(New Collection, Nothing, CollectionExt.Make(1))
    
    ' Act
    Dim Actual As Object
    Set Actual = CollectionExt.Min(Source, New TestCollectionComparer)

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
    Set Actual = CollectionExt.Min(CollectionExt.Make(Nothing), New TestCollectionComparer)
    
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
    CollectionExt.Min Nothing, New LongComparer
    
    ' Assert
ErrHandler:
    Lapis.ExUnit.IsException ExpectedError, Err.Number, GetSig(MethodName)

End Sub


Private Sub MaxReturnsArgumentNullErrorWhenComparerIsNothingTest()

    On Error GoTo ErrHandler
    Const ExpectedError As Long = ErrorCode.ArgumentNull
    Const MethodName = "MaxReturnsArgumentNullErrorWhenComparerIsNothingTest"

    ' Act
    CollectionExt.Min New Collection, Nothing
    
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


Private Sub SumReturnsValueWhenSourceContainsIntegersTest()

    On Error GoTo ErrHandler
    Const MethodName = "SumReturnsValueWhenSourceContainsIntegersTest"
    
    ' Act
    Dim Actual As Variant
    Actual = CollectionExt.Sum(CollectionExt.Make(1, 2, 3), New ValueTypeToNumericConverter)

    ' Assert
    ExUnit.AreEqual 6, Actual, GetSig(MethodName)
    
    Exit Sub
ErrHandler:
    Lapis.ExUnit.TestFailRunTime GetSig(MethodName)

End Sub


Private Sub SumReturnsValueWhenSourceContainsReferencedTypesTest()

    On Error GoTo ErrHandler
    Const MethodName = "SumReturnsValueWhenSourceContainsReferencedTypesTest"

    ' Arrange
    Dim Selector As New PropertyToNumericConverter
    Selector.PropertyName = "Count"
    Selector.CallType = VbCallType.VbMethod ' Collection.Count is not a property, it is a method.
    
    ' Act
    Dim Source As Collection
    Set Source = CollectionExt.Make(CollectionExt.Make(1), CollectionExt.Make(1))

    ' Assert
    ExUnit.AreEqual 2, CollectionExt.Sum(Source, Selector), GetSig(MethodName)

    Exit Sub
ErrHandler:
    Lapis.ExUnit.TestFailRunTime GetSig(MethodName)

End Sub


Private Sub SumReturnsValueWhenSourceIsEmptyTest()

    On Error GoTo ErrHandler
    Const MethodName = "SumReturnsValueWhenSourceIsEmptyTest"

    ' Act
    Dim Actual As Variant
    Actual = CollectionExt.Sum(CollectionExt.Make(Nothing), New ValueTypeToNumericConverter)

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
    Actual = CollectionExt.Sum(CollectionExt.Make(1, Nothing, 2), New ValueTypeToNumericConverter)

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
    CollectionExt.Sum Nothing, New ValueTypeToNumericConverter
    
    ' Assert
ErrHandler:
    Lapis.ExUnit.IsException ExpectedError, Err.Number, GetSig(MethodName)

End Sub


Private Sub SumReturnsArgumentNullErrorWhenSelectorIsNothingTest()

    On Error GoTo ErrHandler
    Const ExpectedError As Long = ErrorCode.ArgumentNull
    Const MethodName = "SumReturnsArgumentNullErrorWhenSelectorIsNothingTest"

    ' Act
    CollectionExt.Sum New Collection, Nothing

    ' Assert
ErrHandler:
    Lapis.ExUnit.IsException ExpectedError, Err.Number, GetSig(MethodName)

End Sub


Private Sub AverageReturnsValueWhenSourceContainsIntegersTest()

    On Error GoTo ErrHandler
    Const MethodName = "AverageReturnsValueWhenSourceContainsIntegersTest"

    ' Act
    Dim Actual As Variant
    Actual = CollectionExt.Average(CollectionExt.Make(1, 2, 3), New ValueTypeToNumericConverter)
    
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
    CollectionExt.Average Nothing, New ValueTypeToNumericConverter

    ' Assert
ErrHandler:
    Lapis.ExUnit.IsException ExpectedError, Err.Number, GetSig(MethodName)

End Sub


Private Sub AverageReturnsArgumentNullErrorWhenSelectorIsEmptyTest()

    On Error GoTo ErrHandler
    Const ExpectedError As Long = ErrorCode.ArgumentNull
    Const MethodName = "AverageReturnsArgumentNullErrorWhenSelectorIsEmptyTest"

    ' Act
    CollectionExt.Average New Collection, Nothing
    
    ' Assert
ErrHandler:
    Lapis.ExUnit.IsException ExpectedError, Err.Number, GetSig(MethodName)

End Sub


Private Sub AverageReturnsValueWhenSourceContainsReferencedTypesTest()

    On Error GoTo ErrHandler
    Const MethodName = "AverageReturnsValueWhenSourceContainsReferencedTypesTest"

    ' Arrange
    Dim Source As New Collection
    Source.Add CollectionExt.Make(1)
    Source.Add CollectionExt.Make(1, 2)
    
    Dim Selector As New PropertyToNumericConverter
    Selector.PropertyName = "Count"
    Selector.CallType = VbCallType.VbMethod ' Collection.Count is not a property, it is a method.
    
    ' Act
    Dim Actual As Variant
    Actual = CollectionExt.Average(Source, Selector)
    
    ' Assert
    ExUnit.AreEqual Round(1.5, 2), Round(Actual, 2), GetSig(MethodName)

    Exit Sub
ErrHandler:
    Lapis.ExUnit.TestFailRunTime GetSig(MethodName)

End Sub


Private Sub AverageReturnsValueWhenSourceIsEmptyTest()

    On Error GoTo ErrHandler
    Const MethodName = "AverageReturnsValueWhenSourceIsEmptyTest"

    ' Act
    Dim Actual As Variant
    Actual = CollectionExt.Average(New Collection, New ValueTypeToNumericConverter)
    
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
    Actual = CollectionExt.Average(Source, New ValueTypeToNumericConverter)
    
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
    Actual = CollectionExt.Average(CollectionExt.Make(Nothing, Nothing), New ValueTypeToNumericConverter)

    ' Assert
    ExUnit.AreEqual 0, Actual, GetSig(MethodName)

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


Private Sub AllReturnsFalseWhenSourceDoesNotSatisfyConditionTest()

    On Error GoTo ErrHandler
    Const MethodName = "AllReturnsFalseWhenSourceDoesNotSatisfyConditionTest"

    ' Arrange
    Dim Predicate As New Predicate
    Predicate.ComparisonValue = 1
    Predicate.Comparer = New LongComparer
    
    ' Act
    Dim Actual As Boolean
    Actual = CollectionExt.All(CollectionExt.Make(1, 2, 3), Predicate)
    
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
    Dim Predicate As New Predicate
    Predicate.Operator = ComparisonOperator.EqualTo
    Predicate.ComparisonValue = 1
    Predicate.Comparer = New LongComparer
    
    ' Act
    Dim Actual As Boolean
    Actual = CollectionExt.All(CollectionExt.Make(1, 1, 1), Predicate)
    
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
    CollectionExt.All Nothing, New Predicate
    
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
    Dim Predicate As New Predicate
    Predicate.ComparisonValue = 1
    Predicate.Comparer = New LongComparer
    
    ' Act
    Dim Actual As Boolean
    Actual = CollectionExt.All(New Collection, Predicate)
    
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
    Dim Predicate As New Predicate
    Predicate.ComparisonValue = 1
    Predicate.Comparer = New LongComparer
    
    ' Act
    Dim Actual As Boolean
    Actual = CollectionExt.All(CollectionExt.Make(1, Nothing, 1), Predicate)
    
    ' Assert
    ExUnit.IsFalse Actual, GetSig(MethodName)

    Exit Sub
ErrHandler:
    Lapis.ExUnit.TestFailRunTime GetSig(MethodName)

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

    ' Arrange
    Dim Predicate As New Predicate
    Predicate.Comparer = New LongComparer
    Predicate.ComparisonValue = 2

    ' Act
    Dim Actual As Boolean
    Actual = CollectionExt.Some(CollectionExt.Make(1, 2, 3), Predicate)

    ' Assert
    ExUnit.IsTrue Actual, GetSig(MethodName)

    Exit Sub
ErrHandler:
    Lapis.ExUnit.TestFailRunTime GetSig(MethodName)

End Sub


Private Sub SomeReturnsFalseWhenNoItemSatisfyConditionTest()

    On Error GoTo ErrHandler
    Const MethodName = "SomeReturnsFalseWhenNoItemSatisfyConditionTest"

    ' Arrange
    Dim Predicate As New Predicate
    Predicate.Comparer = New LongComparer
    Predicate.ComparisonValue = 0

    ' Act
    Dim Actual As Boolean
    Actual = CollectionExt.Some(CollectionExt.Make(1, 2, 3), Predicate)

    ' Assert
    ExUnit.IsFalse Actual, GetSig(MethodName)

    Exit Sub
ErrHandler:
    Lapis.ExUnit.TestFailRunTime GetSig(MethodName)

End Sub


Private Sub SomeReturnsFalseWhenSourceIsEmptyTest()

    On Error GoTo ErrHandler
    Const MethodName = "SomeReturnsFalseWhenSourceIsEmptyTest"

    ' Arrange
    Dim Predicate As New Predicate
    Predicate.Comparer = New LongComparer
    Predicate.ComparisonValue = 0

    ' Act & Assert
    ExUnit.IsFalse CollectionExt.Some(New Collection), GetSig(MethodName)
    ExUnit.IsFalse CollectionExt.Some(New Collection, Predicate), GetSig(MethodName)

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


Private Sub SkinReturnsEmptyCollectionWhenCountIsGreaterThanNumberOfItemsInSourceTest()

    On Error GoTo ErrHandler
    Const MethodName = "SkinReturnsEmptyCollectionWhenCountIsGreaterThanNumberOfItemsInSourceTest"
    
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


Private Sub FirstReturnsValueWhenSourceDoesSatisfyConditionTest()

    On Error GoTo ErrHandler
    Const MethodName = "FirstReturnsValueWhenSourceDoesSatisfyConditionTest"

    ' Arrange
    Dim Predicate As New Predicate
    With Predicate
        .ComparisonValue = 3
        .Comparer = New LongComparer
    End With

    ' Act & Assert
    ExUnit.AreEqual 3, CollectionExt.First(CollectionExt.Make(1, 2, 3), Predicate), GetSig(MethodName)

    Exit Sub
ErrHandler:
    Lapis.ExUnit.TestFailRunTime GetSig(MethodName)

End Sub


Private Sub FirstReturnsArgumentNullErrorWhenSourceIsNothingTest()

    On Error GoTo ErrHandler
    Const ExpectedError As Long = ErrorCode.ArgumentNull
    Const MethodName = "FirstReturnsArgumentNullErrorWhenSourceIsNothingTest"

    ' Arrange
    Dim Predicate As New Predicate
    With Predicate
        .ComparisonValue = 3
        .Comparer = New LongComparer
    End With

    ' Act
    CollectionExt.First Nothing, Predicate

    ' Assert
ErrHandler:
    Lapis.ExUnit.IsException ExpectedError, Err.Number, GetSig(MethodName)

End Sub


Private Sub FirstReturnsArgumentNullErrorWhenPredicateIsNothingTest()

    On Error GoTo ErrHandler
    Const ExpectedError As Long = ErrorCode.ArgumentNull
    Const MethodName = "FirstReturnsArgumentNullErrorWhenPredicateIsNothingTest"

    ' Act
    CollectionExt.First CollectionExt.Make(1, 2, 3), Nothing

    ' Assert
ErrHandler:
    Lapis.ExUnit.IsException ExpectedError, Err.Number, GetSig(MethodName)

End Sub


Private Sub FirstReturnsInvalidOperationErrorWhenSourceDoesNotSatisfyConditionTest()

    On Error GoTo ErrHandler
    Const ExpectedError As Long = ErrorCode.InvalidOperation
    Const MethodName = "FirstReturnsInvalidOperationErrorWhenSourceDoesNotSatisfyConditionTest"

    ' Arrange
    Dim Predicate As New Predicate
    With Predicate
        .ComparisonValue = 5
        .Comparer = New LongComparer
    End With

    ' Act
    CollectionExt.First CollectionExt.Make(1, 2, 3), Predicate

    ' Assert
ErrHandler:
    Lapis.ExUnit.IsException ExpectedError, Err.Number, GetSig(MethodName)

End Sub


Private Sub FirstReturnsInvalidOperationErrorWhenSourceIsEmptyTest()

    On Error GoTo ErrHandler
    Const ExpectedError As Long = ErrorCode.InvalidOperation
    Const MethodName = "FirstReturnsInvalidOperationErrorWhenSourceIsEmptyTest"

    ' Arrange
    Dim Predicate As New Predicate
    With Predicate
        .ComparisonValue = 3
        .Comparer = New LongComparer
    End With

    ' Act
    CollectionExt.First New Collection, Predicate

ErrHandler:
    Lapis.ExUnit.IsException ExpectedError, Err.Number, GetSig(MethodName)

End Sub


Private Sub LastReturnsValueWhenSourceDoesSatisfyConditionTest()

    On Error GoTo ErrHandler
    Const MethodName = "LastReturnsValueWhenSourceDoesSatisfyConditionTest"

    ' Arrange
    Dim Predicate As New Predicate
    With Predicate
        .ComparisonValue = 1
        .Operator = ComparisonOperator.GreaterThan
        .Comparer = New LongComparer
    End With

    ' Act & Assert
    ExUnit.AreEqual 3, CollectionExt.Last(CollectionExt.Make(1, 2, 3), Predicate), GetSig(MethodName)
    
    Exit Sub
ErrHandler:
    Lapis.ExUnit.TestFailRunTime GetSig(MethodName)

End Sub


Private Sub LastReturnsArgumentNullErrorWhenSourceIsNothingTest()

    On Error GoTo ErrHandler
    Const ExpectedError As Long = ErrorCode.ArgumentNull
    Const MethodName = "LastReturnsArgumentNullErrorWhenSourceIsNothingTest"

    ' Arrange
    Dim Predicate As New Predicate
    With Predicate
        .ComparisonValue = 1
        .Operator = ComparisonOperator.GreaterThan
        .Comparer = New LongComparer
    End With

    ' Act
    CollectionExt.Last Nothing, Predicate

    ' Assert
ErrHandler:
    Lapis.ExUnit.IsException ExpectedError, Err.Number, GetSig(MethodName)

End Sub


Private Sub LastReturnsArgumentNullErrorWhenPredicateIsNothingTest()

    On Error GoTo ErrHandler
    Const ExpectedError As Long = ErrorCode.ArgumentNull
    Const MethodName = "LastReturnsArgumentNullErrorWhenPredicateIsNothingTest"
    
    ' Act
    CollectionExt.Last CollectionExt.Make(1, 2, 3), Nothing

    ' Assert
ErrHandler:
    Lapis.ExUnit.IsException ExpectedError, Err.Number, GetSig(MethodName)

End Sub


Private Sub LastReturnsInvalidOperationErrorWhenSourceDoesNotSatisfyConditionTest()

    On Error GoTo ErrHandler
    Const ExpectedError As Long = ErrorCode.InvalidOperation
    Const MethodName = "LastReturnsInvalidOperationErrorWhenSourceDoesNotSatisfyConditionTest"

    ' Arrange
    Dim Predicate As New Predicate
    With Predicate
        .ComparisonValue = 5
        .Comparer = New LongComparer
    End With

    ' Act
    CollectionExt.Last CollectionExt.Make(1, 2, 3), Predicate

    ' Assert
ErrHandler:
    Lapis.ExUnit.IsException ExpectedError, Err.Number, GetSig(MethodName)

End Sub


Private Sub LastReturnsInvalidOperationErrorWhenSourceIsEmptyTest()

    On Error GoTo ErrHandler
    Const ExpectedError As Long = ErrorCode.InvalidOperation
    Const MethodName = "LastReturnsInvalidOperationErrorWhenSourceIsEmptyTest"

    ' Arrange
    Dim Predicate As New Predicate
    With Predicate
        .ComparisonValue = 5
        .Comparer = New LongComparer
    End With

    ' Act
    CollectionExt.Last New Collection, Predicate

    ' Assert

ErrHandler:
    Lapis.ExUnit.IsException ExpectedError, Err.Number, GetSig(MethodName)

End Sub


Private Sub SelectOneReturnsValueWhenSourceDoesSatisfyConditionTest()

    On Error GoTo ErrHandler
    Const MethodName = "SelectOneReturnsValueWhenSourceDoesSatisfyConditionTest"

    ' Arrange
    Dim Predicate As New Predicate
    With Predicate
        .ComparisonValue = 2
        .Comparer = New LongComparer
    End With

    ' Act & Assert
    ExUnit.AreEqual 2, CollectionExt.SelectOne(CollectionExt.Make(1, 2, 3), Predicate), GetSig(MethodName)

    Exit Sub
ErrHandler:
    Lapis.ExUnit.TestFailRunTime GetSig(MethodName)

End Sub


Private Sub SelectOneReturnsArgumentNullErrorWhenSourceIsNothingTest()

    On Error GoTo ErrHandler
    Const ExpectedError As Long = ErrorCode.ArgumentNull
    Const MethodName = "SelectOneReturnsArgumentNullErrorWhenSourceIsNothingTest"

    ' Arrange
    Dim Predicate As New Predicate
    With Predicate
        .ComparisonValue = 3
        .Comparer = New LongComparer
    End With

    ' Act
    CollectionExt.SelectOne Nothing, Predicate

    ' Assert
ErrHandler:
    Lapis.ExUnit.IsException ExpectedError, Err.Number, GetSig(MethodName)

End Sub


Private Sub SelectOneReturnsArgumentNullErrorWhenPredicateIsNothingTest()

    On Error GoTo ErrHandler
    Const ExpectedError As Long = ErrorCode.ArgumentNull
    Const MethodName = "SelectOneReturnsArgumentNullErrorWhenPredicateIsNothingTest"

    ' Act
    CollectionExt.SelectOne CollectionExt.Make(1, 2, 3), Nothing

    ' Assert
ErrHandler:
    Lapis.ExUnit.IsException ExpectedError, Err.Number, GetSig(MethodName)

End Sub


Private Sub SelectOneReturnsInvalidOperationErrorWhenSourceDoesNotSatisfyConditionTest()

    On Error GoTo ErrHandler
    Const ExpectedError As Long = ErrorCode.InvalidOperation
    Const MethodName = "SelectOneReturnsInvalidOperationErrorWhenSourceDoesNotSatisfyConditionTest"

    ' Arrange
    Dim Predicate As New Predicate
    With Predicate
        .ComparisonValue = 5
        .Comparer = New LongComparer
    End With

    ' Act
    CollectionExt.SelectOne CollectionExt.Make(1, 2, 3), Predicate

    ' Assert
ErrHandler:
    Lapis.ExUnit.IsException ExpectedError, Err.Number, GetSig(MethodName)

End Sub


Private Sub SelectOneReturnsInvalidOperationErrorWhenSourceIsEmptyTest()

    On Error GoTo ErrHandler
    Const ExpectedError As Long = ErrorCode.InvalidOperation
    Const MethodName = "SelectOneReturnsInvalidOperationErrorWhenSourceIsEmptyTest"

    ' Arrange
    Dim Predicate As New Predicate
    With Predicate
        .ComparisonValue = 3
        .Comparer = New LongComparer
    End With

    ' Act
    CollectionExt.SelectOne New Collection, Predicate

ErrHandler:
    Lapis.ExUnit.IsException ExpectedError, Err.Number, GetSig(MethodName)

End Sub


Private Sub SelectOneReturnsInvalidOperationErrorWhenSourceDoesSatisfyConditionMultipleTimesTest()

    On Error GoTo ErrHandler
    Const ExpectedError As Long = ErrorCode.InvalidOperation
    Const MethodName = "SelectOneReturnsInvalidOperationErrorWhenSourceDoesSatisfyConditionMultipleTimesTest"

    ' Arrange
    Dim Predicate As New Predicate
    With Predicate
        .ComparisonValue = 2
        .Operator = ComparisonOperator.GreaterThanOrEqualTo
        .Comparer = New LongComparer
    End With

    ' Act
    ExUnit.AreEqual 3, CollectionExt.SelectOne(CollectionExt.Make(1, 2, 3), Predicate), GetSig(MethodName)
    
    ' Assert
ErrHandler:
    Lapis.ExUnit.IsException ExpectedError, Err.Number, GetSig(MethodName)

End Sub


Private Sub CountTest()

    On Error GoTo ErrHandler
    Const MethodName = "CountTest"

    ' Arrange
    Dim Predicate As New Predicate
    With Predicate
        .ComparisonValue = 2
        .Operator = ComparisonOperator.GreaterThanOrEqualTo
        .Comparer = New LongComparer
    End With

    ' Act & Assert
    ExUnit.AreEqual 2, CollectionExt.Count(CollectionExt.Make(1, 2, 3), Predicate), GetSig(MethodName)

    Exit Sub
ErrHandler:
    Lapis.ExUnit.TestFailRunTime GetSig(MethodName)

End Sub


Private Sub CountReturnsArgumentNullErrorWhenSourceIsNothingTest()

    On Error GoTo ErrHandler
    Const ExpectedError As Long = ErrorCode.ArgumentNull
    Const MethodName = "CountReturnsArgumentNullErrorWhenSourceIsNothingTest"

    ' Arrange
    Dim Predicate As New Predicate
    With Predicate
        .ComparisonValue = 3
        .Comparer = New LongComparer
    End With

    ' Act
    CollectionExt.Count Nothing, Predicate

    ' Assert
ErrHandler:
    Lapis.ExUnit.IsException ExpectedError, Err.Number, GetSig(MethodName)

End Sub


Private Sub CountReturnsArgumentNullErrorWhenPredicateIsNothingTest()

    On Error GoTo ErrHandler
    Const ExpectedError As Long = ErrorCode.ArgumentNull
    Const MethodName = "CountReturnsArgumentNullErrorWhenPredicateIsNothingTest"

    ' Act
    CollectionExt.Count CollectionExt.Make(1, 2, 3), Nothing

    ' Assert
ErrHandler:
    Lapis.ExUnit.IsException ExpectedError, Err.Number, GetSig(MethodName)

End Sub


Private Sub WhereTest()

    On Error GoTo ErrHandler
    Const MethodName = "WhereTest"
    
    ' Arrange
    Dim Predicate As New Predicate
    With Predicate
        .ComparisonValue = 2
        .Comparer = New LongComparer
    End With

    ' Act
    Dim Actual As Collection
    Set Actual = CollectionExt.Where(CollectionExt.Make(1, 2, 3), Predicate)
    
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
    CollectionExt.Where Nothing, New Predicate

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
