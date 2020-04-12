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

    ' Act & Assert
    CollectionExt.ToString Nothing, New ValueTypeToStringConverter
    

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
    Dim Coll As Collection
    Set Coll = CollectionExt.Make("a", 1, 2.34, #2/2/2020#, &H5CC)
    
    ' Act
    Dim Actual As String
    Actual = CollectionExt.ToString(Coll, New ValueTypeToStringConverter)

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

    ' Arrange
    Dim Coll As New Collection
    Coll.Add ThisWorkbook
    Coll.Add ThisWorkbook

    ' Act
    Dim Actual As String
    Actual = CollectionExt.ToStringByProperty(Coll, "Name")

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
