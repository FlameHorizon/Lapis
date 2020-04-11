Attribute VB_Name = "CollectionTests"
'@Folder("Tests")
Option Explicit

Private Const ModuleName As String = "CollectionTests"


Public Sub Start()

    ToStringValueTypesTest
    ToStringReturnArgumentNullErrorWhenSourceIsNothingTest
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


Private Sub ToStringReturnArgumentNullErrorWhenSourceIsNothingTest()

    On Error GoTo ErrHandler
    Const ExpectedError As Long = ErrorCode.ArgumentNull
    Const MethodName = "ToStringReturnArgumentNullErrorWhenSourceIsNothingTest"

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
    Actual = CollectionExt.ToStringByProperty(CollectionExt.Make(ThisWorkbook, ThisWorkbook), "NotExistsingProperty")

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


