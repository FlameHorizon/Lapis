Attribute VB_Name = "CollectionTests"
'@Folder("Tests")
Option Explicit

Private Const ModuleName As String = "CollectionTests"


Public Sub Start()

    ToStringValueTypesTest
    ToStringReturnArgumentNullErrorWhenCollectionIsNothingTest
    ToStringReturnsArgumentNullErrorWhenConverterIsNothingTest
    ToStringReturnsEmptyStringWhenCollectionIsEmptyTest
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

    ' Arrange
    Dim Coll As New Collection
    Coll.Add "abc"
    Coll.Add "cba"

    ' Act
    Dim Actual As String
    Actual = CollectionExt.ToString(Coll, New ValueTypeToStringConverter)

    ' Assert
    Lapis.ExUnit.AreEqual "abc,cba", Actual, GetSig(MethodName)

    Exit Sub
ErrHandler:
    Lapis.ExUnit.TestFailRunTime GetSig(MethodName)

End Sub


Private Function GetSig(ByVal MethodName As String) As String
    GetSig = ModuleName & "." & MethodName
End Function


Private Sub ToStringReturnArgumentNullErrorWhenCollectionIsNothingTest()

    On Error GoTo ErrHandler
    Const ExpectedError As Long = ErrorCode.ArgumentNull
    Const MethodName = "ToStringReturnArgumentNullErrorWhenCollectionIsNothingTest"

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


Private Sub ToStringReturnsEmptyStringWhenCollectionIsEmptyTest()

    On Error GoTo ErrHandler
    Const MethodName = "ToStringReturnsEmptyStringWhenCollectionIsEmptyTest"
    
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

    ' Arrange
    Dim Coll As New Collection
    Coll.Add "abc"
    Coll.Add "cba"

    ' Act
    Dim Actual As String
    Actual = CollectionExt.ToString(Coll, New ValueTypeToStringConverter, "-")
    
    ' Assert
    ExUnit.AreEqual "abc-cba", Actual, GetSig(MethodName)

    Exit Sub
ErrHandler:
    Lapis.ExUnit.TestFailRunTime GetSig(MethodName)

End Sub


Private Sub ToStringReturnsStringWhenDelimterIsNullStringTest()

    On Error GoTo ErrHandler
    Const MethodName = "ToStringReturnsStringWhenDelimterIsNullStringTest"

    ' Arrange
    Dim Coll As New Collection
    Coll.Add "abc"
    Coll.Add "cba"

    ' Act
    Dim Actual As String
    Actual = CollectionExt.ToString(Coll, New ValueTypeToStringConverter, vbNullString)
    
    ' Assert
    ExUnit.AreEqual "abccba", Actual, GetSig(MethodName)


    Exit Sub
ErrHandler:
    Lapis.ExUnit.TestFailRunTime GetSig(MethodName)

End Sub


Private Sub ToStringReturnsStringWhenItemsAreNumbersTest()

    On Error GoTo ErrHandler
    Const MethodName = "ToStringReturnsStringWhenItemsAreNumbersTest"

    ' Arrange
    Dim Coll As New Collection
    Coll.Add 1
    Coll.Add 3
    Coll.Add 4
    
    ' Act
    Dim Actual As String
    Actual = CollectionExt.ToString(Coll, New ValueTypeToStringConverter)

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
    Dim Coll As New Collection
    Coll.Add "a"
    Coll.Add 1
    Coll.Add 2.34
    Coll.Add #2/2/2020#
    Coll.Add &H5CC
    
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

    ' Arrange
    Dim Coll As New Collection
    Coll.Add ThisWorkbook
    Coll.Add ThisWorkbook

    ' Act
    Dim Actual As String
    Actual = CollectionExt.ToString(Coll, New WorkbookToStringConverter)

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

    ' Arrange
    Dim Coll As New Collection
    Coll.Add "a"

    ' Act
    Dim Actual As String
    Actual = CollectionExt.ToString(Coll, New WorkbookToStringConverter)

    ' Assert
ErrHandler:
    Lapis.ExUnit.IsException ExpectedError, Err.Number, GetSig(MethodName)

End Sub


Private Sub ToStringReturnsInvalidOperationWhenConverterCantConvertDiffObjectTest()

    On Error GoTo ErrHandler
    Const ExpectedError As Long = ErrorCode.InvalidOperation
    Const MethodName = "ToStringReturnsInvalidOperationWhenConverterCantConvertDiffObjectTest"

    ' Arrange
    Dim Coll As New Collection
    Coll.Add New Collection

    ' Act
    Dim Actual As String
    Actual = CollectionExt.ToString(Coll, New WorkbookToStringConverter)

    ' Assert
ErrHandler:
    Lapis.ExUnit.IsException ExpectedError, Err.Number, GetSig(MethodName)

End Sub


Private Sub ToStringReturnsInvalidOperationWhenConverterCantConvertNothingTest()

    On Error GoTo ErrHandler
    Const ExpectedError As Long = ErrorCode.InvalidOperation
    Const MethodName = "ToStringReturnsInvalidOperationWhenConverterCantConvertNothingTest"

    ' Arrange
    Dim Coll As New Collection
    Coll.Add Nothing

    ' Act
    Dim Actual As String
    Actual = CollectionExt.ToString(Coll, New WorkbookToStringConverter)

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

    ' Arrange
    Dim Coll As New Collection
    Coll.Add ThisWorkbook
    Coll.Add ThisWorkbook

    ' Act
    Dim Actual As String
    Actual = CollectionExt.ToStringByProperty(Coll, "NotExistsingProperty")

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
