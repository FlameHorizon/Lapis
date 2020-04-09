Attribute VB_Name = "CollectionTests"
Option Explicit

Private Const ModuleName As String = "CollectionTests"


Public Sub Start()

    IToStringValueTypesTest
    IToStringReturnArgumentNullErrorWhenCollectionIsNothingTest
    IToStringReturnsArgumentNullErrorWhenConverterIsNothingTest
    IToStringReturnsEmptyStringWhenCollectionIsEmptyTest
    IToStringReturnsStringWithDefinedDelimeterTest
    IToStringReturnsStringWhenDelimterIsNullStringTest
    IToStringReturnsStringWhenItemsAreNumbersTest
    IToStringReturnsStringWhenItemsAreMixedValueTypesTest
    IToStringReturnsStringWhenItemsContainsObjectsTest
    IToStringReturnsInvalidOperationWhenConverterCantConvertNonObjectTest
    IToStringReturnsInvalidOperationWhenConverterCantConvertDiffObjectTest
    IToStringReturnsInvalidOperationWhenConverterCantConvertNothingTest
    
    ToStringByPropertyTest
    ToStringByPropertyReturnsArgumentOutOfRangeErrorWhenPropertyDoesNotExistsTest
        
End Sub


Private Sub IToStringValueTypesTest()

    On Error GoTo ErrHandler
    Const MethodName = "IToStringValueTypesTest"

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


Private Sub IToStringReturnArgumentNullErrorWhenCollectionIsNothingTest()

    On Error GoTo ErrHandler
    Const ExpectedError As Long = ErrorCode.ArgumentNull
    Const MethodName = "IToStringReturnArgumentNullErrorWhenCollectionIsNothingTest"

    ' Act & Assert
    CollectionExt.ToString Nothing, New ValueTypeToStringConverter
    

ErrHandler:
    Lapis.ExUnit.IsException ExpectedError, Err.Number, GetSig(MethodName)

End Sub


Private Sub IToStringReturnsArgumentNullErrorWhenConverterIsNothingTest()

    On Error GoTo ErrHandler
    Const ExpectedError As Long = ErrorCode.ArgumentNull
    Const MethodName = "IToStringReturnsArgumentNullErrorWhenConverterIsNothingTest"

    ' Act
    CollectionExt.ToString New Collection, Nothing
    
    ' Assert
ErrHandler:
    Lapis.ExUnit.IsException ExpectedError, Err.Number, GetSig(MethodName)

End Sub


Private Sub IToStringReturnsEmptyStringWhenCollectionIsEmptyTest()

    On Error GoTo ErrHandler
    Const MethodName = "IToStringReturnsEmptyStringWhenCollectionIsEmptyTest"
    
    ' Act
    Dim Actual As String
    Actual = CollectionExt.ToString(New Collection, New ValueTypeToStringConverter)

    ' Assert
    ExUnit.AreEqual vbNullString, Actual, GetSig(MethodName)

    Exit Sub
ErrHandler:
    Lapis.ExUnit.TestFailRunTime GetSig(MethodName)

End Sub


Private Sub IToStringReturnsStringWithDefinedDelimeterTest()

    On Error GoTo ErrHandler
    Const MethodName = "IToStringReturnsStringWithDefinedDelimeterTest"

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


Private Sub IToStringReturnsStringWhenDelimterIsNullStringTest()

    On Error GoTo ErrHandler
    Const MethodName = "IToStringReturnsStringWhenDelimterIsNullStringTest"

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


Private Sub IToStringReturnsStringWhenItemsAreNumbersTest()

    On Error GoTo ErrHandler
    Const MethodName = "IToStringReturnsStringWhenItemsAreNumbersTest"

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


Private Sub IToStringReturnsStringWhenItemsAreMixedValueTypesTest()

    On Error GoTo ErrHandler
    Const MethodName = "IToStringReturnsStringWhenItemsAreMixedValueTypesTest"

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


Private Sub IToStringReturnsStringWhenItemsContainsObjectsTest()

    On Error GoTo ErrHandler
    Const MethodName = "IToStringReturnsStringWhenItemsContainsObjectsTest"

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


Private Sub IToStringReturnsInvalidOperationWhenConverterCantConvertNonObjectTest()

    On Error GoTo ErrHandler
    Const ExpectedError As Long = ErrorCode.InvalidOperation
    Const MethodName = "IToStringReturnsInvalidOperationWhenConverterCantConvertNonObjectTest"

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


Private Sub IToStringReturnsInvalidOperationWhenConverterCantConvertDiffObjectTest()

    On Error GoTo ErrHandler
    Const ExpectedError As Long = ErrorCode.InvalidOperation
    Const MethodName = "IToStringReturnsInvalidOperationWhenConverterCantConvertDiffObjectTest"

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


Private Sub IToStringReturnsInvalidOperationWhenConverterCantConvertNothingTest()

    On Error GoTo ErrHandler
    Const ExpectedError As Long = ErrorCode.InvalidOperation
    Const MethodName = "IToStringReturnsInvalidOperationWhenConverterCantConvertNothingTest"

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
