Attribute VB_Name = "ComparersTests"
'@Folder("Tests")
Option Explicit

Private Const ModuleName As String = "ComparersTest"


Public Sub Start()
    
    LongComparerReturnsOneWhenNumberIsComparedToNothingTest
    
    ' Default comparers
    DefaultReturnsLongLongComparerWhenNumberWithoutDecimalTest
    DefaultReturnsStringComparerWhenValueIsStringTest
    DefaultReturnsDecimalComaprerWhenValueWithDecimalTest
    DefaultReturnsNothingWhenTypeIsNotHandledTest

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


Private Sub DefaultReturnsLongLongComparerWhenNumberWithoutDecimalTest()

    On Error GoTo ErrHandler
    Const MethodName = "DefaultReturnsLongLongComparerWhenNumberWithoutDecimalTest"
 
    ' Act & Assert
    Lapis.ExUnit.IsTrue TypeOf Comparers.Default(CBool(1)) Is LongLongComparer, GetSig(MethodName)
    Lapis.ExUnit.IsTrue TypeOf Comparers.Default(CByte(1)) Is LongLongComparer, GetSig(MethodName)
    Lapis.ExUnit.IsTrue TypeOf Comparers.Default(CInt(1)) Is LongLongComparer, GetSig(MethodName)
    Lapis.ExUnit.IsTrue TypeOf Comparers.Default(CLng(1)) Is LongLongComparer, GetSig(MethodName)
    Lapis.ExUnit.IsTrue TypeOf Comparers.Default(CLngLng(1)) Is LongLongComparer, GetSig(MethodName)

    Exit Sub
ErrHandler:
    Lapis.ExUnit.TestFailRunTime GetSig(MethodName)

End Sub


Private Sub DefaultReturnsStringComparerWhenValueIsStringTest()

    On Error GoTo ErrHandler
    Const MethodName = "DefaultReturnsStringComparerWhenValueIsStringTest"
    
    ' Act & Assert
    Lapis.ExUnit.IsTrue TypeOf Comparers.Default("1") Is OrdinalIgnoreCaseStringComparer, GetSig(MethodName)

    Exit Sub
ErrHandler:
    Lapis.ExUnit.TestFailRunTime GetSig(MethodName)

End Sub


Private Sub DefaultReturnsDecimalComaprerWhenValueWithDecimalTest()

    On Error GoTo ErrHandler
    Const MethodName = "DefaultReturnsDecimalComaprerWhenValueWithDecimalTest"
    
    ' Act & Assert
    Lapis.ExUnit.IsTrue TypeOf Comparers.Default(CSng(1)) Is DecimalComparer, GetSig(MethodName)
    Lapis.ExUnit.IsTrue TypeOf Comparers.Default(CDbl(1)) Is DecimalComparer, GetSig(MethodName)
    Lapis.ExUnit.IsTrue TypeOf Comparers.Default(CCur(1)) Is DecimalComparer, GetSig(MethodName)
    Lapis.ExUnit.IsTrue TypeOf Comparers.Default(CDec(1)) Is DecimalComparer, GetSig(MethodName)
    Lapis.ExUnit.IsTrue TypeOf Comparers.Default(CDate(1)) Is DecimalComparer, GetSig(MethodName)

    Exit Sub
ErrHandler:
    Lapis.ExUnit.TestFailRunTime GetSig(MethodName)

End Sub


Private Sub DefaultReturnsNothingWhenTypeIsNotHandledTest()

    On Error GoTo ErrHandler
    Const MethodName = "DefaultReturnsNothingWhenTypeIsNotHandledTest"
    
    ' Act & Assert
    Lapis.ExUnit.IsNothing Comparers.Default(Array()), GetSig(MethodName)

    Exit Sub
ErrHandler:
    Lapis.ExUnit.TestFailRunTime GetSig(MethodName)

End Sub
