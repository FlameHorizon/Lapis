Attribute VB_Name = "StringHelperTests"
Option Explicit
'@Folder("Tests")

Private Const ModuleName As String = "StringHelperTests"


Public Sub Start()

    IndexOfAnyTest
    IndexOfAnyReturnsMinusWhenCharNotItStringTest
    IndexOfAnyReturnsZeroWhenAnyOfIsEmptyStringTest
    IndexOfAnyThrowsArgumentExceptionWhenAnyOfArrayIsNotInitalizedTest
    
    RemoveTest
    RemoveThrowsArgumentOutOfRangeWhenStartIndexIsNegetiveTest
    RemoveThrowsArgumentOutOfRangeWhenStartIndexIsGreaterThenLengthTest
    
    RemoveRangeTest
    RemoveRangeReturnsEmptyStringWhenLengthIsZeroTest
    RemoveRangeThrowsArgumentOutOfRangeWhenStartIndexIsNegetiveTest
    RemoveRangeThrowsArgumentOutOfRangeWhenCountIsNegetiveTest
    RemoveRangeThrowsArgumentOutOfRangeWhenStartIndexPlusCountIndexIsNotReachableTest
    
    ConcatTest
    
    StartsWithTest
    StartsWithReturnsTrueWhenValueStringIsNullStringTest
    StartsWithReturnsTrueWhenStrAndValueAreTheSameTest
    StartsWithThrowsArgumentOutOfRangeWhenStringComparisonIsDatabaseCompareTest
    StartsWithThrowsArgumentExceptionWhenStrIsNullStringTest
    
    EndsWithTest
    EndsWithReturnsTrueWhenValueStringIsNullStringTest
    EndsWithReturnsTrueWhenStrAndValueAreTheSameTest
    EndsWithThrowsArgumentOutOfRangeWhenStringComparisonIsDatabaseCompareTest
    EndsWithThrowsArgumentExceptionWhenStrIsNullStringTest
    
    InsertTest
    InsertThrowsArgumentExceptionWhenValueIsNullStringTest
    InsertThrowsArgumentOutOfRangeWhenStartIndexIsNegativeTest
    InsertThrowsArgumentOutOfRangeWhenStartIndexIsGreaterThanStrLenTest
    
    LastIndexOfTest
    LastIndexOfReturnsMinusOneWhenValueNotFound
    LastIndexOfThrowsArgumentExceptionWhenStrIsNullStringTest
    LastIndexOfThrowsArgumentOutOfRangeWhenStartIndexIsLesserThanZeroTest
    LastIndexOfThrowsArgumentOutOfRangeWhenStartIndexIsGreaterThanStrLen
    LastIndexOfThrowsArgumentOutOfRangeWhenWhenCountIsLesserThanZeroTest
    LastIndexOfThrowsArgumentOutOfRangeWhenStartIndexMinusCountIsLesserThanZeroTest
    LastIndexOfThrowsThrowsArgumentOutOfRangeWhenWhenComparisonMethodIsDatabaseTest
    
    LastIndexOfAnyTest
    LastIndexOfAnyReturnsMinusWhenCharNotItStringTest
    LastIndexOfAnyThrowsArgumentExceptionWhenAnyOfIsEmpty
    LastIndexOfAnyThrowsArgumentExceptionWhenAnyOfArrayIsNotInitalizedTest

End Sub


Private Sub IndexOfAnyTest()

    On Error GoTo ErrHandler
    Const MethodName = "IndexOfAnyTest"
    ExUnit.AreEqual 0, StringH.IndexOfAny("ab", StringArray("a")), GetFullSig(MethodName)
    ExUnit.AreEqual 1, StringH.IndexOfAny("ab", StringArray("b")), GetFullSig(MethodName)
    
    Exit Sub
ErrHandler:
    ExUnit.TestFailRunTime GetFullSig(MethodName)

End Sub


Private Function GetFullSig(ByVal MethodName As String) As String
    GetFullSig = ModuleName & "." & MethodName
End Function


Private Sub IndexOfAnyReturnsMinusWhenCharNotItStringTest()

    On Error GoTo ErrHandler
    Const MethodName = "IndexOfAnyReturnsZeroWhenCharNotItStringTest"
    ExUnit.AreEqual -1, StringH.IndexOfAny("ab", StringArray("z")), GetFullSig(MethodName)
    
    Exit Sub
ErrHandler:
    ExUnit.TestFailRunTime GetFullSig(MethodName)

End Sub


Private Sub IndexOfAnyThrowsArgumentExceptionWhenAnyOfArrayIsNotInitalizedTest()

    On Error GoTo ErrHandler
    Const MethodName = "IndexOfAnyThrowsArgumentExceptionWhenAnyOfArrayIsNotInitalizedTest"
    Dim NotInit() As String
    StringH.IndexOfAny "ab", NotInit
    
ErrHandler:
    ExUnit.IsException ExceptionCode.ArgumentException, Err.Number, GetFullSig(MethodName)
    On Error GoTo 0
    
End Sub


Private Sub IndexOfAnyReturnsZeroWhenAnyOfIsEmptyStringTest()

    Const MethodName = "IndexOfAnyReturnsZeroWhenAnyOfIsEmptyStringTest"
    Dim EmptyArr(0) As String
    ExUnit.AreEqual 0, StringH.IndexOfAny("ab", EmptyArr), GetFullSig(MethodName)
    
End Sub


Private Sub RemoveTest()

    On Error GoTo ErrHandler
    Const MethodName = "RemoveTest"
    ExUnit.AreEqual "foo", StringH.Remove("foo-bar", 3), GetFullSig(MethodName)
    
    Exit Sub
ErrHandler:
    ExUnit.TestFailRunTime GetFullSig(MethodName)
    On Error GoTo 0
        
End Sub


Private Sub RemoveThrowsArgumentOutOfRangeWhenStartIndexIsNegetiveTest()

    On Error GoTo ErrHandler
    Const MethodName = "RemoveThrowsArgumentOutOfRangeWhenStartIndexIsNegetiveTest"
    Dim NotInit() As String
    ExUnit.AreEqual "foo", StringH.Remove("foo-bar", -1), GetFullSig(MethodName)
    
    Exit Sub
ErrHandler:
    ExUnit.IsException ExceptionCode.ArgumentOutOfRange, Err.Number, GetFullSig(MethodName)
    On Error GoTo 0

End Sub


Private Sub RemoveThrowsArgumentOutOfRangeWhenStartIndexIsGreaterThenLengthTest()

    On Error GoTo ErrHandler
    Const MethodName = "RemoveThrowsArgumentOutOfRangeWhenStartIndexIsGreaterThenLengthTest"
    Dim NotInit() As String
    ExUnit.AreEqual "foo", StringH.Remove("foo-bar", 10), GetFullSig(MethodName)
    
    Exit Sub
ErrHandler:
    ExUnit.IsException ExceptionCode.ArgumentOutOfRange, Err.Number, GetFullSig(MethodName)
    On Error GoTo 0

End Sub


Private Sub RemoveRangeTest()

    On Error GoTo ErrHandler
    Const MethodName = "RemoveRangeTest"
    ExUnit.AreEqual "foobar", StringH.RemoveRange("foo-bar", 3, 1), GetFullSig(MethodName)
    ExUnit.AreEqual "bar", StringH.RemoveRange("foo-bar", 0, 4), GetFullSig(MethodName)
    
    Exit Sub
ErrHandler:
    ExUnit.TestFailRunTime GetFullSig(MethodName)
    On Error GoTo 0
    
End Sub


Private Sub RemoveRangeReturnsEmptyStringWhenLengthIsZeroTest()

    On Error GoTo ErrHandler
    Const MethodName = "RemoveRangeReturnsEmptyStringWhenLengthIsZeroTest"
    ExUnit.AreEqual vbNullString, StringH.RemoveRange("foo-bar", 0, Len("foo-bar")), GetFullSig(MethodName)

    Exit Sub
ErrHandler:
    ExUnit.TestFailRunTime GetFullSig(MethodName)
    On Error GoTo 0
    
End Sub


Private Sub RemoveRangeThrowsArgumentOutOfRangeWhenStartIndexIsNegetiveTest()

    On Error GoTo ErrHandler
    Const MethodName = "RemoveRangeThrowsArgumentOutOfRangeWhenStartIndexIsNegetiveTest"
    StringH.RemoveRange "a", -1, 0

ErrHandler:
    ExUnit.IsException ExceptionCode.ArgumentOutOfRange, Err.Number, GetFullSig(MethodName)
    On Error GoTo 0

End Sub


Private Sub RemoveRangeThrowsArgumentOutOfRangeWhenCountIsNegetiveTest()

    On Error GoTo ErrHandler
    Const MethodName = "RemoveRangeThrowsArgumentOutOfRangeWhenCountIsNegetiveTest"
    StringH.RemoveRange "a", 1, -1

ErrHandler:
    ExUnit.IsException ExceptionCode.ArgumentOutOfRange, Err.Number, GetFullSig(MethodName)
    On Error GoTo 0
    
End Sub


Private Sub RemoveRangeThrowsArgumentOutOfRangeWhenStartIndexPlusCountIndexIsNotReachableTest()

    On Error GoTo ErrHandler
    Const MethodName = "RemoveRangeThrowsArgumentOutOfRangeWhenStartIndexPlusCountIndexIsNotReachableTest"
    StringH.RemoveRange "foo-bar", 3, 10

ErrHandler:
    ExUnit.IsException ExceptionCode.ArgumentOutOfRange, Err.Number, GetFullSig(MethodName)
    On Error GoTo 0
    
End Sub


Private Sub ConcatTest()
    
    On Error GoTo ErrHandler
    Const MethodName = "ConcatTest"
    ExUnit.AreEqual "aaa bbb", StringH.Concat("aaa", " ", "bbb"), GetFullSig(MethodName)
    
    Exit Sub
ErrHandler:
    ExUnit.TestFailRunTime GetFullSig(MethodName)
    On Error GoTo 0
    
End Sub


Private Sub EndsWithTest()

    On Error GoTo ErrHandler
    Const MethodName = "EndsWithTest"
    ExUnit.IsTrue StringH.EndsWith("AvvS", "s", VbCompareMethod.vbTextCompare), GetFullSig(MethodName)
    ExUnit.IsFalse StringH.EndsWith("AvvS", "s", VbCompareMethod.vbBinaryCompare), GetFullSig(MethodName)
    
    Exit Sub
ErrHandler:
    ExUnit.TestFailRunTime GetFullSig(MethodName)
    On Error GoTo 0
    
End Sub


Private Sub EndsWithReturnsTrueWhenValueStringIsNullStringTest()

    On Error GoTo ErrHandler
    Const MethodName = "EndsWithReturnsTrueWhenValueStringIsNullStringTest"
    ExUnit.IsTrue StringH.EndsWith("aaa", vbNullString, vbBinaryCompare), GetFullSig(MethodName)
    
    Exit Sub
ErrHandler:
    ExUnit.TestFailRunTime GetFullSig(MethodName)
    On Error GoTo 0

End Sub


Private Sub EndsWithReturnsTrueWhenStrAndValueAreTheSameTest()

    On Error GoTo ErrHandler
    Const MethodName = "EndsWithReturnsTrueWhenStrAndValueAreTheSameTest"
    ExUnit.IsTrue StringH.EndsWith("aaa", "aaa", vbBinaryCompare), GetFullSig(MethodName)

    Exit Sub
ErrHandler:
    ExUnit.TestFailRunTime GetFullSig(MethodName)
    On Error GoTo 0

End Sub


Private Sub EndsWithThrowsArgumentOutOfRangeWhenStringComparisonIsDatabaseCompareTest()

    On Error GoTo ErrHandler
    Const MethodName = "EndsWithThrowsArgumentOutOfRangeWhenStringComparisonIsDatabaseCompareTest"
    StringH.EndsWith "aaa", "a", VbCompareMethod.vbDatabaseCompare

ErrHandler:
    ExUnit.IsException ExceptionCode.ArgumentOutOfRange, Err.Number, GetFullSig(MethodName)
    On Error GoTo 0

End Sub


Private Sub EndsWithThrowsArgumentExceptionWhenStrIsNullStringTest()

    On Error GoTo ErrHandler
    Const MethodName = "EndsWithThrowsArgumentExceptionWhenStrIsNullStringTest"
    StringH.EndsWith vbNullString, vbNullString, vbBinaryCompare

ErrHandler:
    ExUnit.IsException ExceptionCode.ArgumentException, Err.Number, GetFullSig(MethodName)
    On Error GoTo 0

End Sub


Private Sub StartsWithTest()
        
    On Error GoTo ErrHandler
    Const MethodName = "StartsWithTest"
    ExUnit.IsTrue StringH.StartsWith("Avvs", "a", VbCompareMethod.vbTextCompare), GetFullSig(MethodName)
    ExUnit.IsFalse StringH.StartsWith("Avvs", "a", VbCompareMethod.vbBinaryCompare), GetFullSig(MethodName)
    
    Exit Sub
ErrHandler:
    ExUnit.TestFailRunTime GetFullSig(MethodName)
    On Error GoTo 0
    
End Sub


Private Sub StartsWithReturnsTrueWhenValueStringIsNullStringTest()

    On Error GoTo ErrHandler
    Const MethodName = "StartsWithReturnsTrueWhenValueStringIsNullStringTest"
    ExUnit.IsTrue StringH.StartsWith("aaa", vbNullString, vbBinaryCompare), GetFullSig(MethodName)
    
    Exit Sub
ErrHandler:
    ExUnit.TestFailRunTime GetFullSig(MethodName)
    On Error GoTo 0

End Sub


Private Sub StartsWithReturnsTrueWhenStrAndValueAreTheSameTest()

    On Error GoTo ErrHandler
    Const MethodName = "StartsWithReturnsTrueWhenStrAndValueAreTheSameTest"
    ExUnit.IsTrue StringH.StartsWith("aaa", "aaa", vbBinaryCompare), GetFullSig(MethodName)

    Exit Sub
ErrHandler:
    ExUnit.TestFailRunTime GetFullSig(MethodName)
    On Error GoTo 0

End Sub


Private Sub StartsWithThrowsArgumentOutOfRangeWhenStringComparisonIsDatabaseCompareTest()

    On Error GoTo ErrHandler
    Const MethodName = "StartsWithThrowsArgumentOutOfRangeWhenStringComparisonIsDatabaseCompareTest"
    StringH.StartsWith "aaa", "a", VbCompareMethod.vbDatabaseCompare

ErrHandler:
    ExUnit.IsException ExceptionCode.ArgumentOutOfRange, Err.Number, GetFullSig(MethodName)
    On Error GoTo 0

End Sub


Private Sub StartsWithThrowsArgumentExceptionWhenStrIsNullStringTest()

    On Error GoTo ErrHandler
    Const MethodName = "StartsWithThrowsArgumentExceptionWhenStrIsNullStringTest"
    StringH.StartsWith vbNullString, vbNullString, vbBinaryCompare

ErrHandler:
    ExUnit.IsException ExceptionCode.ArgumentException, Err.Number, GetFullSig(MethodName)
    On Error GoTo 0

End Sub


Private Sub InsertTest()

    On Error GoTo ErrHandler
    Const MethodName = "InsertTest"
    ExUnit.AreEqual "aaa bbb", StringH.Insert("aaabbb", 3, " "), GetFullSig(MethodName)
    ExUnit.AreEqual "aaa", StringH.Insert("", 0, "aaa"), GetFullSig(MethodName)
    
    Exit Sub
ErrHandler:
    ExUnit.TestFailRunTime GetFullSig(MethodName)
    On Error GoTo 0
    
End Sub


Private Sub InsertThrowsArgumentExceptionWhenValueIsNullStringTest()
    
    On Error GoTo ErrHandler
    Const MethodName = "InsertThrowsArgumentExceptionWhenValueIsNullStringTest"
    StringH.Insert "a", 0, vbNullString

ErrHandler:
    ExUnit.IsException ExceptionCode.ArgumentException, Err.Number, GetFullSig(MethodName)
    On Error GoTo 0
    
End Sub


Private Sub InsertThrowsArgumentOutOfRangeWhenStartIndexIsNegativeTest()
    
    On Error GoTo ErrHandler
    Const MethodName = "InsertThrowsArgumentOutOfRangeWhenStartIndexIsNegativeTest"
    StringH.Insert "a", -1, "a"

ErrHandler:
    ExUnit.IsException ExceptionCode.ArgumentOutOfRange, Err.Number, GetFullSig(MethodName)
    On Error GoTo 0
    
End Sub


Private Sub InsertThrowsArgumentOutOfRangeWhenStartIndexIsGreaterThanStrLenTest()
    
    On Error GoTo ErrHandler
    Const MethodName = "InsertThrowsArgumentOutOfRangeWhenStartIndexIsGreaterThanStrLenTest"
    ExUnit.AreEqual "aaa", StringH.Insert("aaabbb", 10, "aaa"), GetFullSig(MethodName)

ErrHandler:
    ExUnit.IsException ExceptionCode.ArgumentOutOfRange, Err.Number, GetFullSig(MethodName)
    On Error GoTo 0
    
End Sub


Private Sub LastIndexOfTest()
    
    On Error GoTo ErrHandler
    Const MethodName = "LastIndexOfTest"
    ExUnit.AreEqual 5, StringH.LastIndexOf("bbbaaa", "a", Len("bbbaaa") - 1, Len("bbbaaa"), VbCompareMethod.vbTextCompare), GetFullSig(MethodName)
    ExUnit.AreEqual 2, StringH.LastIndexOf("bbbaaa", "b", Len("bbbaaa") - 1, Len("bbbaaa"), VbCompareMethod.vbTextCompare), GetFullSig(MethodName)
    
    Exit Sub
ErrHandler:
    ExUnit.TestFailRunTime GetFullSig(MethodName)
    On Error GoTo 0
    
End Sub


Private Sub LastIndexOfReturnsMinusOneWhenValueNotFound()
    
    On Error GoTo ErrHandler
    Const MethodName = "LastIndexOfReturnsMinusOneWhenValueNotFound"
    ExUnit.AreEqual -1, StringH.LastIndexOf("bbbaaa", "z", Len("bbbaaa") - 1, Len("bbbaaa"), VbCompareMethod.vbTextCompare), GetFullSig(MethodName)
    
    Exit Sub
ErrHandler:
    ExUnit.TestFailRunTime GetFullSig(MethodName)
    On Error GoTo 0
    
End Sub


Private Sub LastIndexOfThrowsArgumentExceptionWhenStrIsNullStringTest()
    
    On Error GoTo ErrHandler
    Const MethodName = "LastIndexOfThrowsArgumentExceptionWhenStrIsNullStringTest"
    StringH.LastIndexOf vbNullString, "a", Len("bbbaaa") - 1, Len("bbbaaa"), VbCompareMethod.vbTextCompare
    
ErrHandler:
    ExUnit.IsException ExceptionCode.ArgumentException, Err.Number, GetFullSig(MethodName)
    On Error GoTo 0

End Sub


Private Sub LastIndexOfThrowsArgumentOutOfRangeWhenStartIndexIsLesserThanZeroTest()
    
    On Error GoTo ErrHandler
    Const MethodName = "LastIndexOfThrowsArgumentOutOfRangeWhenStartIndexIsLesserThanZeroTest"
    StringH.LastIndexOf "bbbaaa", "a", -1, Len("bbbaaa"), VbCompareMethod.vbTextCompare

ErrHandler:
    ExUnit.IsException ExceptionCode.ArgumentOutOfRange, Err.Number, GetFullSig(MethodName)
    On Error GoTo 0

End Sub


Private Sub LastIndexOfThrowsArgumentOutOfRangeWhenStartIndexIsGreaterThanStrLen()
    
    On Error GoTo ErrHandler
    Const MethodName = "LastIndexOfThrowsArgumentOutOfRangeWhenStartIndexIsGreaterThanStrLen"
    StringH.LastIndexOf "bbbaaa", "a", Len("bbbaaa") + 1, Len("bbbaaa"), VbCompareMethod.vbTextCompare

ErrHandler:
    ExUnit.IsException ExceptionCode.ArgumentOutOfRange, Err.Number, GetFullSig(MethodName)
    On Error GoTo 0

End Sub


Private Sub LastIndexOfThrowsArgumentOutOfRangeWhenWhenCountIsLesserThanZeroTest()
    
    On Error GoTo ErrHandler
    Const MethodName = "LastIndexOfThrowsArgumentOutOfRangeWhenWhenCountIsLesserThanZeroTest"
    StringH.LastIndexOf "bbbaaa", "a", Len("bbbaaa") - 1, -1, VbCompareMethod.vbTextCompare

ErrHandler:
    ExUnit.IsException ExceptionCode.ArgumentOutOfRange, Err.Number, GetFullSig(MethodName)
    On Error GoTo 0

End Sub


Private Sub LastIndexOfThrowsArgumentOutOfRangeWhenStartIndexMinusCountIsLesserThanZeroTest()
    
    On Error GoTo ErrHandler
    Const MethodName = "LastIndexOfThrowsArgumentOutOfRangeWhenStartIndexMinusCountIsLesserThanZeroTest"
    StringH.LastIndexOf "bbbaaa", "a", Len("bbbaaa") - 1, Len("bbbaaa") + 1, VbCompareMethod.vbTextCompare

ErrHandler:
    ExUnit.IsException ExceptionCode.ArgumentOutOfRange, Err.Number, GetFullSig(MethodName)
    On Error GoTo 0

End Sub


Private Sub LastIndexOfThrowsThrowsArgumentOutOfRangeWhenWhenComparisonMethodIsDatabaseTest()
    
    On Error GoTo ErrHandler
    Const MethodName = "LastIndexOfThrowsThrowsArgumentOutOfRangeWhenWhenComparisonMethodIsDatabaseTest"
    StringH.LastIndexOf "bbbaaa", "a", Len("bbbaaa") - 1, Len("bbbaaa"), VbCompareMethod.vbDatabaseCompare

ErrHandler:
    ExUnit.IsException ExceptionCode.ArgumentOutOfRange, Err.Number, GetFullSig(MethodName)
    On Error GoTo 0

End Sub


Private Sub LastIndexOfAnyTest()

    On Error GoTo ErrHandler
    Const MethodName = "LastIndexOfAnyTest"
    ExUnit.AreEqual 0, StringH.LastIndexOfAny("abcd", StringArray("a", "d")), GetFullSig(MethodName)
    ExUnit.AreEqual 3, StringH.LastIndexOfAny("abcd", StringArray("d", "a")), GetFullSig(MethodName)
    ExUnit.AreEqual 1, StringH.LastIndexOfAny("abcd", StringArray("b", "c")), GetFullSig(MethodName)
    
    Exit Sub
ErrHandler:
    ExUnit.TestFailRunTime GetFullSig(MethodName)
    
End Sub


Private Sub LastIndexOfAnyReturnsMinusWhenCharNotItStringTest()

    On Error GoTo ErrHandler
    Const MethodName = "LastIndexOfAnyReturnsZeroWhenCharNotItStringTest"
    ExUnit.AreEqual -1, StringH.LastIndexOfAny("ab", StringArray("z")), GetFullSig(MethodName)
    
    Exit Sub
ErrHandler:
    ExUnit.TestFailRunTime GetFullSig(MethodName)

End Sub


Private Sub LastIndexOfAnyThrowsArgumentExceptionWhenAnyOfArrayIsNotInitalizedTest()

    On Error GoTo ErrHandler
    Const MethodName = "LastIndexOfAnyThrowsArgumentExceptionWhenAnyOfArrayIsNotInitalizedTest"
    Dim NotInit() As String
    StringH.LastIndexOfAny "ab", NotInit
    
ErrHandler:
    ExUnit.IsException ExceptionCode.ArgumentException, Err.Number, GetFullSig(MethodName)
    On Error GoTo 0
    
End Sub


Private Sub LastIndexOfAnyThrowsArgumentExceptionWhenAnyOfIsEmpty()
    
    On Error GoTo ErrHandler
    Const MethodName = "LastIndexOfAnyThrowsArgumentExceptionWhenAnyOfIsEmpty"
    Dim EmptyArr(0) As String
    ExUnit.AreEqual 0, StringH.LastIndexOfAny("ab", EmptyArr), GetFullSig(MethodName)
    
ErrHandler:
    ExUnit.IsException ExceptionCode.ArgumentException, Err.Number, GetFullSig(MethodName)
    On Error GoTo 0
    
End Sub

