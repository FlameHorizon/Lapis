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
<<<<<<< HEAD
    ExUnit.AreEqual 0, StringH.IndexOfAny("ab", StringArray("a")), GetSig(MethodName)
    ExUnit.AreEqual 1, StringH.IndexOfAny("ab", StringArray("b")), GetSig(MethodName)
=======
    ExUnit.AreEqual 0, StringExt.IndexOfAny("ab", StringArray("a")), GetFullSig(MethodName)
    ExUnit.AreEqual 1, StringExt.IndexOfAny("ab", StringArray("b")), GetFullSig(MethodName)
>>>>>>> 524a3180e197414b591569bde713af7aadc226c6
    
    Exit Sub
ErrHandler:
    ExUnit.TestFailRunTime GetSig(MethodName)

End Sub


Private Function GetSig(ByVal MethodName As String) As String
    GetSig = ModuleName & "." & MethodName
End Function


Private Sub IndexOfAnyReturnsMinusWhenCharNotItStringTest()

    On Error GoTo ErrHandler
    Const MethodName = "IndexOfAnyReturnsZeroWhenCharNotItStringTest"
<<<<<<< HEAD
    ExUnit.AreEqual -1, StringH.IndexOfAny("ab", StringArray("z")), GetSig(MethodName)
=======
    ExUnit.AreEqual -1, StringExt.IndexOfAny("ab", StringArray("z")), GetFullSig(MethodName)
>>>>>>> 524a3180e197414b591569bde713af7aadc226c6
    
    Exit Sub
ErrHandler:
    ExUnit.TestFailRunTime GetSig(MethodName)

End Sub


Private Sub IndexOfAnyThrowsArgumentExceptionWhenAnyOfArrayIsNotInitalizedTest()

    On Error GoTo ErrHandler
    Const MethodName = "IndexOfAnyThrowsArgumentExceptionWhenAnyOfArrayIsNotInitalizedTest"
    Dim NotInit() As String
    StringExt.IndexOfAny "ab", NotInit
    
ErrHandler:
    ExUnit.IsException ExceptionCode.ArgumentException, Err.Number, GetSig(MethodName)
    
End Sub


Private Sub IndexOfAnyReturnsZeroWhenAnyOfIsEmptyStringTest()
    
    On Error GoTo ErrHandler
    Const MethodName = "IndexOfAnyReturnsZeroWhenAnyOfIsEmptyStringTest"
    Dim EmptyArr(0) As String
<<<<<<< HEAD
    ExUnit.AreEqual 0, StringH.IndexOfAny("ab", EmptyArr), GetSig(MethodName)
=======
    ExUnit.AreEqual 0, StringExt.IndexOfAny("ab", EmptyArr), GetFullSig(MethodName)
>>>>>>> 524a3180e197414b591569bde713af7aadc226c6
    
    Exit Sub
ErrHandler:
    ExUnit.TestFailRunTime GetSig(MethodName)
    
End Sub


Private Sub RemoveTest()

    On Error GoTo ErrHandler
    Const MethodName = "RemoveTest"
<<<<<<< HEAD
    ExUnit.AreEqual "foo", StringH.Remove("foo-bar", 3), GetSig(MethodName)
=======
    ExUnit.AreEqual "foo", StringExt.Remove("foo-bar", 3), GetFullSig(MethodName)
>>>>>>> 524a3180e197414b591569bde713af7aadc226c6
    
    Exit Sub
ErrHandler:
    ExUnit.TestFailRunTime GetSig(MethodName)
        
End Sub


Private Sub RemoveThrowsArgumentOutOfRangeWhenStartIndexIsNegetiveTest()

    On Error GoTo ErrHandler
    Const MethodName = "RemoveThrowsArgumentOutOfRangeWhenStartIndexIsNegetiveTest"
<<<<<<< HEAD
    ExUnit.AreEqual "foo", StringH.Remove("foo-bar", -1), GetSig(MethodName)
=======
    ExUnit.AreEqual "foo", StringExt.Remove("foo-bar", -1), GetFullSig(MethodName)
>>>>>>> 524a3180e197414b591569bde713af7aadc226c6
    
ErrHandler:
    ExUnit.IsException ExceptionCode.ArgumentOutOfRange, Err.Number, GetSig(MethodName)

End Sub


Private Sub RemoveThrowsArgumentOutOfRangeWhenStartIndexIsGreaterThenLengthTest()

    On Error GoTo ErrHandler
    Const MethodName = "RemoveThrowsArgumentOutOfRangeWhenStartIndexIsGreaterThenLengthTest"
<<<<<<< HEAD
    ExUnit.AreEqual "foo", StringH.Remove("foo-bar", 10), GetSig(MethodName)
=======
    ExUnit.AreEqual "foo", StringExt.Remove("foo-bar", 10), GetFullSig(MethodName)
>>>>>>> 524a3180e197414b591569bde713af7aadc226c6
    
ErrHandler:
    ExUnit.IsException ExceptionCode.ArgumentOutOfRange, Err.Number, GetSig(MethodName)

End Sub


Private Sub RemoveRangeTest()

    On Error GoTo ErrHandler
    Const MethodName = "RemoveRangeTest"
<<<<<<< HEAD
    ExUnit.AreEqual "foobar", StringH.RemoveRange("foo-bar", 3, 1), GetSig(MethodName)
    ExUnit.AreEqual "bar", StringH.RemoveRange("foo-bar", 0, 4), GetSig(MethodName)
=======
    ExUnit.AreEqual "foobar", StringExt.RemoveRange("foo-bar", 3, 1), GetFullSig(MethodName)
    ExUnit.AreEqual "bar", StringExt.RemoveRange("foo-bar", 0, 4), GetFullSig(MethodName)
>>>>>>> 524a3180e197414b591569bde713af7aadc226c6
    
    Exit Sub
ErrHandler:
    ExUnit.TestFailRunTime GetSig(MethodName)
    
End Sub


Private Sub RemoveRangeReturnsEmptyStringWhenLengthIsZeroTest()

    On Error GoTo ErrHandler
    Const MethodName = "RemoveRangeReturnsEmptyStringWhenLengthIsZeroTest"
<<<<<<< HEAD
    ExUnit.AreEqual vbNullString, StringH.RemoveRange("foo-bar", 0, Len("foo-bar")), GetSig(MethodName)
=======
    ExUnit.AreEqual vbNullString, StringExt.RemoveRange("foo-bar", 0, Len("foo-bar")), GetFullSig(MethodName)
>>>>>>> 524a3180e197414b591569bde713af7aadc226c6

    Exit Sub
ErrHandler:
    ExUnit.TestFailRunTime GetSig(MethodName)
    
End Sub


Private Sub RemoveRangeThrowsArgumentOutOfRangeWhenStartIndexIsNegetiveTest()

    On Error GoTo ErrHandler
    Const MethodName = "RemoveRangeThrowsArgumentOutOfRangeWhenStartIndexIsNegetiveTest"
    StringExt.RemoveRange "a", -1, 0

ErrHandler:
    ExUnit.IsException ExceptionCode.ArgumentOutOfRange, Err.Number, GetSig(MethodName)

End Sub


Private Sub RemoveRangeThrowsArgumentOutOfRangeWhenCountIsNegetiveTest()

    On Error GoTo ErrHandler
    Const MethodName = "RemoveRangeThrowsArgumentOutOfRangeWhenCountIsNegetiveTest"
    StringExt.RemoveRange "a", 1, -1

ErrHandler:
    ExUnit.IsException ExceptionCode.ArgumentOutOfRange, Err.Number, GetSig(MethodName)
    
End Sub


Private Sub RemoveRangeThrowsArgumentOutOfRangeWhenStartIndexPlusCountIndexIsNotReachableTest()

    On Error GoTo ErrHandler
    Const MethodName = "RemoveRangeThrowsArgumentOutOfRangeWhenStartIndexPlusCountIndexIsNotReachableTest"
    StringExt.RemoveRange "foo-bar", 3, 10

ErrHandler:
    ExUnit.IsException ExceptionCode.ArgumentOutOfRange, Err.Number, GetSig(MethodName)
    
End Sub


Private Sub ConcatTest()
    
    On Error GoTo ErrHandler
    Const MethodName = "ConcatTest"
<<<<<<< HEAD
    ExUnit.AreEqual "aaa bbb", StringH.Concat("aaa", " ", "bbb"), GetSig(MethodName)
=======
    ExUnit.AreEqual "aaa bbb", StringExt.Concat("aaa", " ", "bbb"), GetFullSig(MethodName)
>>>>>>> 524a3180e197414b591569bde713af7aadc226c6
    
    Exit Sub
ErrHandler:
    ExUnit.TestFailRunTime GetSig(MethodName)
    
End Sub


Private Sub EndsWithTest()

    On Error GoTo ErrHandler
    Const MethodName = "EndsWithTest"
<<<<<<< HEAD
    ExUnit.IsTrue StringH.EndsWith("AvvS", "s", VbCompareMethod.vbTextCompare), GetSig(MethodName)
    ExUnit.IsFalse StringH.EndsWith("AvvS", "s", VbCompareMethod.vbBinaryCompare), GetSig(MethodName)
=======
    ExUnit.IsTrue StringExt.EndsWith("AvvS", "s", VbCompareMethod.vbTextCompare), GetFullSig(MethodName)
    ExUnit.IsFalse StringExt.EndsWith("AvvS", "s", VbCompareMethod.vbBinaryCompare), GetFullSig(MethodName)
>>>>>>> 524a3180e197414b591569bde713af7aadc226c6
    
    Exit Sub
ErrHandler:
    ExUnit.TestFailRunTime GetSig(MethodName)
    
End Sub


Private Sub EndsWithReturnsTrueWhenValueStringIsNullStringTest()

    On Error GoTo ErrHandler
    Const MethodName = "EndsWithReturnsTrueWhenValueStringIsNullStringTest"
<<<<<<< HEAD
    ExUnit.IsTrue StringH.EndsWith("aaa", vbNullString, vbBinaryCompare), GetSig(MethodName)
=======
    ExUnit.IsTrue StringExt.EndsWith("aaa", vbNullString, vbBinaryCompare), GetFullSig(MethodName)
>>>>>>> 524a3180e197414b591569bde713af7aadc226c6
    
    Exit Sub
ErrHandler:
    ExUnit.TestFailRunTime GetSig(MethodName)

End Sub


Private Sub EndsWithReturnsTrueWhenStrAndValueAreTheSameTest()

    On Error GoTo ErrHandler
    Const MethodName = "EndsWithReturnsTrueWhenStrAndValueAreTheSameTest"
<<<<<<< HEAD
    ExUnit.IsTrue StringH.EndsWith("aaa", "aaa", vbBinaryCompare), GetSig(MethodName)
=======
    ExUnit.IsTrue StringExt.EndsWith("aaa", "aaa", vbBinaryCompare), GetFullSig(MethodName)
>>>>>>> 524a3180e197414b591569bde713af7aadc226c6

    Exit Sub
ErrHandler:
    ExUnit.TestFailRunTime GetSig(MethodName)

End Sub


Private Sub EndsWithThrowsArgumentOutOfRangeWhenStringComparisonIsDatabaseCompareTest()

    On Error GoTo ErrHandler
    Const MethodName = "EndsWithThrowsArgumentOutOfRangeWhenStringComparisonIsDatabaseCompareTest"
    StringExt.EndsWith "aaa", "a", VbCompareMethod.vbDatabaseCompare

ErrHandler:
    ExUnit.IsException ExceptionCode.ArgumentOutOfRange, Err.Number, GetSig(MethodName)

End Sub


Private Sub EndsWithThrowsArgumentExceptionWhenStrIsNullStringTest()

    On Error GoTo ErrHandler
    Const MethodName = "EndsWithThrowsArgumentExceptionWhenStrIsNullStringTest"
    StringExt.EndsWith vbNullString, vbNullString, vbBinaryCompare

ErrHandler:
    ExUnit.IsException ExceptionCode.ArgumentException, Err.Number, GetSig(MethodName)

End Sub


Private Sub StartsWithTest()
        
    On Error GoTo ErrHandler
    Const MethodName = "StartsWithTest"
<<<<<<< HEAD
    ExUnit.IsTrue StringH.StartsWith("Avvs", "a", VbCompareMethod.vbTextCompare), GetSig(MethodName)
    ExUnit.IsFalse StringH.StartsWith("Avvs", "a", VbCompareMethod.vbBinaryCompare), GetSig(MethodName)
=======
    ExUnit.IsTrue StringExt.StartsWith("Avvs", "a", VbCompareMethod.vbTextCompare), GetFullSig(MethodName)
    ExUnit.IsFalse StringExt.StartsWith("Avvs", "a", VbCompareMethod.vbBinaryCompare), GetFullSig(MethodName)
>>>>>>> 524a3180e197414b591569bde713af7aadc226c6
    
    Exit Sub
ErrHandler:
    ExUnit.TestFailRunTime GetSig(MethodName)
    
End Sub


Private Sub StartsWithReturnsTrueWhenValueStringIsNullStringTest()

    On Error GoTo ErrHandler
    Const MethodName = "StartsWithReturnsTrueWhenValueStringIsNullStringTest"
<<<<<<< HEAD
    ExUnit.IsTrue StringH.StartsWith("aaa", vbNullString, vbBinaryCompare), GetSig(MethodName)
=======
    ExUnit.IsTrue StringExt.StartsWith("aaa", vbNullString, vbBinaryCompare), GetFullSig(MethodName)
>>>>>>> 524a3180e197414b591569bde713af7aadc226c6
    
    Exit Sub
ErrHandler:
    ExUnit.TestFailRunTime GetSig(MethodName)

End Sub


Private Sub StartsWithReturnsTrueWhenStrAndValueAreTheSameTest()

    On Error GoTo ErrHandler
    Const MethodName = "StartsWithReturnsTrueWhenStrAndValueAreTheSameTest"
<<<<<<< HEAD
    ExUnit.IsTrue StringH.StartsWith("aaa", "aaa", vbBinaryCompare), GetSig(MethodName)
=======
    ExUnit.IsTrue StringExt.StartsWith("aaa", "aaa", vbBinaryCompare), GetFullSig(MethodName)
>>>>>>> 524a3180e197414b591569bde713af7aadc226c6

    Exit Sub
ErrHandler:
    ExUnit.TestFailRunTime GetSig(MethodName)

End Sub


Private Sub StartsWithThrowsArgumentOutOfRangeWhenStringComparisonIsDatabaseCompareTest()

    On Error GoTo ErrHandler
    Const MethodName = "StartsWithThrowsArgumentOutOfRangeWhenStringComparisonIsDatabaseCompareTest"
    StringExt.StartsWith "aaa", "a", VbCompareMethod.vbDatabaseCompare

ErrHandler:
    ExUnit.IsException ExceptionCode.ArgumentOutOfRange, Err.Number, GetSig(MethodName)

End Sub


Private Sub StartsWithThrowsArgumentExceptionWhenStrIsNullStringTest()

    On Error GoTo ErrHandler
    Const MethodName = "StartsWithThrowsArgumentExceptionWhenStrIsNullStringTest"
    StringExt.StartsWith vbNullString, vbNullString, vbBinaryCompare

ErrHandler:
    ExUnit.IsException ExceptionCode.ArgumentException, Err.Number, GetSig(MethodName)

End Sub


Private Sub InsertTest()

    On Error GoTo ErrHandler
    Const MethodName = "InsertTest"
<<<<<<< HEAD
    ExUnit.AreEqual "aaa bbb", StringH.Insert("aaabbb", 3, " "), GetSig(MethodName)
    ExUnit.AreEqual "aaa", StringH.Insert("", 0, "aaa"), GetSig(MethodName)
=======
    ExUnit.AreEqual "aaa bbb", StringExt.Insert("aaabbb", 3, " "), GetFullSig(MethodName)
    ExUnit.AreEqual "aaa", StringExt.Insert(vbNullString, 0, "aaa"), GetFullSig(MethodName)
>>>>>>> 524a3180e197414b591569bde713af7aadc226c6
    
    Exit Sub
ErrHandler:
    ExUnit.TestFailRunTime GetSig(MethodName)
    
End Sub


Private Sub InsertThrowsArgumentExceptionWhenValueIsNullStringTest()
    
    On Error GoTo ErrHandler
    Const MethodName = "InsertThrowsArgumentExceptionWhenValueIsNullStringTest"
    StringExt.Insert "a", 0, vbNullString

ErrHandler:
    ExUnit.IsException ExceptionCode.ArgumentException, Err.Number, GetSig(MethodName)
    
End Sub


Private Sub InsertThrowsArgumentOutOfRangeWhenStartIndexIsNegativeTest()
    
    On Error GoTo ErrHandler
    Const MethodName = "InsertThrowsArgumentOutOfRangeWhenStartIndexIsNegativeTest"
    StringExt.Insert "a", -1, "a"

ErrHandler:
    ExUnit.IsException ExceptionCode.ArgumentOutOfRange, Err.Number, GetSig(MethodName)
    
End Sub


Private Sub InsertThrowsArgumentOutOfRangeWhenStartIndexIsGreaterThanStrLenTest()
    
    On Error GoTo ErrHandler
    Const MethodName = "InsertThrowsArgumentOutOfRangeWhenStartIndexIsGreaterThanStrLenTest"
<<<<<<< HEAD
    ExUnit.AreEqual "aaa", StringH.Insert("aaabbb", 10, "aaa"), GetSig(MethodName)
=======
    ExUnit.AreEqual "aaa", StringExt.Insert("aaabbb", 10, "aaa"), GetFullSig(MethodName)
>>>>>>> 524a3180e197414b591569bde713af7aadc226c6

ErrHandler:
    ExUnit.IsException ExceptionCode.ArgumentOutOfRange, Err.Number, GetSig(MethodName)
    
End Sub


Private Sub LastIndexOfTest()
    
    On Error GoTo ErrHandler
    Const MethodName = "LastIndexOfTest"
<<<<<<< HEAD
    ExUnit.AreEqual 5, StringH.LastIndexOf("bbbaaa", "a"), GetSig(MethodName)
    ExUnit.AreEqual 2, StringH.LastIndexOf("bbbaaa", "b"), GetSig(MethodName)
=======
    ExUnit.AreEqual 5, StringExt.LastIndexOf("bbbaaa", "a"), GetFullSig(MethodName)
    ExUnit.AreEqual 2, StringExt.LastIndexOf("bbbaaa", "b"), GetFullSig(MethodName)
>>>>>>> 524a3180e197414b591569bde713af7aadc226c6
    
    Exit Sub
ErrHandler:
    ExUnit.TestFailRunTime GetSig(MethodName)
    
End Sub


Private Sub LastIndexOfReturnsMinusOneWhenValueNotFound()
    
    On Error GoTo ErrHandler
    Const MethodName = "LastIndexOfReturnsMinusOneWhenValueNotFound"
<<<<<<< HEAD
    ExUnit.AreEqual -1, StringH.LastIndexOf("bbbaaa", "z"), GetSig(MethodName)
=======
    ExUnit.AreEqual -1, StringExt.LastIndexOf("bbbaaa", "z"), GetFullSig(MethodName)
>>>>>>> 524a3180e197414b591569bde713af7aadc226c6
    
    Exit Sub
ErrHandler:
    ExUnit.TestFailRunTime GetSig(MethodName)
    
End Sub


Private Sub LastIndexOfThrowsArgumentExceptionWhenStrIsNullStringTest()
    
    On Error GoTo ErrHandler
    Const MethodName = "LastIndexOfThrowsArgumentExceptionWhenStrIsNullStringTest"
    StringExt.LastIndexOf vbNullString, "a"
    
ErrHandler:
    ExUnit.IsException ExceptionCode.ArgumentException, Err.Number, GetSig(MethodName)

End Sub


Private Sub LastIndexOfThrowsArgumentOutOfRangeWhenStartIndexIsLesserThanZeroTest()
    
    On Error GoTo ErrHandler
    Const MethodName = "LastIndexOfThrowsArgumentOutOfRangeWhenStartIndexIsLesserThanZeroTest"
    StringExt.LastIndexOf "bbbaaa", "a", -1, Len("bbbaaa"), VbCompareMethod.vbTextCompare

ErrHandler:
    ExUnit.IsException ExceptionCode.ArgumentOutOfRange, Err.Number, GetSig(MethodName)

End Sub


Private Sub LastIndexOfThrowsArgumentOutOfRangeWhenStartIndexIsGreaterThanStrLen()
    
    On Error GoTo ErrHandler
    Const MethodName = "LastIndexOfThrowsArgumentOutOfRangeWhenStartIndexIsGreaterThanStrLen"
    StringExt.LastIndexOf "bbbaaa", "a", Len("bbbaaa") + 1, Len("bbbaaa"), VbCompareMethod.vbTextCompare

ErrHandler:
    ExUnit.IsException ExceptionCode.ArgumentOutOfRange, Err.Number, GetSig(MethodName)

End Sub


Private Sub LastIndexOfThrowsArgumentOutOfRangeWhenWhenCountIsLesserThanZeroTest()
    
    On Error GoTo ErrHandler
    Const MethodName = "LastIndexOfThrowsArgumentOutOfRangeWhenWhenCountIsLesserThanZeroTest"
    StringExt.LastIndexOf "bbbaaa", "a", Len("bbbaaa") - 1, -1, VbCompareMethod.vbTextCompare

ErrHandler:
    ExUnit.IsException ExceptionCode.ArgumentOutOfRange, Err.Number, GetSig(MethodName)

End Sub


Private Sub LastIndexOfThrowsArgumentOutOfRangeWhenStartIndexMinusCountIsLesserThanZeroTest()
    
    On Error GoTo ErrHandler
    Const MethodName = "LastIndexOfThrowsArgumentOutOfRangeWhenStartIndexMinusCountIsLesserThanZeroTest"
    StringExt.LastIndexOf "bbbaaa", "a", Len("bbbaaa") - 1, Len("bbbaaa") + 1, VbCompareMethod.vbTextCompare

ErrHandler:
    ExUnit.IsException ExceptionCode.ArgumentOutOfRange, Err.Number, GetSig(MethodName)

End Sub


Private Sub LastIndexOfThrowsThrowsArgumentOutOfRangeWhenWhenComparisonMethodIsDatabaseTest()
    
    On Error GoTo ErrHandler
    Const MethodName = "LastIndexOfThrowsThrowsArgumentOutOfRangeWhenWhenComparisonMethodIsDatabaseTest"
    StringExt.LastIndexOf "bbbaaa", "a", Len("bbbaaa") - 1, Len("bbbaaa"), VbCompareMethod.vbDatabaseCompare

ErrHandler:
    ExUnit.IsException ExceptionCode.ArgumentOutOfRange, Err.Number, GetSig(MethodName)

End Sub


Private Sub LastIndexOfAnyTest()

    On Error GoTo ErrHandler
    Const MethodName = "LastIndexOfAnyTest"
<<<<<<< HEAD
    ExUnit.AreEqual 0, StringH.LastIndexOfAny("abcd", StringArray("a", "d")), GetSig(MethodName)
    ExUnit.AreEqual 3, StringH.LastIndexOfAny("abcd", StringArray("d", "a")), GetSig(MethodName)
    ExUnit.AreEqual 1, StringH.LastIndexOfAny("abcd", StringArray("b", "c")), GetSig(MethodName)
=======
    ExUnit.AreEqual 0, StringExt.LastIndexOfAny("abcd", StringArray("a", "d")), GetFullSig(MethodName)
    ExUnit.AreEqual 3, StringExt.LastIndexOfAny("abcd", StringArray("d", "a")), GetFullSig(MethodName)
    ExUnit.AreEqual 1, StringExt.LastIndexOfAny("abcd", StringArray("b", "c")), GetFullSig(MethodName)
>>>>>>> 524a3180e197414b591569bde713af7aadc226c6
    
    Exit Sub
ErrHandler:
    ExUnit.TestFailRunTime GetSig(MethodName)
    
End Sub


Private Sub LastIndexOfAnyReturnsMinusWhenCharNotItStringTest()

    On Error GoTo ErrHandler
    Const MethodName = "LastIndexOfAnyReturnsZeroWhenCharNotItStringTest"
<<<<<<< HEAD
    ExUnit.AreEqual -1, StringH.LastIndexOfAny("ab", StringArray("z")), GetSig(MethodName)
=======
    ExUnit.AreEqual -1, StringExt.LastIndexOfAny("ab", StringArray("z")), GetFullSig(MethodName)
>>>>>>> 524a3180e197414b591569bde713af7aadc226c6
    
    Exit Sub
ErrHandler:
    ExUnit.TestFailRunTime GetSig(MethodName)

End Sub


Private Sub LastIndexOfAnyThrowsArgumentExceptionWhenAnyOfArrayIsNotInitalizedTest()

    On Error GoTo ErrHandler
    Const MethodName = "LastIndexOfAnyThrowsArgumentExceptionWhenAnyOfArrayIsNotInitalizedTest"
    Dim NotInit() As String
    StringExt.LastIndexOfAny "ab", NotInit
    
ErrHandler:
    ExUnit.IsException ExceptionCode.ArgumentException, Err.Number, GetSig(MethodName)
    
End Sub


Private Sub LastIndexOfAnyThrowsArgumentExceptionWhenAnyOfIsEmpty()
    
    On Error GoTo ErrHandler
    Const MethodName = "LastIndexOfAnyThrowsArgumentExceptionWhenAnyOfIsEmpty"
    Dim EmptyArr(0) As String
<<<<<<< HEAD
    ExUnit.AreEqual 0, StringH.LastIndexOfAny("ab", EmptyArr), GetSig(MethodName)
=======
    ExUnit.AreEqual 0, StringExt.LastIndexOfAny("ab", EmptyArr), GetFullSig(MethodName)
>>>>>>> 524a3180e197414b591569bde713af7aadc226c6
    
ErrHandler:
    ExUnit.IsException ExceptionCode.ArgumentException, Err.Number, GetSig(MethodName)
    
End Sub


