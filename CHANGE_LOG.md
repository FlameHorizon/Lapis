# Change log

## Changes in 0.9.0

* Added following methods to `StringExt`:

    * `Compare` - Compares two specified Strings, ignoring or honoring their case, and returns an integer that indicates their relative position in the sort order.

    * `Concat` - Concatenates the members of a constructed Variant ParamArray of type String.

    * `IsEmpty` - Indicates whether the specified string is null or an empty string (vbNullString).

    * `IsEmptyOrWhiteSpace` - Indicates whether a specified string is empty, or consists only of white-space characters.

    * `PadLeft` - Returns a new string that right-aligns the characters in this instance by padding them on the left with a specified Unicode character, for a specified total length.

    * `PadRight` - Returns a new string that left-aligns the characters in this string by padding them on the right with a specified Unicode character, for a specified total length.

    * `Replace` - Returns a new string in which all occurrences of a specified string in the current instance are replaced with another specified string, using the provided comparison type.

    * `Substring` - Retrieves a substring from this instance. The substring starts at a specified character position and has a specified length. If length is not specified, selection continues to the end of the string.

    * `ToLower` - Returns a copy of this string converted to lowercase.

    * `ToUpper` - The uppercase equivalent of the current string.

    * `Trim` - Removes all leading and trailing occurrences of a set of characters specified in an array from the current string.

    * `TrimEnd` - Removes all the trailing occurrences of a set of characters specified in an array from the current string.

    * `TrimStart` - Removes all the leading occurrences of a set of characters specified in an array from the current string.

* Changed signature of `StringExt.IndexOf` method. `Count` argument is no longer accepted.

* Methods are now listed in alphabetical order. This should allow easier navigation though module documentation.

* Fixed casing for boolean literals in documentation.

* Fixed a bug where `StringExt.StartsWith` would not pickup symbols like `#` and `*`.

## Changes in 0.8.0

* Updated or created entries in `CollectionExt` documentation.

* Added following method to `CollectionExt`:

    * Convert

* Removed following methods from `CollectionExt`:
    
    * ToString

    * ToStringByProperty

    * DistinctValues

* Fixed a bug in MergeSort algorithm which would not allow to sort items which aren't reference type items.

## Changes in 0.7.0

* ExUnit, instead of displaying error numbers when test fails, now shows name of the error. This applies for both, custom defined exceptions in `Exception` class as well as for VBA's error numbers.

```vb
Option Explicit
Private Sub RemoveThrowsArgumentOutOfRangeWhenStartIndexIsNegetiveTest()

    On Error GoTo ErrHandler
    Const MethodName = "RemoveThrowsArgumentOutOfRangeWhenStartIndexIsNegetiveTest"
    ExUnit.AreEqual "foo", StringExt.Remove("foo-bar", -1), "StringHelperTests." & MethodName

ErrHandler:
    ExUnit.IsException ExceptionCode.ArgumentOutOfRange, Err.Number, "StringHelperTests." & MethodName

End Sub

' This example produces the following results:
'
' FAILED; Expected exception [ArgumentOutOfRange], Actual exception [ArgumentNull]; StringHelperTests.RemoveThrowsArgumentOutOfRangeWhenStartIndexIsNegetiveTest
```

* Added `ExceptionCodeEnum.ToString` and `ExceptionCodeEnum.TryToString` to retrieve custom exception names.

* Added `ErrorNumber` enumerator which lists error numbers for every built-in VBA error.

* Added `ErrorNumberEnum.ToString` and `ErrorNumberEnum.TryToString` to retrieve error names for built-in VBA errors.

* TODO: Tests which do fail assertion should have more better descriptions.

* Instead of displaying each assertion result separately, `ExUnit` now groups test results based on the source. Each test method is represented by one `TestResult` object.

> Remarks: This behavior is expected only when ExUnit is setup like in following example.

```vb
Option Explicit
Public Sub TestRunner()

    ExUnit.ShowPassingTest = True
    ExUnit.Setup

    TestMethod

    ExUnit.PrintTestResults New TestResultImmediatePrinter
    ExUnit.Teardown

End Sub


Private Sub TestMethod()

    On Error GoTo ErrHandler
    Const MethodName = "TestMethod"
    ExUnit.AreEqual 1, 1, "TestModule." & MethodName
    ExUnit.AreEqual 2, 2, "TestModule." & MethodName

    Exit Sub
ErrHandler:
    ExUnit.TestFailRunTime "TestModule." & MethodName

End Sub

' This example produces the following results:
'
' PASSED; TestModule.TestMethod
```

* Users can now have access to the test results by calling `ExUnit.TestResults` at the end of each test run.

* ExUnit can now print test results into any place which is handled by `ITestResultPrinter` interface. At moment there is only one class which impelemnts mentioned interface, it's `TestResultImmediatePrinter`.

---

## Changes in 0.6.0

* Exposed `ExUnit.TestPass` and `ExUnit.TestFail` methods to better express some test's intentions.
* Added `StringExt.Remove` which returns new string in which a specified number of characters from the current string are deleted.

```vb
Public Sub Start()
    Debug.Print StringExt.Remove("foo-bar", 3)
End Sub

' This example produces the following results:
'
' foo
```

* Added `StringExt.RemoveRange` which returns a new string in which a specified number of characters in the current instance beginning at a specified position have been deleted.

```vb
Option Explicit
Public Sub Start()
    Debug.Print StringExt.RemoveRange("foo-bar", 3, 1)
End Sub

' This example produces the following results:
'
' foobar
```

* Added `StringExt.Concat` which concatenates elements of a specified string array.

```vb
Option Explicit
Public Sub Start()
    Debug.Print StringExt.Concat("A", "bbb", "c", "12#")
End Sub

' This example produces the following results:
'
' Abbbc12#
```

* Added `StringExt.StartsWith` which determines whether the beginning of this string instance matches a specified string.

```vb
Option Explicit
Public Sub Start()

    Test("John", "J", VbCompareMethod.vbTextCompare)
    Test("John", "J", VbCompareMethod.vbBinaryCompare)

End Sub


Private Sub Test(ByVal x as String, ByVal y as String, ByVal CompareMethod as vbCompareMethod)

    If CompareMethod = vbTextCompare Then
        Debug.Print "vbCompareMethod.vbTextCompare"

    ElseIf CompareMethod = vbBinaryCompare Then
        Debug.Print "vbCompareMethod.vbBinaryCompare"
    End if

    If StringExt.StartsWith(x, y, CompareMethod) Then
        Debug.Print x & " ends with " & y
    Else
        Debug.Print x & " does not ends with " & y
    End if

    Debug.Print vbNewLine

End Sub

' This example produces the following results:
'
' vbCompareMethod.vbTextCompare
' John starts with J
'
' vbCompareMethod.vbBinaryCompare
' John does not starts with J
```

* Added `StringExt.EndsWith` which determines whether the end of this string instance matches the specified string when compared using the specified comparison option.

```vb
Option Explicit
Public Sub Start()

    Test("John", "N", VbCompareMethod.vbTextCompare)
    Test("John", "N", VbCompareMethod.vbBinaryCompare)

End Sub


Private Sub Test(ByVal x as String, ByVal y as String, ByVal CompareMethod as vbCompareMethod)

    If CompareMethod = vbTextCompare Then
        Debug.Print "vbCompareMethod.vbTextCompare"

    ElseIf CompareMethod = vbBinaryCompare Then
        Debug.Print "vbCompareMethod.vbBinaryCompare"
    End if

    If StringExt.EndsWith(x, y, CompareMethod) Then
        Debug.Print x & " ends with " & y
    Else
        Debug.Print x & " does not ends with " & y
    End if
    Debug.Print vbNewLine

End Sub

' This example produces the following results:
'
' vbCompareMethod.vbTextCompare
' John ends with N
'
' vbCompareMethod.vbBinaryCompare
' John does not ends with N
```

* Added `StringExt.Insert` which returns a new string in which a specified string is inserted at a specified index position in the instance.

```vb
Option Explicit
Public Sub Start()

    Dim Original as String: Original = "aaabbb"
    Dim Modified as String
    Modified = StringExt.Insert(Original, 3, " ")
    Debug.Print "This is original string: " & Original
    Debug.Print "This is modified string: " & Modified

End Sub

' This example produces the following results:
'
' This is original string: aaabbb
' This is modified string: aaa bbb
```

* Added `StringExt.LastIndexOf` which reports the zero-based index position of the last occurrence of a specified string within this instance.

```vb
Option Explicit
Public Sub Start()

    Debug.Print "Comparison: vbTextCompare. Location: " & StringExt.LastIndexOf(Text, "b", 0, Len(Text), vbCompareMethod.vbTextCompare
    Debug.Print "Comparison: vbBinaryCompare. Location:  " & StringExt.LastIndexOf(Text, "B", 0, Len(Text), vbCompareMethod.vbBinaryCompare)

End Sub

' This example produces the following results:
'
' Comparison: vbTextCompare. Location: 3
' Comparison: vbBinaryCompare. Location: -1
```

* Added `StringExt.LastIndexOfAny` which reports the zero-based index position of the last occurrence in this instance of one or more characters.

```vb
Option Explicit
Public Sub Start()

    Debug.Print "The last character occurrence location: " & StringExt.LastIndexOfAny("abcd", StringArray("a", "d"))
    Debug.Print "The last character occurrence location: " & StringExt.LastIndexOfAny("abcd", StringArray("d", "a"))
    Debug.Print "The last character occurrence location: " & StringExt.LastIndexOfAny("abcd", StringArray("b", "c"))

End Sub

' This example produces the following results:
' The last character occurrence location: 0
' The last character occurrence location: 3
' The last character occurrence location: 1
```
