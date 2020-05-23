# StringExt.LastIndexOf Method

Returns the zero-based index of the last occurrence of the specified string in the current string.

```vb
Public Function LastIndexOf(ByVal Str As String, ByVal Value As String, Optional ByVal StartIndex As Long = System.LongMinValue, Optional ByVal Count As Long = System.LongMinValue, Optional ByVal ComparisonMethod As VbCompareMethod = VbCompareMethod.vbTextCompare) As Long
```

### Parameters

**Str** `String` <br>
A source string which will be examined.

**Value** `String` <br>
The string to seek.

**StartIndex** `Long` <br>
The search starting position. The search proceeds from startIndex toward the beginning of this instance.

**Count** `Long` <br>
The number of character positions to examine. If value is not provided, the default value will be `Len(Str) - StartIndex`.

**ComparisonMethod** `VbCompareMethod` <br>
One of the enumeration values that specifies the rules for the search.

### Returns

`Long` <br>
The zero-based index position of the value parameter from the start of the current instance if that string is found, or -1 if it is not. If value is `vbEmptyString`, the return value is `StartIndex`.

### Errors

`OnArgumentError` <br>
When `Str` is empty string.

-or-

`Value` is empty string.

`ArgumentOutOfRangeException` <br>
`StartIndex` or `Count` is smaller than 0.

-or-

`StartIndex` is greater than the length of `Str`

-or-

`Len(Str) - Count` is smaller than `StartIndex`

-or-

`ComparisonMethod` is not `vbTextCompare` or `vbBinaryCompare`.

## Examples

The following example demonstrates how last index of letter a can be found.

```vb
Option Explicit
Public Sub Start()

    Debug.Print "Case insensitive search: " & StringExt.LastIndexOf("aaAAaa", "a")
    Debug.Print "Case sensitive search: " & StringExt.LastIndexOf("aaAAaa", "A", , , VbCompareMethod.vbBinaryCompare)
    Debug.Print "Case insensitive search: " & StringExt.LastIndexOf("aaAAaa", "a", 5, , VbCompareMethod.vbTextCompare)
    Debug.Print "Case sensitive search: " & StringExt.LastIndexOf("aaAAaa", "A", 5, , VbCompareMethod.vbBinaryCompare)

End Sub

' The example displays the following output:
' Case insensitive search: 0
' Case sensitive search: 2
' Case insensitive search with specified StartIndex: 2
' Case sensitive search with specified StartIndex: 4
```

### Remarks

Index numbering starts from zero. That is, the first character in the string is at index zero and the last is at Length - 1.
