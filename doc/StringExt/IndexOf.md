# StringExt.IndexOf Method

Returns the zero-based index of the first occurrence of the specified string in the current string

```vb
Public Function IndexOf(ByVal Str As String, ByVal Value As String, Optional ByVal StartIndex As Long = System.LongMinValue, Optional ByVal ComparisonMethod As VbCompareMethod = VbCompareMethod.vbTextCompare) As Long
```

### Parameters

**Str** `String` <br>
A source string which will be examined.

**Value** `String` <br>
The string to seek.

**StartIndex** `Long` <br>
The search starting position. If value is not provided, the default value will be 0. That means, search will start from the beginning of `Str`.

**ComparisonMethod** `VbCompareMethod` <br>
One of the enumeration values that specifies the rules for the search.

### Returns

`Long` <br>
The zero-based index position of the value parameter from the start of the current instance if that string is found, or -1 if it is not. If value is Empty, the return value is startIndex.

### Errors

`OnArgumentError` <br>
When `Str` is empty string.

`OnArgumentOutOfRangeException` <br>
`StartIndex` or `count` is smaller than 0.

-or-

`ComparisonMethod` is not `vbTextCompare` or `vbBinaryCompare`.

## Examples

The following example finds the first occurrence of either letter a in the string.

```vb
Option Explicit
Public Sub Start()
    Debug.Print "Case insensitive search: " & StringExt.IndexOf("aaAAaa", "a")
    Debug.Print "Case sensitive search: " & StringExt.IndexOf("aaAAaa", "A", , vbCompareMethod.vbBinaryCompare)
    Debug.Print "Case insensitive search: " & StringExt.IndexOf("aaAAaa", "a", 2, vbCompareMethod.vbTextCompare)
    Debug.Print "Case sensitive search: " & StringExt.IndexOf("aaAAaa", "a", 2, vbCompareMethod.vbBinaryCompare)

End Sub

' The example displays the following output:
' Case insensitive search: 0
' Case sensitive search: 2
' Case insensitive search with specified StartIndex: 2
' Case sensitive search with specified StartIndex: 4
```

### Remarks

Index numbering starts from 0 (zero). The startIndex parameter can range from 0 to the length of the string instance.
