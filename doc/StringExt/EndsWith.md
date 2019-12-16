# EndsWith

Determines whether the end of this string instance matches a specified string.

```vb
Public Function EndsWith(ByVal Str As String, ByVal Value As String, ByVal StringComparison As VbCompareMethod) As Boolean
```

### Parameters

**Str** `String` <br>
A source string which will be examined.

**Value** `String` <br>
The string to seek.

**ComparisonMethod** `VbCompareMethod` <br>
One of the enumeration values that specifies the rules for the search.

### Returns

`Long` <br>
`true` if value matches the end of this instance; otherwise, `false`.

### Exceptions

`ArgumentException` <br>
When `Str` is empty string.

-or-

`ComparisonMethod` is not `vbTextCompare` or `vbBinaryCompare`.

## Examples

The following example indicates whether each string in an array ends with a period (".").

```vb
Option Explicit
Public Sub Start()

    Dim Arr(2) As String
    Arr(0) = "This is a string."
    Arr(1) = "Hello!"
    Arr(2) = "Nothing"

    Dim Item As Variant
    For Each Item In Arr
        Dim EndsWithPeriod As Boolean
        EndsWithPeriod = StringH.EndsWith(Item, ".", VbCompareMethod.vbTextCompare)
        Debug.Print Item & " ends with a period " & EndsWithPeriod
    Next Item

End Sub

' The example displays the following output:
' This is a string. ends with a period True
' Hello! ends with a period False
' Nothing ends with a period False
```

### Remarks

This method compares value to the substring at the end of this instance that is the same length as value, and returns an indication whether they are equal. To be equal, value must be a reference to this same instance or match the end of this instance.

This method performs a word (case-sensitive and culture-sensitive) comparison.