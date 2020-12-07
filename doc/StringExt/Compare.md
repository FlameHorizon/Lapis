# StringExt.Compare Method

Compares two specified Strings, ignoring or honoring their case, and returns an integer that indicates their relative position in the sort order.

```vb
Public Function Compare(ByVal Str1 As String, ByVal Str2 As String, ByVal IgnoreCase As Boolean) As Long
```

### Parameters

**Str1** `String` <br>
The first string to compare.

**Str2** `String` <br>
The second string to compare.

**IgnoreCase** `Boolean` <br>
`True` to ignore case during the comparison; otherwise, `False`.

### Returns

`Long` <br>
A long value that indicates the lexical relationship between the two comparands.

|Value|Condition|
|---|---|
|Less than zero|`strA` precedes `strB` in the sort order|
|Zero|`strA` occurs in the same position as `strB` in the sort order.|
|Greater than zero|`strA` follows `strB` in the sort order.|

## Examples

The following example demonstrates the Compare(String, String, Boolean) method in use.

```vb
Option Explicit
' Main Module

Public Sub Start()

    Dim StringUpper As String: StringUpper = "ABC"
    Dim StringLower As String: StringLower = "abc"

    Debug.Print "Comparing " & StringUpper & " and " & StringLower & "."

    Debug.Print "The Strings are equal when capitalized? " & _
                IIf(StringExt.Compare(UCase(StringUpper), UCase(StringLower), IgnoreCase:=False) = 0, "true", "false")
                
    Debug.Print "The Strings are equal when case is ignored? " & _
                IIf(StringExt.Compare(StringUpper, StringLower, IgnoreCase:=True) = 0, "true", "false")
                
End Sub

' The example displays the following output:
' Comparing ABC and abc.
' The Strings are equal when capitalized? true
' The Strings are equal when case is ignored? true
```

