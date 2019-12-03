# Insert

Returns a new string in which a specified string is inserted at a specified index position in this instance.

```vb
Public Function Insert(ByVal Str As String, ByVal StartIndex As Long, ByVal Value As String) As String
```

### Parameters

**Str** `String` <br>
The string to which value will be inserted.

**StartIndex** `Long` <br>
The zero-based index position of the insertion.

**Value** `String` <br>
The string to insert

### Returns

`String` <br>
A new string that is equivalent to this instance, but with value inserted at position startIndex.

### Exceptions

`ArgumentException` <br>
When `Str` is empty string.

`ArgumentOutOfRangeException` <br>
`StartIndex` is negative or greater than the length of this instance.

## Examples

The following example inserts a space character in the fourth character position (the character at index 3) of a string.

```vb
Option Explicit
Public Sub Start()

    Dim Original As String: Original = "aaabbb"
    Debug.Print "The original string " & Original
    
    Dim Modified As String: Modified = StringH.Insert(Original, 3, " ")
    Debug.Print "The modified string: " & Modified

End Sub

' The example displays the following output:
' The original string: aaabbb
' The modified string: aaa bbb
```

### Remarks

If `startIndex` is equal to the length of this instance, value is appended to the end of this instance.