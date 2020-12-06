# StringExt.PadRight Method

Returns a new string that left-aligns the characters in this string by padding them on the right with a specified Unicode character, for a specified total length.

```vb
Public Function PadRight(ByVal Str As String, ByVal TotalWidth As Long, ByVal PaddingChar As String) As String
```

### Parameters

**Str** `String` <br>
A source string.

**TotalWidth** `Long` <br>
The number of characters in the resulting string, equal to the number of original characters plus any additional padding characters.

**PaddingChar** `String` <br>
A Unicode padding character.

### Returns

`String` <br>
A new string that is equivalent to this instance, but left-aligned and padded on the right with as many `paddingChar` characters as needed to create a length of `totalWidth`. However, if `totalWidth` is less than the length of this instance, the method returns a reference to the existing instance. If `totalWidth` is equal to the length of this instance, the method returns a new string that is identical to this instance.

### Errors

`OnArgumentOutOfRange` <br>
`TotalWidth` is less than zero

## Examples

The following example demonstrates the PadRight method.

```vb
Option Explicit

Public Sub Start()

    Dim Str As String: Str = "forty-two"
    Dim Pad As String: Pad = "."
    
    Debug.Print StringExt.PadRight(Str, 15, Pad)
    Debug.Print StringExt.PadRight(Str, 2, Pad)

End Sub

' The example displays the following output:
'
' forty-two......
' forty-two

```

