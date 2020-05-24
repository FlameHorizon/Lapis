# StringExt.Concat Method

Concatenates the elements of a specified String array.

```vb
Public Function Concat(ParamArray Items() As Variant) As String
```

### Parameters

**Items** `ParamArray Variant` <br>
An array of strings to join.

### Returns

`String` <br>
The concatenated elements of `Items`.

## Examples

The following example takes three strings and joins them into one.

```vb
Option Explicit
Public Sub Start()
    Debug.Print StringExt.Concat("'01", "-", "114'")
End Sub

' The example displays the following output:
' '01-114'
```

