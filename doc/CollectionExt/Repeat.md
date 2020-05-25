# CollectionExt.Repeat Method

Generates a sequence that contains one repeated value.

```vb
Public Function Repeat(ByVal Element As Variant, ByVal Count As Long) As Collection
```

### Parameters

**Element** `Variant` <br>
The value to be repeated.

**Count** `Long` <br>
The number of times to repeat the value in the generated sequence.

### Returns

`Collection` <br>
A collection that contains a repeated value.

### Errors

`OnArgumentOutOfRange` <br>
`Count` is less than 0.

## Examples

The following code example demonstrates how to use Repeat to generate a sequence of a repeated value.

```vb
' Main module
Option Explicit

Public Sub Start()

    Dim Strings As Collection
    Set Strings = CollectionExt.Repeat("I like programming", 10)
    
    Dim Str As Variant
    For Each Str In Strings
        Debug.Print Str
    Next Str
    
End Sub


' The example displays the following output:
' I like programming
' I like programming
' I like programming
' I like programming
' I like programming
' I like programming
' I like programming
' I like programming
' I like programming
' I like programming
```

