# CollectionExt.Reverse Method

Inverts the order of the elements in a sequence.

```vb
Public Function Reverse(ByVal Source As Collection) As Collection
```

### Parameters

**Source** `Collection` <br>
A sequence of values to reverse.

### Returns

`Collection` <br>
A sequence whose elements correspond to those of the input sequence in reverse order.

### Errors

`OnArgumentNull` <br>
`Source` is `Nothing`

## Examples

The following code example demonstrates how to use Reverse to reverse the order of elements in an array.

```vb
' Start module
Option Explicit


Public Sub Start()

    Dim Apple As Collection
    Set Apple = CollectionExt.Make("a", "p", "p", "l", "e")
    
    Dim Reversed As Collection
    Set Reversed = CollectionExt.Reverse(Apple)
    
    Debug.Print CollectionExt.ToString(Reversed, Factory.GetValueTypeToStringConverter, " ")
    
End Sub
```

