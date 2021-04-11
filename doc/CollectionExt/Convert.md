# CollectionExt.Convert Method

Projects each element of a sequence into a new form.

```vb
Public Function Convert(ByVal Source As Collection, ByVal Selector As ICallable) As Collection
```

### Parameters

**Source** `Collection` <br>
A sequence of values to invoke a transform function on.

**Selector** `ICallable` <br>
A transform function to apply to each element.

### Returns

`Collection` <br>
A set whose elements are the result of invoking the transform function on each element of `Source`

### Errors

`OnArgumentNull` <br>
`Source` or `Converter` is `Nothing`

## Examples

The following code example demonstrates how to use `Convert` to return a set of results.

```vb
Option Explicit

Public Sub Start()

    Dim Numbers As Collection
    Set Numbers = Lapis.CollectionExt.Range(1, 10)
    
    Dim Squares As Collection
    Set Squares = Lapis.CollectionExt.Convert(Numbers, Lambda.Create("$1 * $1"))

    Dim Item As Variant
    For Each Item In Squares
        Debug.Print Item
    Next Item

End Sub


' This code produces the following output:
'
' 1
' 4
' 9
' 16
' 25
' 36
' 49
' 64
' 81
' 100
```

