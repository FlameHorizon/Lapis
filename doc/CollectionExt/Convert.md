# CollectionExt.Convert Method

Projects each element of a sequence into a new form.

```vb
Public Function Convert(ByVal Source As Collection, ByVal Converter As Lapis.IConverter) As Collection
```

### Parameters

**Source** `Collection` <br>
A sequence of values to invoke a transform function on.

**Converter** `Lapis.IConverter` <br>
A transform function to apply to each element.

### Returns

`Collection` <br>
A set whose elements are the result of invoking the transform function on each element of `Source`

### Errors

`OnArgumentNull` <br>
`Source` or `Converter` is `Nothing`

## Examples

The following code example demonstrates how to use Convert(Collection, Lapis.IConverter) to force immediate query evaluation and return a set of results.

```vb
' MathSquareConverter class module
Option Explicit

Implements Lapis.IConverter


Public Function Convert(ByVal Item As Variant) As Variant
    Convert = Item * Item
End Function


Private Function IConverter_Convert(ByVal Item As Variant) As Variant
    IConverter_Convert = Me.Convert(Item)
End Function
```

```vb
Option Explicit

Public Sub Start()

    Dim Numbers As Collection
    Set Numbers = Lapis.CollectionExt.Range(1, 10)
    
    Dim Squares As Collection
    SetSquares = Lapis.CollectionExt.Convert(Numbers, New MathSquareConverter)

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

