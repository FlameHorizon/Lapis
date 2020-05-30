# CollectionExt.Sum Method

Computes the sum of a sequence of Int32 values.

```vb
Public Function Sum(ByVal Source As Collection, ByVal Selector As Lapis.IConverter) As Variant
```

### Parameters

**Source** `Collection` <br>
A sequence of values to calculate the sum of.

**Selector** `Lapis.IConverter` <br>
Method how elements of array will be converted into summable values.

### Returns

`Variant` <br>
The sum of the values in the sequence.

### Errors

`OnArgumentNull` <br>
`Source` is `Nothing`

-or-

`Selector` is `Nothing`

## Examples

The following code example demonstrates how to use Sum to sum the values of a sequence.

```vb
' Start module
Option Explicit


Public Sub Start()

    Dim Numbers As Collection
    Set Numbers = CollectionExt.Make(2, 3, 5, 10)
    
    Dim Sum As Long
    Sum = CollectionExt.Sum(Numbers, Factory.GetValueTypeToNumericConverter)
    
    Debug.Print "The sum fo the numbers is " & Sum
    
End Sub

' This code produces the following output:

' The sum fo the numbers is 20
```

