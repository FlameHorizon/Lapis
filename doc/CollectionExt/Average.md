# CollectionExt.Average Method

Computes the average of a sequence of values that is obtained by invoking a projection function on each element of the input sequence.

```vb
Public Function Average(ByVal Source As Collection, ByVal Selector As Lapis.IConverter) As Variant
```

### Parameters

**Source** `Collection` <br>
A sequence of values that are used to calculate an average.

**Selector** `Lapis.IConverter` <br>
A projection function to apply to each element.

### Returns

`Variant` <br>
The average of the sequence of values.

### Errors

`OnArgumentNull` <br>
`Source` is `Nothing`

-or-

`Selector` is `Nothing`

## Examples

' Start module
Option Explicit


Public Sub Start()

    Dim Fruits As Collection
    Set Fruits = CollectionExt.Make("apple", "banana", "mango", "orange", "passionfruit", "grape")
    
    Dim Average As Double
    Average = CollectionExt.Average(Fruits, New StringToLengthConverter)
    
    Debug.Print "The average string length is " & Average
    
End Sub

' This code produces the following output:

' The average string length is 6,5

