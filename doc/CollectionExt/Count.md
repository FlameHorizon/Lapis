# CollectionExt.Count Method

Returns a number that represents how many elements in the specified sequence satisfy a condition. If `Predicate` is not defined, returns the number of elements in a sequence.

```vb
Public Function Count(ByVal Source As Collection, Optional ByVal Predicate As ICallable) As Long
```

### Parameters

**Source** `Collection` <br>
A sequence that contains elements to be tested and counted.

**Predicate** `ICallable` <br>
Optional. A function to test each element for a condition.

### Returns

`Long` <br>
A number that represents how many elements in the sequence satisfy the condition in the predicate function.

### Errors

`OnArgumentNull` <br>
`Source` is `Nothing`

## Examples

The following code example demonstrates how to use `Count` to count the elements in a set.

```vb
Option Explicit

Public Sub Start()

    Dim Numbers As Collection
    Set Numbers = Lapis.CollectionExt.Make(1, 2, 3, 4, 5, 6, 7, 8, 9, 10)

    Dim Value As Variant
    Value = Lapis.CollectionExt.Count(Numbers)
    
    Debug.Print "There are " & Value & " numbers."

End Sub

' This code produces the following output:
'
' There are 10 numbers.
```

The following code example demonstrates how to use `Count` to count the elements in a set that satisfy a condition.

```vb
Option Explicit

Public Sub Start()

    Dim Numbers As Collection
    Set Numbers = Lapis.CollectionExt.Make(1, 2, 3, 4, 5, 6, 7, 8, 9, 10)

    Dim Value As Variant
    Value = Lapis.CollectionExt.Count(Numbers, Lambda.Create("$1 >= 5"))
    
    Debug.Print "There are " & Value & " numbers greater than or equal to 5."

End Sub

' This code produces the following output:
'
' There are 6 numbers greater than or equal to 5.
```

