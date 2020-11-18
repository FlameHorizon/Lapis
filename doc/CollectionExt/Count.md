# CollectionExt.Count Method

Returns a number that represents how many elements in the specified sequence satisfy a condition.

```vb
Public Function Count(ByVal Source As Collection, ByVal Predicate As Predicate) As Long
```

### Parameters

**Source** `Collection` <br>
A sequence that contains elements to be tested and counted.

**Predicate** `Predicate` <br>
A function to test each element for a condition.

### Returns

`Long` <br>
A number that represents how many elements in the sequence satisfy the condition in the predicate function.

### Errors

`OnArgumentNull` <br>
`Source` or `Predicate` is `Nothing`

## Examples

The following code example demonstrates how to use Count(Collection, Pred) to count the elements in an array that satisfy a condition.

```
Option Explicit

Public Sub Start()

    Dim Numbers As Collection
    Set Numbers = Lapis.CollectionExt.Make(1, 2, 3, 4, 5, 6, 7, 8, 9, 10)

    Dim Pred As Lapis.Predicate
    Set Pred = Lapis.Factory.GetPredicate
    
    With Pred
        .Operator = Lapis.ComparisonOperator.GreaterThanOrEqualTo
        .ComparisonValue = 5
        .Comparer = Lapis.Factory.GetLongComparer
    End With
    
    Dim Value As Variant
    Value = Lapis.CollectionExt.Count(Numbers, Pred)
    
    Debug.Print "There are " & Value & " numbers greater than or equal to 5."

End Sub

' This code produces the following output:
'
' There are 6 numbers greater than or equal to 5.
```

