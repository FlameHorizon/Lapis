# CollectionExt.Last Method

Returns the last element of a sequence.

```vb
Public Function Last(ByVal Source As Collection, ByVal Predicate As Predicate) As Variant
```

### Parameters

**Source** `Collection` <br>
A collection to return an element from.

**Predicate** `Predicate` <br>
A function to test each element for a condition.

### Returns

`Variant` <br>
The last element in the sequence that passes the test in the specified predicate function.

### Errors

`OnArgumentNull` <br>
`Source` or `Predicate` is `Nothing`

`OnInvalidOperation` <br>
When the number of elements in Source is zero

-or-

No element satisfies the condition in predicate.

## Examples

The following code example demonstrates how to use Last(Source, Predicate) to return the last element of a collection that satisfies a condition.

```
Option Explicit

Public Sub Start()

    Dim Numbers As Collection
    Set Numbers = Lapis.CollectionExt.Make(9, 34, 65, 92, 87, 435, 3, 54, _
                                           83, 23, 87, 67, 12, 19)
    
    Dim Pred As Lapis.Predicate
    Set Pred = Lapis.Factory.GetPredicate
    
    With Pred
        .Operator = Lapis.ComparisonOperator.GreaterThan
        .ComparisonValue = 80
        .Comparer = Lapis.Factory.GetLongComparer
    End With
    
    Dim Last As Variant
    Last = Lapis.CollectionExt.Last(Numbers, Pred)
    
    Debug.Print Last

End Sub

' This code produces the following output:
'
' 87
```

