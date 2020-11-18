# CollectionExt.SelectOne Method

Returns the only element of a sequence that satisfies a specified condition, and throws an exception if more than one such element exists.

```vb
Public Function SelectOne(ByVal Source As Collection, ByVal Predicate As Predicate) As Variant
```

### Parameters

**Source** `Collection` <br>
A Collection to return a single element from.

**Predicate** `Predicate` <br>
A function to test an element for a condition.

### Returns

`Variant` <br>
The single element of the input sequence that satisfies a condition.

### Errors

`OnArgumentNull` <br>
`Source` or `Predicate` is `Nothing`

`OnInvalidOperation` <br>
When the number of elements in `Source is zero

-or-

No element satisfies the condition in `predicate`.

-or-

More than one element satisfies the condition in `predicate`

## Examples

The following code example demonstrates how to use SelectOne(Collection, Predicate) to select the only element of a set that satisfies a condition.

```
Option Explicit

Public Sub Start()

    Dim Fruits As Collection
    Set Fruits = Lapis.CollectionExt.Make("apple", "banana", "mango", "orange", "passionfruit", "grape")

    Dim Pred As Lapis.Predicate
    Set Pred = Lapis.Factory.GetPredicate
    
    With Pred
        .Operator = Lapis.ComparisonOperator.GreaterThanOrEqualTo
        .ComparisonValue = 10
        .Comparer = Lapis.Factory.GetWordLengthComparer
    End With
    
    Dim Fruit1 As Variant
    Fruit1 = Lapis.CollectionExt.SelectOne(Fruits, Pred)
    
    Debug.Print Fruit1

End Sub

' This code produces the following output:
'
' passionfruit
```

