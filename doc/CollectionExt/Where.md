# CollectionExt.Where Method

Filters a sequence of values based on a predicate

```vb
Public Function Where(ByVal Source As Collection, ByVal Predicate As Predicate) As Collection
```

### Parameters

**Source** `Collection` <br>
A sequence to filter.

**Predicate** `Predicate` <br>
A function to test each element for a condition.

### Returns

`Collection` <br>
A set that contains elements from the input sequence that satisfy the condition.

### Errors

`OnArgumentNull` <br>
`Source` or `Predicate` is `Nothing`

## Examples

The following code example demonstrates how to use Where(Collection, Predicate) to filter a sequence.

```
Option Explicit

Public Sub Start()

    Dim Fruits As Collection
    Set Fruits = Lapis.CollectionExt.Make("apple", "passionfruit", "banana", "mango", _
                                         "orange", "blueberry", "grape", "strawberry")
    
    Dim Pred As Lapis.Predicate
    Set Pred = Lapis.Factory.GetPredicate
    With Pred
        .ComparisonValue = 6
        .Operator = Lapis.ComparisonOperator.LessThan
        .Comparer = Lapis.Factory.GetWordLengthComparer
    End With
    
    Dim Query As Collection
    Set Query = Lapis.CollectionExt.Where(Fruits, Pred)
    
    Dim Fruit As Variant
    For Each Fruit In Query
        Debug.Print Fruit
    Next Fruit

End Sub

' This code produces the following output:
' apple
' mango
' grape
```

