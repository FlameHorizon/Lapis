# CollectionExt.SelectOne Method

Returns the only element of a sequence that satisfies a specified condition, and throws an exception if more than one such element exists. If `Predicate` is not specified, returns the only element of a sequence, and throws an exception if there is not exactly one element in the sequence.

```vb
Public Function SelectOne(ByVal Source As Collection, Optional ByVal Predicate As ICallable) As Variant
```

### Parameters

**Source** `Collection` <br>
A Collection to return a single element from.

**Predicate** `ICallable` <br>
Optional. A function to test an element for a condition.

### Returns

`Variant` <br>
The single element of the input sequence that satisfies a condition.

### Errors

`OnArgumentNull` <br>
`Source` is `Nothing`

`OnInvalidOperation` <br>
When the number of elements in `Source` is zero

-or-

No element satisfies the condition in `Predicate`.

-or-

More than one element satisfies the condition in `Predicate`

## Examples

The following code example demonstrates how to use `SelectOne` to select the only element of a set.

```vb
Option Explicit

Public Sub Start()

    Dim Fruits As Collection
    Set Fruits = Lapis.CollectionExt.Make("orange")

    Dim Fruit1 As Variant
    Fruit1 = Lapis.CollectionExt.SelectOne(Fruits)
    
    Debug.Print Fruit1

End Sub

' This code produces the following output:
'
' orange
```

The following code example demonstrates that `SelectOne` throws an exception when the sequence does not contain exactly one element.

```vb
Option Explicit

Public Sub Start()

    On Error GoTo ErrHandler

    Dim Fruits As Collection
    Set Fruits = Lapis.CollectionExt.Make("orange", "apple")

    Dim Fruit1 As Variant
    Fruit1 = Lapis.CollectionExt.SelectOne(Fruits)
    
    Debug.Print Fruit1
    Exit Sub

ErrHandler:
    If Err.Number = ErrorCode.InvalidOperation Then
        Debug.Print "The collection does not contain exactly one element."
    End If

End Sub

' This code produces the following output:
'
' The collection does not contain exactly one element.
```

The following code example demonstrates how to use `SelectOne` to select the only element of a set that satisfies a condition.

```vb
Option Explicit

Public Sub Start()

    Dim Fruits As Collection
    Set Fruits = Lapis.CollectionExt.Make("apple", "banana", "mango", "orange", "passionfruit", "grape")

    Dim Fruit1 As Variant
    Fruit1 = Lapis.CollectionExt.SelectOne(Fruits, Lambda.Create("len($1) >= 10"))
    
    Debug.Print Fruit1

End Sub

' This code produces the following output:
'
' passionfruit
```

