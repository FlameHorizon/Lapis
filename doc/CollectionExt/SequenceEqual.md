# CollectionExt.SequenceEqual Method

Determines whether two sequences are equal by comparing their elements by using a specified IEqualityComparer.

```vb
Public Function SequenceEqual(ByVal First As Collection, ByVal Second As Collection, ByVal Comparer As IEqualityComparer) As Boolean
```

### Parameters

**First** `Collection` <br>
A set to compare to `Second`.

**Second** `Collection` <br>
A set to compare to the first sequence.

**Comparer** `IEqualityComparer` <br>
An `IEqualityComparer` to use to compare elements.

### Returns

`Boolean` <br>
`True` if the two source sequences are of equal length and their corresponding elements compare equal according to comparer; otherwise, `False`.

### Errors

`OnArgumentNull` <br>
`First`, `Second` or `Converter` is `Nothing`

## Examples

The following example shows how to implement an equality comparer that can be used in the SequenceEqual(Collection, Collection, IEqualityComparer) method.

```vb
' Product class
Option Explicit

Public Name As String
Public Code As Long
```

```vb
' ProductComparer class
Option Explicit

Implements Lapis.IEqualityComparer

' Determines whether the specified objects are equal.
Public Function Equals(ByVal x As Variant, ByVal y As Variant) As Boolean

    If x Is Nothing Or y Is Nothing Then
        Equals = False
        Exit Function
    End If
    
    Equals = (x.Name = y.Name) And (x.Code = y.Code)
    
End Function


Private Function IEqualityComparer_Equals(ByVal x As Variant, ByVal y As Variant) As Boolean
    IEqualityComparer_Equals = Me.Equals(x, y)
End Function
```

```vb
Option Explicit
' Main module

Public Sub Start()

    Dim StoreA As Collection
    Set StoreA = Lapis.CollectionExt.Make(MakeProduct("Apple", 9), _
                                          MakeProduct("Orange", 4))
    
    Dim StoreB As Collection
    Set StoreB = Lapis.CollectionExt.Make(MakeProduct("Apple", 9), _
                                          MakeProduct("Orange", 4))
    
    Dim EqualAB As Boolean
    EqualAB = Lapis.CollectionExt.SequenceEqual(StoreA, StoreB, New ProductComparer)
    
    Debug.Print "Equal? " & EqualAB

End Sub

' This code produces the following output:
'
' Equal? True


Private Function MakeProduct(ByVal Name As String, ByVal Code As Long) As Product

    Dim Output As New Product
    Output.Name = Name
    Output.Code = Code
    Set MakeProduct = Output

End Function

```

