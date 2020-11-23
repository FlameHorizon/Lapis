# CollectionExt.Contains Method

Determines whether a sequence contains a specified element.

```vb
Public Function Contains(ByVal Source As Collection, ByVal Value As Variant, Optional ByVal Comparer As IEqualityComparer) As Boolean
```

### Parameters

**Source** `Collection` <br>
A sequence in which to locate a value.

**Value** `Variant` <br>
The value to locate in the sequence.

**Comparer** `IEqualityComparer` <br>
An equality comparer to compare values.

### Returns

`Boolean` <br>
`True` if the source sequence contains an element that has the specified value; otherwise, `False`

### Errors

`OnArgumentNull` <br>
`Source` is `Nothing`

## Examples

The following example shows how to implement an equality comparer that can be used in the Contains method.

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
Public Sub Start()

    Dim Fruits As Collection
    Set Fruits = Lapis.CollectionExt.Make(MakeProduct("Apple", 9), _
                                          MakeProduct("Orange", 4), _
                                          MakeProduct("Lemon", 12))
    
    Dim Apple As Product: Set Apple = MakeProduct("Apple", 9)
    Dim Kiwi As Product: Set Kiwi = MakeProduct("Kiwi", 8)
    Dim ProdComp As New ProductComparer
    
    Dim HasApple As Boolean: HasApple = Lapis.CollectionExt.Contains(Fruits, Apple, ProdComp)
    Dim HasKiwi As Boolean: HasKiwi = Lapis.CollectionExt.Contains(Fruits, Kiwi, ProdComp)
    
    Debug.Print "Apple? " & HasApple
    Debug.Print "Kiwi? " & HasKiwi

End Sub

' This code produces the following output:
'
' Apple? True
' Kiwi? False

Private Function MakeProduct(ByVal Name As String, ByVal Code As Long) As Product
    
    Dim Output As New Product
    Output.Name = Name
    Output.Code = Code
    Set MakeProduct = Output
    
End Function
```

