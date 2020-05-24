# CollectionExt.Except Method

Produces the set difference of two sequences by using the specified IEqualityComparer to compare values.

```vb
Public Function Except(ByVal First As Collection, ByVal Second As Collection, Optional ByVal Comparer As IEqualityComparer) As Collection
```

### Parameters

**First** `Collection` <br>
A collection whose elements that are not also in `Second` will be returned.

**Second** `Collection` <br>
A collection whose elements that also occur in the first sequence will cause those elements to be removed from the returned sequence.

**Comparer** `IEqualityComparer` <br>
A comparer to compare items. Parameter can be omitted if item implements `IComparable` interface or item is value type.

### Returns

`Collection` <br>
A sequence that contains the set difference of the elements of two sequences.

### Errors

`OnArgumentNull` <br>
`First` is nothing

-or-

`Second` is nothing

## Examples

If you want to compare sequence of objects of some custom data type, you have to implement the IEqualityComparer interface in a helper class. The following code example shows how to implement this interface in a custom data type and provided Equals methods. The following example shows how to implement and equality comparer that can be used in the Except method.

```vb
' Product class
Option Explicit
Public Name As String
Public Code As Long
```

```vb
' ProductComparer class
' A custom comparer for the Product class
Option Explicit
Implements Lapis.IEqualityComparer

Public Function Equals(ByVal x As Product, ByVal y As Product) As Boolean

    If x Is y Then
        Equals = True
    
    ElseIf x Is Nothing Or y Is Nothing Then
        Equals = False
    
    Else
        Equals = (x.Code = y.Code And x.Name = y.Name)
    
    End If

End Function

Private Function IEqualityComparer_Equals(ByVal x As Variant, ByVal y As Variant) As Boolean
    IEqualityComparer_Equals = Me.Equals(x, y)
End Function
```

```vb
' Main module
Option Explicit
Public Sub Start()
    
    Dim Fruits1 As Collection
    Set Fruits1 = CollectionExt.Make(MakeProduct("apple", 9), _
                                     MakeProduct("orange", 4), _
                                     MakeProduct("lemon", 12))
    
    Dim Fruits2 As Collection
    Set Fruits2 = CollectionExt.Make(MakeProduct("apple", 9))
    
    ' Get all the elements from the first collection
    ' except for the elements from the second collection.
    
    Dim Except As Collection
    Set Except = CollectionExt.Except(Fruits1, Fruits2, New ProductComparer)
    
    Dim Item As Product
    For Each Item In Except
        Debug.Print Item.Name & " " & Item.Code
    Next Item
    
End Sub

Private Function MakeProduct(ByVal Name As String, ByVal Code As Long) As Product
    
    Dim Output As New Product
    Output.Name = Name
    Output.Code = Code
    Set MakeProduct = Output
    
End Function

' This code produces the following output:
' orange 4
' lemon 12
```

