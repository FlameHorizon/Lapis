# CollectionExt.Distinct Method

Returns distinct elements from a sequence by using a specified IEqualityComparer to compare values.

```vb
Public Function Distinct(ByVal Source As Collection, Optional ByVal Comparer As IEqualityComparer) As Collection
```

### Parameters

**Source** `Collection` <br>
The sequence to remove duplicate elements from.

**Comparer** `IEqualityComparer` <br>
An IEqualityComparer to compare values.

### Returns

`Collection` <br>
A set that contains distinct elements from the source sequence.

### Errors

`OnArgumentNull` <br>
`Source` or `Comparer` is `Nothing`

## Examples

The following example shows how to implement an equality comparer that can be used in the Distinct method.

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


Public Sub Start()

    Dim Products As Collection
    Set Products = Lapis.CollectionExt.Make(MakeProduct("Apple", 9), _
                                            MakeProduct("Orange", 4), _
                                            MakeProduct("Apple", 9), _
                                            MakeProduct("Lemon", 12))
    ' Exclude duplicates

    Dim NoDuplicates As Collection
    Set NoDuplicates = Lapis.CollectionExt.Distinct(Products, New ProductComparer)
    
    Dim Item As Variant
    For Each Item In NoDuplicates
        Debug.Print Item.Name & " " & Item.Code
    Next Item

End Sub

' This code produces the following output:
'
' Apple 9
' Orange 4
' Lemon 12

Private Function MakeProduct(ByVal Name As String, ByVal Code As Long) As Product
    
    Dim Output As New Product
    Output.Name = Name
    Output.Code = Code
    Set MakeProduct = Output
    
End Function
```

