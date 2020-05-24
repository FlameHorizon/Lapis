# CollectionExt.Make Method

Creates a new collection based on to list of arguments.

```vb
Public Function Make(ParamArray Items() As Variant) As Collection
```

### Parameters

**Items** `ParamArray Variant` <br>
Array of items which new collection will contain.

### Returns

`Collection` <br>
New instance of a collection which contains given items

## Examples

The following example takes three strings and puts them into one collection.

```vb
Option Explicit
Public Sub Start()

    Dim Cars As Collection
    Set Cars = CollectionExt.Make("Toyota", "Porsche", "Fiat")

    Dim Item As Variant
    For Each Item In Cars
        Debug.Print Item
    Next Item

End Sub

' The example displays the following output:
' Toyota
' Porsche
' Fiat 

```

