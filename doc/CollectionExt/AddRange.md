# CollectionExt.AddRange Method

Adds the elements of the specified collection to the end of the set.

```vb
Public Sub AddRange(ByRef Container As Collection, ByVal Items As Collection)
```

### Parameters

**Container** `Collection` <br>
The collection to which elements will be added.

**Items** `Collection` <br>
The collection whose elements should be added to the end of the collection. The collection itself cannot be null, but it can contain elements that are `Nothing`

### Errors

`OnArgumentNull` <br>
`Container` is `Nothing

-or-

`Items` is `Nothing`

## Examples

The following example demonstrates the AddRange method that act on ranges.

```vb
Option Explicit

Public Sub Start()

    Dim Dinosaurs As Collection
    Set Dinosaurs = Lapis.CollectionExt.Make("Brachiosaurus", "Amargasaurus", "Mamenchisaurus")

    Dim Item As Variant
    For Each Item In Dinosaurs
        Debug.Print Item
    Next Item
    
    Dim MoreDinosaurs As Collection
    Set MoreDinosaurs = Lapis.CollectionExt.Make("Tyrannosaurus Rex", "Stegosaurus", "Triceratops")

    Lapis.CollectionExt.AddRange Dinosaurs, MoreDinosaurs
    Debug.Print vbNewLine
    For Each Item In Dinosaurs
        Debug.Print Item
    Next Item

End Sub

' This code example produces the following output:
'
' Brachiosaurus
' Amargasaurus
' Mamenchisaurus
'
'
' Brachiosaurus
' Amargasaurus
' Mamenchisaurus
' Tyrannosaurus Rex
' Stegosaurus
' Triceratops
```

