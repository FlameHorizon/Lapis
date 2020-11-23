# CollectionExt.ToArray Method

Creates an array from a collection.

```vb
Public Function ToArray(ByRef Items As Collection) As Variant()
```

### Parameters

**Items** `Collection` <br>
A collection to create an array from.

### Returns

`Variant()` <br>
An array that contains the elements from the input sequence.

### Errors

`OnArgumentNull` <br>
`Items` is `Nothing`

## Examples

The following code example demonstrates how to use ToArray to force immediate query evaluation and return an array of results.

```vb
' Package class module
Option Explicit

Public Company As String
Public Weight As Double
```

```vb
Option Explicit

Public Sub Start()

    Dim Packages As New Collection
    With Packages
        .Add MakePackage("Coho Vineyard", 25.2)
        .Add MakePackage("Lucerne Publishing", 15)
        .Add MakePackage("Wingtip Toys", 6)
        .Add MakePackage("Adventure Works", 33.9)
    End With

    Dim Converter As Lapis.PropertyToStringConverter
    Set Converter = Lapis.Factory.PropertyToStringConverter
    Converter.PropertyName = "Company"

    Dim Companies() As Variant
    Companies = Lapis.CollectionExt.ToArray(Lapis.CollectionExt.Convert(Packages, Converter))

    Dim Item As Variant
    For Each Item In Companies
        Debug.Print Item
    Next Item

End Sub


' This code produces the following output:
'
' Coho Vineyard
' Lucerne Publishing
' Wingtip Toys
' Adventure Works


Private Function MakePackage(ByVal Company As String, ByVal Weight As Double) As Package

    Dim Output As New Package
    Output.Company = Company
    Output.Weight = Weight
    Set MakePackage = Output

End Function
```

