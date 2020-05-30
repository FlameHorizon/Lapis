# CollectionExt.Sort Method

Sorts given collection using merge sort according to defined comparer.

```vb
Public Function Sort(ByVal Items As Collection, ByVal Comparer As Lapis.IComparer) As Collection
```

### Parameters

**Items** `Collection` <br>
Items which will be sorted.

**Comparer** `Lapis.IComparer` <br>
Comparer which defined how items will be sorted.

### Returns

`Collection` <br>
Sorted sequence of items.

## Examples

The following code example demonstrates how to use Sort method to sort the elements of a sequence using custom comparers.

```vb
' Car class module
Option Explicit

Public Manufacturer As String
Public Year As Long
```

```vb
' CarManufacturerComparer class module
Option Explicit

Implements IComparer


Public Function Compare(ByRef x As Variant, ByRef y As Variant) As Long
    
    If System.IsNothing(x) And System.IsNothing(y) Then
        Compare = 0
    
    Else
        Compare = StringComparers.OrdinalIgnoreCase.Compare(x.Manufacturer, y.Manufacturer)
        
    End If
    
End Function


Private Function IComparer_Compare(ByRef x As Variant, ByRef y As Variant) As Long
    IComparer_Compare = Me.Compare(x, y)
End Function
```

```vb
' CarYearComparer class module
Option Explicit

Implements IComparer


Public Function Compare(ByRef x As Variant, ByRef y As Variant) As Long

    If System.IsNothing(x) And System.IsNothing(y) Then
        Compare = 0
       
    Else
        Compare = Lapis.Factory.GetLongComparer.Compare(x.Year, y.Year)
        
    End If

End Function


Private Function IComparer_Compare(ByRef x As Variant, ByRef y As Variant) As Long
    IComparer_Compare = Me.Compare(x, y)
End Function
```

```vb
' Start module
Option Explicit


Public Sub Start()

    Dim Cars As Collection
    Set Cars = CollectionExt.Make(MakeCar("Toyota", 1990), _
                                  MakeCar("Porsche", 1995), _
                                  MakeCar("Fiat", 1989))
    
    Debug.Print "Cars sorted by the Year of production"
    PrintCars CollectionExt.Sort(Cars, New CarYearComparer)
    Debug.Print vbNewLine
    
    Debug.Print "Cars sorted by the Manufacturer name"
    PrintCars CollectionExt.Sort(Cars, New CarManufacturerComparer)
    
End Sub


Private Function MakeCar(ByVal Manufacturer As String, ByVal Year As Long) As Car
    
    Dim Output As New Car
    Output.Manufacturer = Manufacturer
    Output.Year = Year
    Set MakeCar = Output
    
End Function


Private Sub PrintCars(ByVal Cars As Collection)

    Dim Car As Car
    For Each Car In Cars
        Debug.Print "Car manufacturer: " & Car.Manufacturer & ", Production year: " & Car.Year
    Next Car

End Sub

' This code produces the following output:

' Cars sorted by the Year of production
' Car manufacturer: Fiat, Production year: 1989
' Car manufacturer: Toyota, Production year: 1990
' Car manufacturer: Porsche, Production year: 1995


' Cars sorted by the Manufacturer name
' Car manufacturer: Fiat, Production year: 1989
' Car manufacturer: Porsche, Production year: 1995
' Car manufacturer: Toyota, Production year: 1990
```

