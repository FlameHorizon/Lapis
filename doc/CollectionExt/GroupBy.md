# CollectionExt.GroupBy Method

Returns a dictionary with grouped values where key is a unique value and item is a collection of items which matches key.

```vb
Public Function GroupBy(ByVal Items As Collection, ByVal PropertyName As String) As Scripting.Dictionary
```

### Parameters

**Items** `Collection` <br>
A set which will be grouped.

**PropertyName** `String` <br>
A property of the object which will by used as key.

### Returns

`Scripting.Dictionary` <br>
Grouped dictionary base on value.

### Errors

`OnArgumentNull` <br>
`Items` is `Nothing

`OnArgumentError` <br>
`PropertyName` is `vbNullString`

## Examples

The following code example demonstrates how to use GroupBy(Collection, String) to group the project elements of sequence.

```vb
Option Explicit

' Reservation class
Public Id As Long
Public BookedOn As Date
```

```vb
Option Explicit

'Main module
Public Sub Start()

    Dim Reservations As Collection
    Set Reservations = Lapis.CollectionExt.Make(MakeReservation(1, #1/1/2020#), _
                                                MakeReservation(2, #1/1/2020#), _
                                                MakeReservation(3, #1/2/2020#), _
                                                MakeReservation(4, #1/3/2020#))
    
    Dim GroupedByDate As Scripting.Dictionary
    Set GroupedByDate = Lapis.CollectionExt.GroupBy(Reservations, "BookedOn")
    
    Dim DateKey As Variant
    For Each DateKey In GroupedByDate.Keys
        Debug.Print "DateKey: " & DateKey & ", elements: " & GroupedByDate.Item(DateKey).Count
    Next DateKey

End Sub


Private Function MakeReservation(ByVal Id As Long, ByVal BookedOn As Date) As Reservation
    
    Dim Output As New Reservation
    Output.Id = Id
    Output.BookedOn = BookedOn
    Set MakeReservation = Output
    
End Function

```

