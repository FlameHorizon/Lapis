# CollectionExt.ToString Method

Returns string based on the given set using value converter

```vb
Public Function ToString(ByVal Source As Collection, ByVal Converter As IToString, Optional ByVal Delimiter As String = ",") As String
```

### Parameters

**Source** `Collection` <br>
Set of items which will be converted.

**Converter** `IToString` <br>
Contains information how single item of set can be converted to string

**Delimiter** `String` <br>
Optional. Defines delimiter which will be used between each item.

### Returns

`String` <br>
Text which will represent all items in the source.

### Errors

`OnArgumentNull` <br>
`Source` or `Converter` is `Nothing`

`OnInvalidOperation` <br>
Converting a value to string fails

## Examples

The following code example demonstrates how to use ToString(Collection, IToString, String) to covert set to string separated by comma.

```
Option Explicit

Public Sub Start()

    Dim Fruits As Collection
    Set Fruits = Lapis.CollectionExt.Make("apple", "passionfruit", "banana", "mango", _
                                         "orange", "blueberry", "grape", "strawberry")
    
    Dim Result As String
    Result = Lapis.CollectionExt.ToString(Fruits, _
                                          Lapis.Factory.GetValueTypeToStringConverter, _
                                          Delimiter:=", ")
    Debug.Print Result

End Sub

' This code produces the following output:

' apple, passionfruit, banana, mango, orange, blueberry, grape, strawberry
```

