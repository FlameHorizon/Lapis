# StringExt.IsEmptyOrWhiteSpace Method

Indicates whether a specified string is empty, or consists only of white-space characters.

```vb
Public Function IsEmptyOrWhiteSpace(ByVal Str As String) As Boolean
```

### Parameters

**Str** `String` <br>
A source string to test.

### Returns

`Boolean` <br>
`True` if the value parameter empty, or if value consists exclusively of white-space characters

## Examples

The following example creates a string Collection, and then passes each element of the Collection to the `IsEmptyOrWhiteSpace` method.

```vb
Option Explicit

Public Sub Start()

    Dim Values As Collection
    Set Values = CollectionExt.Make(vbNullString, _
                                    "", _
                                    "ABCDE", _
                                    String(20, " "), _
                                    vbTab)
    
    Dim Str As Variant
    For Each Str In Values
        Debug.Print StringExt.IsEmptyOrWhiteSpace(Str)
    Next Str
    
End Sub

' The example displays the following output:
'
' True
' True
' False
' True
' True
```

