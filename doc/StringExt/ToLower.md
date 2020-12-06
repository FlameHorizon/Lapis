# StringExt.ToLower Method

Returns a copy of this string converted to lowercase.

```vb
Public Function ToLower(ByVal Str As String) As String
```

### Parameters

**Str** `String` <br>
A source string which will be turned into lower cased version.

### Returns

`String` <br>
A string in lowercase.

## Examples

The following example converts string to lowercase.

```vb
Option Explicit

Public Sub Start()

    Dim BigLetters As String: BigLetters = "ABCDEFG"
    Debug.Print StringExt.ToLower(BigLetters)
    
End Sub

' The example displays the following output:
'
' abcdefg
```

