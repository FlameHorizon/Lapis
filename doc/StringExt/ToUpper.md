# StringExt.ToUpper Method

The uppercase equivalent of the current string.

```vb
Public Function ToUpper(ByVal Str As String) As String
```

### Parameters

**Str** `String` <br>
A source string which will be turned into upper cased version.

### Returns

`String` <br>
The uppercase equivalent of the current string.

## Examples

The following example uses the Trim(String, ParamArray) method to remove space, asterisk (*), and apostrophe (') characters from a string.

```vb
Option Explicit

Public Sub Start()

    Dim ToUpperLetters As String: ToUpperLetters = "abcDefg"
    Debug.Print StringExt.ToUpper(ToUpperLetters)

End Sub

' The example displays the following output:
'
' ABCDEFG
```

