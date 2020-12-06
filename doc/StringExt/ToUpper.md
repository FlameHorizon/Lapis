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

The following example calls the ToUpper method to convert a series of one-character strings that contain each character in the Basic Latin, Latin-1 Supplement, and Latin Extended-A character sets. It then displays each string whose uppercase character is different from its lowercase character.

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

