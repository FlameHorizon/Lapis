# StringExt.Trim Method

Removes all leading and trailing occurrences of a set of characters specified in an array from the current string.

```vb
Public Function Trim(ByVal Str As String, ParamArray TrimChars() As Variant) As String
```

### Parameters

**Str** `String` <br>
A source string which will be trimmed.

**TrimChars** `ParamArray Variant` <br>
An array of Unicode characters to remove, or `vbNullString`.

### Returns

`String` <br>
The string that remains after all occurrences of the characters in the `trimChars` parameter are removed from the start and end of the current string. If `trimChars` is an empty array, white-space characters are removed instead. If no characters can be trimmed from the current instance, the method returns the current instance unchanged.

## Examples

The following example uses the Trim(String, ParamArray) method to remove space, asterisk (*), and apostrophe (') characters from a string.

```vb
Option Explicit

Public Sub Start()

    Dim Banner As String: Banner = "*** Banner - Buffalo Bull EFB ***"
    Dim Result As String: Result = StringExt.Trim(Banner, "*", " ", "'")
    
    Debug.Print "Trimmer " & Banner & " to '" & Result & "'"

End Sub

' The example displays the following output:
'
' Trimmer *** Banner - Buffalo Bull EFB *** to 'Banner - Buffalo Bull EFB'
```

