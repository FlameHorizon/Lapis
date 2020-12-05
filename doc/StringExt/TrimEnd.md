# StringExt.TrimEnd Method

Removes all the trailing occurrences of a set of characters specified in an array from the current string.

```vb
Public Function TrimEnd(ByVal Str As String, ParamArray TrimChars() As Variant) As String
```

### Parameters

**Str** `String` <br>
A source string which will be trimmed.

**TrimChars** `ParamArray Variant` <br>
An array of Unicode characters to remove, or `vbNullString`.

### Returns

`String` <br>
The string that remains after all occurrences of characters in the `TrimChars` parameter are removed from the end of the current string. If `TrimChars` is `vbNullString` or an empty array, white-space characters are removed instead. If no characters can be trimmed from the current instance, the method returns the current instance unchanged.

## Examples

The following example demonstrates how you can use the TrimEnd(String, ParamArray) method to trim white space or punctuation marks from the end of a string.

```vb
Option Explicit

Public Sub Start()

    Dim Sentence As String
    Sentence = "The dog had a bone, a ball, and other toys."
    
    Dim Words As Variant: Words = VBA.Split(Sentence)
    Dim Word As Variant
    For Each Word In Words
        Debug.Print StringExt.TrimEnd(Word, ",", ".", " ")
    Next Word

End Sub

' The example displays the following output:
'
' The
' dog
' had
' a
' bone
' a
' ball
' and
' other
' toys
```

