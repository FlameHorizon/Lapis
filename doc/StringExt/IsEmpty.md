# StringExt.IsEmpty Method

Indicates whether the specified string is null or an empty string (vbNullString).

```vb
Public Function IsEmpty(ByVal Str As String) As Boolean
```

### Parameters

**Str** `String` <br>
A source string to test.

### Returns

`Boolean` <br>
`True` if the value parameter is an empty string (vbNullString); otherwise, `False`

## Examples

The following example examines three strings and determines whether each string has a value, is an empty string.

```vb
Option Explicit

Public Sub Start()

    Dim Str1 As String: Str1 = "abcd"
    Dim Str2 As String: Str2 = ""
    Dim Str3 As String: Str3 = vbNullString
    
    Debug.Print "String Str1 " & Test(Str1)
    Debug.Print "String Str2 " & Test(Str2)
    Debug.Print "String Str3 " & Test(Str3)

End Sub


Private Function Test(ByVal Str As String) As String

    If StringExt.IsEmpty(Str) Then
        Test = "is empty."
    Else
        Test = "(""" & Str & """) is not empty"
    End If

End Function

' The example displays the following output:
'
' String Str1 ("abcd") is not empty
' String Str2 is empty.
' String Str3 is empty.


```

