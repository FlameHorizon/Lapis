# StringExt.Replace Method

Returns a new string in which all occurrences of a specified string in the current instance are replaced with another specified string, using the provided comparison type.

```vb
Public Function Replace(ByVal Str As String, ByVal OldValue As String, ByVal NewValue As String, ByVal ComparisonType As CompareMethod) As String
```

### Parameters

**Str** `String` <br>
A source string.

**OldValue** `String` <br>
The string to be replaced.

**NewValue** `String` <br>
The string to replace all occurrences of `OldValue`.

**ComparisonType** `CompareMethod` <br>
One of the enumeration values that determines how `OldValue` is searched within this instance.

### Returns

`String` <br>
A string that is equivalent to the current string except that all instances of `OldValue` are replaced with `newValue`. If `OldValue` is not found in the current instance, the method returns the current instance unchanged.

### Errors

`OnArgumentError` <br>
`OldValue` is the empty string (vbNullString).

## Examples

Because this method returns the modified string, you can chain together successive calls to the `Replace` method to perform multiple replacements on the original string. Method calls are executed from left to right. The following example provides an illustration.

```vb
Option Explicit

Public Sub Start()

    Dim Str As String: Str = "aaa"
    Debug.Print "The initial string: " & Str
    
    Str = StringExt.Replace(Str, "a", "b", TextCompare)
    Str = StringExt.Replace(Str, "b", "c", TextCompare)
    Str = StringExt.Replace(Str, "c", "d", TextCompare)
    Debug.Print "The final string: " & Str
    
End Sub

' The example displays the following output:
'
' The initial string: aaa
' The final string: ddd

```

