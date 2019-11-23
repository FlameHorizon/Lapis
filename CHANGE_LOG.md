# Change log

## Changes in 0.6.0

* Exposed `ExUnit.TestPass` and `ExUnit.TestFail` methods to better express some test's intentions.
* Added `StringH.Remove` which returns new string in which a specified number of characters from the current string are deleted.

```vb
Public Sub Start()
    Debug.Print StringH.Remove("foo-bar", 3)
End Sub

' This example produces the following results:
'
' foo
```

* Added `StringH.RemoveRange` which returns a new string in which a specified number of characters in the current instance beginning at a specified position have been deleted.

```vb
Option Explicit
Public Sub Start()
    Debug.Print StringH.RemoveRange("foo-bar", 3, 1)
End Sub

' This example produces the following results:
'
' foobar
```

* Added `StringH.Concat` which concatenates elements of a specified string array.

```vb
Option Explicit
Public Sub Start()
    Debug.Print StringH.Concat("A", "bbb", "c", "12#")
End Sub

' This example produces the following results:
'
' Abbbc12#
```

* Added `StringH.StartsWith` which determines whether the beginning of this string instance matches a specified string.

```vb
Option Explicit
Public Sub Start()

    Test("John", "J", VbCompareMethod.vbTextCompare)
    Test("John", "J", VbCompareMethod.vbBinaryCompare)

End Sub


Private Sub Test(ByVal x as String, ByVal y as String, ByVal CompareMethod as vbCompareMethod)

    If CompareMethod = vbTextCompare Then
        Debug.Print "vbCompareMethod.vbTextCompare"

    ElseIf CompareMethod = vbBinaryCompare Then
        Debug.Print "vbCompareMethod.vbBinaryCompare"
    End if

    If StringH.StartsWith(x, y, CompareMethod) Then
        Debug.Print x & " ends with " & y
    Else
        Debug.Print x & " does not ends with " & y
    End if

    Debug.Print vbNewLine

End Sub

' This example produces the following results:
'
' vbCompareMethod.vbTextCompare
' John starts with J
'
' vbCompareMethod.vbBinaryCompare
' John does not starts with J
```

* Added `StringH.EndsWith` which determines whether the end of this string instance matches the specified string when compared using the specified comparison option.

```vb
Option Explicit
Public Sub Start()

    Test("John", "N", VbCompareMethod.vbTextCompare)
    Test("John", "N", VbCompareMethod.vbBinaryCompare)

End Sub


Private Sub Test(ByVal x as String, ByVal y as String, ByVal CompareMethod as vbCompareMethod)

    If CompareMethod = vbTextCompare Then
        Debug.Print "vbCompareMethod.vbTextCompare"

    ElseIf CompareMethod = vbBinaryCompare Then
        Debug.Print "vbCompareMethod.vbBinaryCompare"
    End if

    If StringH.EndsWith(x, y, CompareMethod) Then
        Debug.Print x & " ends with " & y
    Else
        Debug.Print x & " does not ends with " & y
    End if
    Debug.Print vbNewLine

End Sub

' This example produces the following results:
'
' vbCompareMethod.vbTextCompare
' John ends with N
'
' vbCompareMethod.vbBinaryCompare
' John does not ends with N
```

* Added `StringH.Insert` which returns a new string in which a specified string is inserted at a specified index position in the instance.

```vb
Option Explicit
Public Sub Start()

    Dim Original as String: Original = "aaabbb"
    Dim Modified as String
    Modified = StringH.Insert(Original, 3, " ")
    Debug.Print "This is original string: " & Original
    Debug.Print "This is modified string: " & Modified

End Sub

' This example produces the following results:
'
' This is original string: aaabbb
' This is modified string: aaa bbb
```

* Added `StringH.LastIndexOf` which reports the zero-based index position of the last occurrence of a specified string within this instance.

```vb
Option Explicit
Public Sub Start()

    Debug.Print "Comparison: vbTextCompare. Location: " & StringH.LastIndexOf(Text, "b", 0, Len(Text), vbCompareMethod.vbTextCompare
    Debug.Print "Comparison: vbBinaryCompare. Location:  " & StringH.LastIndexOf(Text, "B", 0, Len(Text), vbCompareMethod.vbBinaryCompare)

End Sub

' This example produces the following results:
'
' Comparison: vbTextCompare. Location: 3
' Comparison: vbBinaryCompare. Location: -1
```

* Added `StringH.LastIndexOfAny` which reports the zero-based index position of the last occurrence in this instance of one or more characters.

```vb
Option Explicit
Public Sub Start()

    Debug.Print "The last character occurrence location: " & StringH.LastIndexOfAny("abcd", StringArray("a", "d"))
    Debug.Print "The last character occurrence location: " & StringH.LastIndexOfAny("abcd", StringArray("d", "a"))
    Debug.Print "The last character occurrence location: " & StringH.LastIndexOfAny("abcd", StringArray("b", "c"))

End Sub

' This example produces the following results:
' The last character occurrence location: 0
' The last character occurrence location: 3
' The last character occurrence location: 1
```
