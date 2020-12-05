# StringExt.TrimStart Method

Removes all the leading occurrences of a set of characters specified in an array from the current string.

```vb
Public Function TrimStart(ByVal Str As String, ParamArray TrimChars() As Variant) As String
```

### Parameters

**Str** `String` <br>
A source string which will be trimmed.

**TrimChars** `ParamArray Variant` <br>
An array of Unicode characters to remove, or `vbNullString`.

### Returns

`String` <br>
The string that remains after all occurrences of characters in the `TrimChars` parameter are removed from the start of the current string. If `TrimChars` is `vbNullString` or an empty array, white-space characters are removed instead. If no characters can be trimmed from the current instance, the method returns the current instance unchanged.

## Examples

The following example demonstrates the basic functionality of the `TrimStart` method:

```vb 
Option Explicit

Public Sub Start()

    Dim LineWithLeadingSpaces As String: LineWithLeadingSpaces = "    Noice"
    Dim LineWithLeadingSymbols As String: LineWithLeadingSymbols = "$$$$Noice"
    Dim LineWithLeadingUnderscores As String: LineWithLeadingUnderscores = "____Noice"
    Dim LineWithLeadingLetter As String: LineWithLeadingLetter = "xxxxNoice"
    
    ' Add lines to array to work with them all of the examples
    Dim Lines As Collection
    Set Lines = CollectionExt.Make(LineWithLeadingSpaces, _
                                    LineWithLeadingSymbols, _
                                    LineWithLeadingUnderscores, _
                                    LineWithLeadingLetter)
                                    
    Debug.Print "These are initial values we are going to work with."
    Dim Line As Variant
    For Each Line In Lines
        Debug.Print Line
    Next Line
    
    ' A basic demonstration of TrimStart in action
    Dim LineAfterTrimStart As String
    LineAfterTrimStart = StringExt.TrimStart(LineWithLeadingSpaces, " ")
    Debug.Print vbNewLine & "Line, after removing spaces: " _
                          & LineAfterTrimStart
    
    ' Since TrimStart accepts a character array of leading items to be removed as an argument,
    ' it's possible to do things like trim multiple pieces of data that each have different
    ' leading characters,
    For Each Line In Lines
        Debug.Print "Line " & Line & " was trimmed to " _
                     & StringExt.TrimStart(Line, " ", "$", "_", "x")
    Next Line
    
    ' Or handle pieces of data that have multiple kinds of leading characters
    Dim LineToBeTrimmed As String: LineToBeTrimmed = "__###__ John Smith"
    Debug.Print "Line " & LineToBeTrimmed & " was trimmed to " _
                & StringExt.TrimStart(LineToBeTrimmed, "_", "#", " ")
    
End Sub

' This code produces the following output:
' These are initial values we are going to work with.
'     Noice
' $$$$Noice
' ____Noice
' xxxxNoice
'
' Line, after removing spaces: Noice
' Line     Noice was trimmed to Noice
' Line $$$$Noice was trimmed to Noice
' Line ____Noice was trimmed to Noice
' Line xxxxNoice was trimmed to Noice
' Line __###__ John Smith was trimmed to John Smith
```

