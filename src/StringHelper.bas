Attribute VB_Name = "StringHelper"
' VBA GOLEM Built-In Module. All modifications need to be processed via GIT Pull Request.
' Module providing Useful functions for handling Strings

Option Explicit


Private Const ModuleName As String = "StringHelper"


' If str parameter is empty, then function returns -1,
' if no element was found in str, then function returns 0 else,
' returns the index of the first occurrence of any character in value.
' Note: Function does not check if array is initialized properly.
Public Function IndexOfAny(ByVal Str As String, ByRef AnyOf() As String) As Long
    
    Const MethodName = "IndexOfAny"
    
    If ArrayHelper.IsInitialized(ArrayHelper.ToVariantArray(AnyOf)) = False Then
        Exception.ArgumentException "AnyOf", ModuleName & "." & MethodName
    End If

    IndexOfAny = -1

    If Str = vbNullString Then
        Exit Function
    End If

    Dim i As Long
    For i = 0 To UBound(AnyOf)
        IndexOfAny = IndexOf(Str, AnyOf(i), 0, Len(Str), vbTextCompare)

        If IndexOfAny <> 0 Then
            Exit Function
        End If
    Next i

    IndexOfAny = 0

End Function


' Returns the zero-based index of the first occurrence of the specified string in the current string.
' Throws ArgumentException when Str is empty string.
' Throws ArgumentOutOfRangeException when StartIndex points outside of Value bound or
' Count exceedes availablue scope of Value.
' Throws ArgumentException when CompareMethod value is not valid.
' Params:
' Str - String to look in.
' Value - String to look for in Str.
' StartIndex - Index where search will start at.
' Count - Number of characters after StartIndex which will be checked.
' ComparisonMethod - defines rule for the search.
Public Function IndexOf(ByVal Str As String, _
                        ByVal Value As String, _
                        ByVal StartIndex As Long, _
                        ByVal Count As Long, _
                        ByVal ComparisonMethod As VbCompareMethod) As Long
    
    Const MethodName = "IndexOf"
    
    If Str = vbNullString Then
        Exception.ArgumentException "Str", _
                                    "Str value can't be empty string. " _
                                  & ModuleName & "." & MethodName
    End If
    
    If StartIndex < 0 Or StartIndex > Len(Str) Then
        Exception.ArgumentOutOfRangeException "StartIndex", ModuleName & "." & MethodName
    End If
    
    If Count < 0 Or StartIndex > Len(Str) - Count Then
        Exception.ArgumentOutOfRangeException "Count", ModuleName & "." & MethodName
    End If
    
    Select Case ComparisonMethod
        Case VbCompareMethod.vbBinaryCompare, VbCompareMethod.vbDatabaseCompare, VbCompareMethod.vbTextCompare
            IndexOf = InternalIndexOf(Str, Value, StartIndex, Count, ComparisonMethod)
            
        Case Else
            Exception.ArgumentException "ComparisonMethod", "Not supported string comparison. " & ComparisonMethod & "." & MethodName
            
    End Select
    
End Function


Private Function InternalIndexOf(ByVal Str As String, _
                                 ByVal Value As String, _
                                 ByVal StartIndex As Long, _
                                 ByVal Count As Long, _
                                 ByVal ComparisonMethod As VbCompareMethod) As Long
    
    InternalIndexOf = InStr(StartIndex + 1, Str, Value, ComparisonMethod) - 1
    
    If InternalIndexOf > Count - 1 Then
        InternalIndexOf = -1
    End If
    
End Function


' Returns a value indicating whether a specified substring occurs within this string.
' Params:
' Str - String to look in.
' Value - String to look for in Str.
Public Function Contains(ByVal Str As String, ByVal Value As String) As Boolean
    Contains = IndexOf(Str, Value, 0, Len(Str), vbTextCompare) >= 0
End Function



