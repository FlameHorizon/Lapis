Attribute VB_Name = "StringH"
Option Explicit
'@Folder("Helper")

Private Const ModuleName As String = "StringHelper"


' If str parameter is empty, then function returns -1,
' if no element was found in str, then function returns 0 else,
' returns the index of the first occurrence of any character in value.
' Note: Function does not check if array is initialized properly.
Public Function IndexOfAny(ByVal Str As String, ByRef AnyOf() As String) As Long
    
    Const MethodName = "IndexOfAny"
    
    If ArrayH.IsInitialized(ArrayH.ToVariantArray(AnyOf)) = False Then
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
        Case VbCompareMethod.vbBinaryCompare, VbCompareMethod.vbTextCompare
            IndexOf = InternalIndexOf(Str, Value, StartIndex, Count, ComparisonMethod)
        
        Case Else
            Exception.ArgumentOutOfRangeException "ComparisonMethod", "Not supported string comparison. " & MethodName & "." & MethodName
            
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


' Returns a new string in which all the characters in the current instance,
' beginning at a specified position and continuing through the last position, have been deleted.
Public Function Remove(ByVal Str As String, ByVal StartIndex As Long) As String
    
    Const MethodName = "Remove"
    If StartIndex < 0 Then
        Exception.ArgumentOutOfRangeException "StartIndex", ModuleName & "." & MethodName
    End If
    
    If StartIndex >= VBA.Len(Str) Then
        Exception.ArgumentOutOfRangeException "StartIndex", ModuleName & "." & MethodName
    End If
    
    ' Vba.Mid method does exactly the same thing as
    ' String.Substring method in .NET.
    Remove = VBA.Mid$(Str, 1, StartIndex)
    
End Function


' Returns a new string in which a specified number of characters in the current instance
' beginning at a specified position have been deleted.
Public Function RemoveRange(ByVal Str As String, ByVal StartIndex As Long, ByVal Count As Long) As String

    Const MethodName = "RemoveRange"
    If StartIndex < 0 Then
        Exception.ArgumentOutOfRangeException "StartIndex", ModuleName & "." & MethodName
    End If
    
    If Count < 0 Then
        Exception.ArgumentOutOfRangeException "Count", ModuleName & "." & MethodName
    End If
    
    If Count > Len(Str) - StartIndex Then
        Exception.ArgumentOutOfRangeException "Count", ModuleName & "." & MethodName
    End If
    
    If Len(Str) - Count = 0 Then
        RemoveRange = vbNullString
        Exit Function
    End If
    
    Dim Lhs As String
    Lhs = VBA.Mid$(Str, 1, StartIndex)
    
    Dim Rhs As String
    Rhs = VBA.Mid$(Str, (StartIndex + 1) + Count, Len(Str))
    
    RemoveRange = Lhs + Rhs

End Function


' Concatenates the members of a constructed Variant ParamArray of type String.
Public Function Concat(ParamArray Items() As Variant) As String
    
    Dim Output As String
    Dim Item As Variant
    For Each Item In Items
        Output = Output + Item
    Next Item
    
    Concat = Output
    
End Function


Public Function StartsWith(ByVal Str As String, ByVal Value As String, ByVal StringComparison As VbCompareMethod) As Boolean

    Const MethodName = "StartsWith"
    
    If StringComparison = VbCompareMethod.vbDatabaseCompare Then
        Exception.ArgumentOutOfRangeException "StringComparison", ModuleName & "." & MethodName
    End If
    
    If Str = vbNullString Then
        Exception.ArgumentException "Str", ModuleName & "." & MethodName
    End If
    
    If Len(Value) = 0 Then
        StartsWith = True
        Exit Function
    End If
    
    If Str = Value Then
        StartsWith = True
        Exit Function
    End If
    
    If StringComparison = VbCompareMethod.vbBinaryCompare Then
        StartsWith = (Str Like Value & "*")
    Else
        StartsWith = (VBA.LCase$(Str) Like VBA.LCase$(Value) & "*")
    End If

End Function


Public Function EndsWith(ByVal Str As String, ByVal Value As String, ByVal StringComparison As VbCompareMethod) As Boolean
    
    Const MethodName = "EndsWith"
    
    If StringComparison = VbCompareMethod.vbDatabaseCompare Then
        Exception.ArgumentOutOfRangeException "StringComparison", ModuleName & "." & MethodName
    End If
    
    If Str = vbNullString Then
        Exception.ArgumentException "Str", ModuleName & "." & MethodName
    End If
    
    If Len(Value) = 0 Then
        EndsWith = True
        Exit Function
    End If
    
    If Str = Value Then
        EndsWith = True
        Exit Function
    End If
    
    If StringComparison = VbCompareMethod.vbBinaryCompare Then
        EndsWith = (Str Like "*" & Value)
    Else
        EndsWith = (VBA.LCase$(Str) Like "*" & VBA.LCase$(Value))
    End If
    
End Function


Public Function Insert(ByVal Str As String, _
                       ByVal StartIndex As Long, _
                       ByVal Value As String) As String
    
    Const MethodName = "Insert"

    If Value = vbNullString Then
        Exception.ArgumentException "Value", ModuleName & "." & MethodName
    End If
    
    If StartIndex < 0 Or StartIndex > Len(Str) Then
        Exception.ArgumentOutOfRangeException "StartIndex", ModuleName & "." & MethodName
    End If
    
    Dim Lhs As String
    Lhs = VBA.Mid$(Str, 1, StartIndex)
    
    Dim Rhs As String
    Rhs = VBA.Mid$(Str, StartIndex + 1, Len(Str))
    
    Insert = Lhs + Value + Rhs
    
End Function


Public Function LastIndexOf(ByVal Str As String, _
                            ByVal Value As String, _
                            ByVal StartIndex As Long, _
                            ByVal Count As Long, _
                            ByVal ComparisonMethod As VbCompareMethod) As Long
    
    Const MethodName = "LastIndexOf"
    
    LastIndexOf = -1
    
    If Str = vbNullString Then
        Exception.ArgumentException "Str", ModuleName & "." & MethodName
    End If

    If Value = vbNullString Then
        Exception.ArgumentException "Value", ModuleName & "." & MethodName
    End If
    
    If Len(Str) = 0 And (StartIndex = -1 Or StartIndex - 0) Then
        LastIndexOf = IIf(Len(Value) = 0, 0, -1)
    End If
    
    If StartIndex < 0 Or StartIndex > Len(Str) Then
        Exception.ArgumentOutOfRangeException "StartIndex", ModuleName & "." & MethodName
    End If
    
    If StartIndex = Len(Str) Then
        StartIndex = StartIndex - 1
        If Count > 0 Then
            Count = Count - 1
        End If
        
        If Len(Value) = 0 And Count >= 0 And StartIndex - Count + 1 >= 0 Then
            LastIndexOf = StartIndex
            Exit Function
        End If
    End If
    
    If Count < 0 Or StartIndex - Count + 1 < 0 Then
        Exception.ArgumentOutOfRangeException "Count", ModuleName & "." & MethodName
    End If
    
    Select Case ComparisonMethod
        Case CompareMethod.BinaryCompare, CompareMethod.TextCompare
            Dim MinIndex As Long
            MinIndex = StartIndex - Count + 1
            
            Dim i As Long
            For i = StartIndex To MinIndex Step -1
                If InStrRev(Str, Value, i + 1, ComparisonMethod) <> 0 Then
                    LastIndexOf = InStrRev(Str, Value, i + 1, ComparisonMethod) - 1
                    Exit For
                End If
            Next i
        Case Else
            Exception.ArgumentOutOfRangeException "ComparisonMethod", "Not supported string comparison. " & MethodName & "." & MethodName
            
    End Select
    
    
End Function


Public Function LastIndexOfAny(ByVal Str As String, ByRef AnyOf() As String) As Long
    
    Const MethodName = "LastIndexOfAny"
    
    If ArrayH.IsInitialized(ArrayH.ToVariantArray(AnyOf)) = False Then
        Exception.ArgumentException "AnyOf", ModuleName & "." & MethodName
    End If

    LastIndexOfAny = -1

    If Str = vbNullString Then
        Exit Function
    End If
    
    Dim i As Long
    For i = 0 To UBound(AnyOf)
        LastIndexOfAny = LastIndexOf(Str, AnyOf(i), Len(Str) - 1, Len(Str), vbTextCompare)

        If LastIndexOfAny <> -1 Then
            Exit Function
        End If
    Next i

End Function



