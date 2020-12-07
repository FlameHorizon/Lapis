Attribute VB_Name = "StringExt"
Option Explicit
'@Folder("Helper")

Private Const ModuleName As String = "StringExt"


' Specifies which portions of the string should be trimmed in a trimming operation.
Private Enum TrimType

    ' Trim from the beginning of the string
    Head = 1
    
    ' Trim from the end of the string.
    Tail
    
    ' Trim from both the beginning and the end of the string.
    Both
    
End Enum


Public Function IndexOfAny(ByVal Str As String, ByRef AnyOf() As String) As Long
    
    Const MethodName = "IndexOfAny"
    
    If ArrayExt.IsInitialized(ArrayExt.ToVariantArray(AnyOf)) = False Then
        Errors.OnArgumentError "AnyOf", ModuleName & "." & MethodName
    End If

    IndexOfAny = -1

    If Str = vbNullString Then
        Exit Function
    End If

    Dim i As Long
    For i = 0 To UBound(AnyOf)
        IndexOfAny = IndexOf(Str, AnyOf(i), 0, vbTextCompare)

        If IndexOfAny <> 0 Then
            Exit Function
        End If
    Next i

    IndexOfAny = 0

End Function


Public Function IndexOf(ByVal Str As String, _
                        ByVal Value As String, _
                        Optional ByVal StartIndex As Long = System.LongMinValue, _
                        Optional ByVal ComparisonMethod As VbCompareMethod = VbCompareMethod.vbTextCompare) As Long
    
    Const MethodName = "IndexOf"
    
    Dim Start As Long
    Start = IIf(StartIndex = System.LongMinValue, 0, StartIndex)
    
    If Str = vbNullString Then
        Errors.OnArgumentError "Str", _
                               "Str value can't be empty string. " _
                               & ModuleName & "." & MethodName
    End If
    
    If Start < 0 Or Start > Len(Str) Then
        Errors.OnArgumentOutOfRange "Start", ModuleName & "." & MethodName
    End If
    
    Select Case ComparisonMethod
        Case VbCompareMethod.vbBinaryCompare, VbCompareMethod.vbTextCompare
            IndexOf = VBA.InStr(Start + 1, Str, Value, ComparisonMethod) - 1
        
        Case Else
            Errors.OnArgumentOutOfRange "ComparisonMethod", "Not supported string comparison. " & MethodName & "." & MethodName
            
    End Select
    
End Function


Public Function Contains(ByVal Str As String, ByVal Value As String) As Boolean
    Contains = IndexOf(Str, Value, 0, VbCompareMethod.vbTextCompare) >= 0
End Function


Public Function Remove(ByVal Str As String, ByVal StartIndex As Long) As String
    
    Const MethodName = "Remove"
    If StartIndex < 0 Then
        Errors.OnArgumentOutOfRange "StartIndex", ModuleName & "." & MethodName
    End If
    
    If StartIndex >= VBA.Len(Str) Then
        Errors.OnArgumentOutOfRange "StartIndex", ModuleName & "." & MethodName
    End If
    
    ' Vba.Mid method does exactly the same thing as
    ' String.Substring method in .NET.
    Remove = VBA.Mid$(Str, 1, StartIndex)
    
End Function


Public Function RemoveRange(ByVal Str As String, ByVal StartIndex As Long, ByVal Count As Long) As String

    Const MethodName = "RemoveRange"
    If StartIndex < 0 Then
        Errors.OnArgumentOutOfRange "StartIndex", ModuleName & "." & MethodName
    End If
    
    If Count < 0 Then
        Errors.OnArgumentOutOfRange "Count", ModuleName & "." & MethodName
    End If
    
    If Count > Len(Str) - StartIndex Then
        Errors.OnArgumentOutOfRange "Count", ModuleName & "." & MethodName
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


Public Function Concat(ParamArray Items() As Variant) As String
    
    Dim Output As String
    Dim Item As Variant
    For Each Item In Items
        Output = Output + Item
    Next Item
    
    Concat = Output
    
End Function


Public Function StartsWith(ByVal Str As String, _
                           ByVal Value As String, _
                           ByVal StringComparison As VbCompareMethod) As Boolean

    Const MethodName = "StartsWith"
    
    If StringComparison = VbCompareMethod.vbDatabaseCompare Then
        Errors.OnArgumentOutOfRange "StringComparison", ModuleName & "." & MethodName
    End If
    
    If Str = vbNullString Then
        Errors.OnArgumentError "Str", _
                                "Value of Str can't be an empty string. " & _
                                ModuleName & "." & MethodName
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
        StartsWith = (VBA.LCase$(Str) Like "[" & VBA.LCase$(Value) & "]" & "*")
    End If

End Function


Public Function EndsWith(ByVal Str As String, _
                         ByVal Value As String, _
                         ByVal StringComparison As VbCompareMethod) As Boolean
    
    Const MethodName = "EndsWith"
    
    If StringComparison = VbCompareMethod.vbDatabaseCompare Then
        Errors.OnArgumentOutOfRange "StringComparison", ModuleName & "." & MethodName
    End If
    
    If Str = vbNullString Then
        Errors.OnArgumentError "Str", ModuleName & "." & MethodName
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
        Errors.OnArgumentError "Value", ModuleName & "." & MethodName
    End If
    
    If StartIndex < 0 Or StartIndex > Len(Str) Then
        Errors.OnArgumentOutOfRange "StartIndex", ModuleName & "." & MethodName
    End If
    
    Dim Lhs As String
    Lhs = VBA.Mid$(Str, 1, StartIndex)
    
    Dim Rhs As String
    Rhs = VBA.Mid$(Str, StartIndex + 1, Len(Str))
    
    Insert = Lhs + Value + Rhs
    
End Function


Public Function LastIndexOf(ByVal Str As String, _
                            ByVal Value As String, _
                            Optional ByVal StartIndex As Long = System.LongMinValue, _
                            Optional ByVal Count As Long = System.LongMinValue, _
                            Optional ByVal ComparisonMethod As VbCompareMethod = VbCompareMethod.vbTextCompare) As Long
    
    Const MethodName = "LastIndexOf"
    
    Dim Start As Long
    Start = IIf(StartIndex = System.LongMinValue, VBA.Len(Str) - 1, StartIndex)

    Dim Cnt As Long
    Cnt = IIf(Count = System.LongMinValue, VBA.Len(Str) - Start, Count)
    
    LastIndexOf = -1
    
    If Str = vbNullString Then
        Errors.OnArgumentError "Str", ModuleName & "." & MethodName
    End If

    If Value = vbNullString Then
        Errors.OnArgumentError "Value", ModuleName & "." & MethodName
    End If
    
    If Len(Str) = 0 And (Start = -1 Or Start = 0) Then
        LastIndexOf = IIf(Len(Value) = 0, 0, -1)
    End If
    
    If Start < 0 Or Start > Len(Str) Then
        Errors.OnArgumentOutOfRange "Start", ModuleName & "." & MethodName
    End If
    
    If Start = Len(Str) Then
        Start = Start - 1
        If Cnt > 0 Then
            Cnt = Cnt - 1
        End If
        
        If Len(Value) = 0 And Cnt >= 0 And Start - Cnt + 1 >= 0 Then
            LastIndexOf = Start
            Exit Function
        End If
    End If
    
    If Cnt < 0 Or Start - Cnt + 1 < 0 Then
        Errors.OnArgumentOutOfRange "Cnt", ModuleName & "." & MethodName
    End If
    
    Select Case ComparisonMethod
        Case CompareMethod.BinaryCompare, CompareMethod.TextCompare
            Dim MinIndex As Long
            MinIndex = Start - Cnt + 1
            
            Dim i As Long
            For i = Start To MinIndex Step -1
                If InStrRev(Str, Value, i + 1, ComparisonMethod) <> 0 Then
                    LastIndexOf = InStrRev(Str, Value, i + 1, ComparisonMethod) - 1
                    Exit For
                End If
            Next i
            
        Case Else
            Errors.OnArgumentOutOfRange "ComparisonMethod", "Not supported string comparison. " & MethodName & "." & MethodName
            
    End Select
    
End Function


Public Function LastIndexOfAny(ByVal Str As String, ByRef AnyOf() As String) As Long
    
    Const MethodName = "LastIndexOfAny"
    
    If ArrayExt.IsInitialized(ArrayExt.ToVariantArray(AnyOf)) = False Then
        Errors.OnArgumentError "AnyOf", ModuleName & "." & MethodName
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


' Compares two specified strings, ignoring or honoring their case,
' and returns an integer that indicates their relative position in the sort order.
Public Function Compare(ByVal Str1 As String, _
                        ByVal Str2 As String, _
                        ByVal IgnoreCase As Boolean) As Long
    
    Dim CompMethod As CompareMethod
    CompMethod = IIf(IgnoreCase, CompareMethod.TextCompare, CompareMethod.BinaryCompare)
    Compare = Strings.StrComp(Mid$(Str1, 1, 1), Mid$(Str2, 1, 1), CompMethod)

End Function


' Removes all the leading occurrences of a set of characters specified in an array from the current string.
Public Function TrimStart(ByVal Str As String, ParamArray TrimChars() As Variant) As String

    If UBound(TrimChars) = -1 Then
        TrimStart = TrimWhiteSpaceHelper(Str, TrimType.Head)
    Else
        TrimStart = TrimHelper(Str, TrimChars, UBound(TrimChars), TrimType.Head)
    End If
    
End Function


' Removes all the trailing occurrences of a set of
' characters specified in an array from the current string.
Public Function TrimEnd(ByVal Str As String, ParamArray TrimChars() As Variant) As String

    If UBound(TrimChars) = -1 Then
        TrimEnd = TrimWhiteSpaceHelper(Str, TrimType.Tail)
    Else
        TrimEnd = TrimHelper(Str, TrimChars, UBound(TrimChars), TrimType.Tail)
    End If

End Function


Private Function TrimWhiteSpaceHelper(ByVal Str As String, _
                                      ByVal TrmType As TrimType) As String
                                          
    Dim Finish As Long: Finish = Len(Str)
    Dim Start As Long: Start = 1
    
    If WorksheetFunction.Bitand(TrmType, TrimType.Head) <> 0 Then
    
        For Start = 1 To Len(Str)
            Dim Ch As String: Ch = VBA.Mid$(Str, Start, 1)
            If Ch <> Chr(32) Then ' blank space
                Exit For
            End If
        Next Start
        
    End If
    
    If WorksheetFunction.Bitand(TrmType, TrimType.Tail) <> 0 Then
        
        For Finish = Len(Str) To Start Step -1
            Ch = VBA.Mid$(Str, Finish, 1)
            If Ch <> Chr(32) Then ' blank space
                Exit For
            End If
        Next Finish
        
    End If
    
    TrimWhiteSpaceHelper = CreateTrimmedString(Str, Start, Finish)
                                          
End Function


' Removes all leading and trailing occurrences of a set of characters specified in an array from the current string.
Public Function Trim(ByVal Str As String, ParamArray TrimChars() As Variant) As String
    
    If UBound(TrimChars) = -1 Then
        Trim = VBA.Trim$(Str)
    Else
        Trim = TrimHelper(Str, TrimChars, UBound(TrimChars) + 1, TrimType.Both)
    End If
    
End Function


Private Function TrimHelper(ByVal Str As String, _
                            ByVal TrimChars As Variant, _
                            ByVal Length As Long, _
                            ByVal TrmType As TrimType) As String

    Dim Finish As Long: Finish = Len(Str)
    Dim Start As Long: Start = 1
    Dim TrimCharsLength As Long: TrimCharsLength = UBound(TrimChars) + 1
    
    If WorksheetFunction.Bitand(TrmType, TrimType.Head) <> 0 Then
    
        For Start = 1 To Len(Str)
            Dim i As Long: i = 0
            Dim Ch As String: Ch = Mid(Str, Start, 1)
            
            For i = 0 To TrimCharsLength - 1
                If TrimChars(i) = Ch Then
                    Exit For
                End If
            Next i
            
            If i = TrimCharsLength Then
                Exit For
            End If
            
        Next Start
    
    End If
    
    If WorksheetFunction.Bitand(TrmType, TrimType.Tail) <> 0 Then
        For Finish = Len(Str) To Start Step -1
        
            i = 0
            Ch = Mid(Str, Finish, 1)
            For i = 0 To TrimCharsLength - 1
                If TrimChars(i) = Ch Then
                    Exit For
                End If
            Next i
            
            If i = TrimCharsLength Then
                Exit For
            End If
            
        Next Finish
    End If
    
    TrimHelper = CreateTrimmedString(Str, Start, Finish)

End Function


Private Function CreateTrimmedString(ByVal Str As String, _
                                     ByVal Start As Long, _
                                     ByVal Finish As Long) As String
    
    Dim Ln As Long: Ln = Finish - Start + 1

    If Ln = Len(Str) Then
        CreateTrimmedString = Str
        
    ElseIf Ln = 0 Then
        CreateTrimmedString = vbNullString
        
    Else
        CreateTrimmedString = VBA.Mid$(Str, Start, Ln)
    End If
    
End Function


' Returns a copy of this string converted to uppercase.
Public Function ToUpper(ByVal Str As String) As String
    ToUpper = VBA.UCase$(Str)
End Function


' Returns a copy of this string converted to lowercase.
Public Function ToLower(ByVal Str As String) As String
    ToLower = VBA.LCase$(Str)
End Function

' Returns a new string that left-aligns the characters in this string
' by padding them on the right with a specified Unicode character,
' for a specified total length.
Public Function PadRight(ByVal Str As String, _
                         ByVal TotalWidth As Long, _
                         ByVal PaddingChar As String) As String

    If TotalWidth < 0 Then
        Errors.OnArgumentOutOfRange "TotalWidth", _
                                    "Argument must be positive value. " _
                                    & ModuleName & ".PadRight"
    End If
    
    Dim OldLength As Long: OldLength = VBA.Len(Str)
    Dim Count As Long: Count = TotalWidth - OldLength
    If Count <= 0 Then
        PadRight = Str
        Exit Function
    End If
    
    PadRight = Str & String(Count, PaddingChar)

End Function


' Returns a new string that right-aligns the characters
' in this instance by padding them on the left with a
' specified Unicode character, for a specified total length.
Public Function PadLeft(ByVal Str As String, _
                        ByVal TotalWidth As Long, _
                        ByVal PaddingChar As String) As String

    Const MethodName = "PadLeft"
    
    If TotalWidth < 0 Then
        Errors.OnArgumentOutOfRange "TotalWidth", _
                                    "Argument must be positive value. " _
                                    & ModuleName & "." & MethodName
    End If
    
    Dim OldLength As Long: OldLength = VBA.Len(Str)
    Dim Count As Long: Count = TotalWidth - OldLength
    If Count <= 0 Then
        PadLeft = Str
        Exit Function
    End If
    
    PadLeft = String(Count, PaddingChar) & Str

End Function


' Retrieves a substring from this instance.
' The substring starts at a specified character position and has a specified length.
' If length is not specified, selection continues to the end of the string.
Public Function Substring(ByVal Str As String, _
                          ByVal StartIndex As Long, _
                          Optional ByVal Length As Long = -1) As String
    
    Const MethodName As String = "Substring"
    
    If Length = -1 Then
        Length = VBA.Len(Str)
    End If
    
    If StartIndex < 1 Then
        Errors.OnArgumentOutOfRange "StartIndex", _
                                    "Value must be positive value. " _
                                    & ModuleName & "." & MethodName
    End If
    
    If StartIndex > VBA.Len(Str) Then
        Errors.OnArgumentOutOfRange "StartIndex", _
                                    "Value must be smaller than Length. " _
                                    & ModuleName & "." & MethodName
    End If
    
    If Length < 0 Then
        Errors.OnArgumentOutOfRange "Length", _
                                    "Value must be non-negative. " _
                                    & ModuleName & "." & MethodName
    End If
    
    If StartIndex > VBA.Len(Str) - Length + 1 Then
        Errors.OnArgumentOutOfRange "Length", _
                                    "Value is out of range. " _
                                    & ModuleName & "." & MethodName
    End If
    
    If StartIndex = 1 And Length = VBA.Len(Str) Then
        Substring = Str
        Exit Function
    End If
    
    Substring = VBA.Mid$(Str, StartIndex, Length)
    
End Function


' Returns a new string in which all occurrences
' of a specified string in the current instance
' are replaced with another specified string,
' using the provided comparison type.
Public Function Replace(ByVal Str As String, _
                        ByVal OldValue As String, _
                        ByVal NewValue As String, _
                        ByVal ComparisonType As CompareMethod) As String
    
    Const MethodName = "Replace"
    
    If OldValue = vbNullString Then
        Errors.OnArgumentError "OldValue", _
                               "Value can't be empty string. " _
                               & ModuleName & "." & MethodName
    End If
    
    Replace = VBA.Replace(Str, OldValue, NewValue, 1, -1, ComparisonType)

End Function

