Attribute VB_Name = "StringExt"
Option Explicit
'@Folder("Helper")

Private Const ModuleName As String = "StringExtelper"


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
        IndexOfAny = IndexOf(Str, AnyOf(i), 0, Len(Str), vbTextCompare)

        If IndexOfAny <> 0 Then
            Exit Function
        End If
    Next i

    IndexOfAny = 0

End Function


Public Function IndexOf(ByVal Str As String, _
                        ByVal Value As String, _
                        Optional ByVal StartIndex As Long = System.LongMinValue, _
                        Optional ByVal Count As Long = System.LongMinValue, _
                        Optional ByVal ComparisonMethod As VbCompareMethod = VbCompareMethod.vbTextCompare) As Long
    
    Const MethodName = "IndexOf"
    
    Dim Start As Long
    Start = IIf(StartIndex = System.LongMinValue, 0, StartIndex)
    
    Dim Cnt As Long
    Cnt = IIf(Count = System.LongMinValue, VBA.Len(Str) - StartIndex, Count)
    
    If Str = vbNullString Then
        Errors.OnArgumentError "Str", _
                               "Str value can't be empty string. " _
                               & ModuleName & "." & MethodName
    End If
    
    If Start < 0 Or Start > Len(Str) Then
        Errors.OnArgumentOutOfRange "Start", ModuleName & "." & MethodName
    End If
    
    If Count < 0 Or Start > Len(Str) - Cnt Then
        Errors.OnArgumentOutOfRange "Cnt", ModuleName & "." & MethodName
    End If
    
    Select Case ComparisonMethod
        Case VbCompareMethod.vbBinaryCompare, VbCompareMethod.vbTextCompare
            IndexOf = InStr(StartIndex + 1, Str, Value, ComparisonMethod) - 1
        
        Case Else
            Errors.OnArgumentOutOfRange "ComparisonMethod", "Not supported string comparison. " & MethodName & "." & MethodName
            
    End Select
    
End Function


Public Function Contains(ByVal Str As String, ByVal Value As String) As Boolean
    Contains = IndexOf(Str, Value, 0, Len(Str), VbCompareMethod.vbTextCompare) >= 0
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


Public Function StartsWith(ByVal Str As String, ByVal Value As String, ByVal StringComparison As VbCompareMethod) As Boolean

    Const MethodName = "StartsWith"
    
    If StringComparison = VbCompareMethod.vbDatabaseCompare Then
        Errors.OnArgumentOutOfRange "StringComparison", ModuleName & "." & MethodName
    End If
    
    If Str = vbNullString Then
        Errors.OnArgumentError "Str", ModuleName & "." & MethodName
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


