Attribute VB_Name = "ArrayExt"
Option Explicit
'@Folder("Helper")

Private Const ModuleName As String = "ArrayHelper"


' Item paramter has to be a simple type.
' Arr has to have only one dimension.
Public Function Exists(ByVal Item As Variant, ByRef Arr As Variant) As Boolean
    Exists = (UBound(Filter(Arr, Item)) > -1)
End Function


' Retruns true if array was initalized.
' In VB, for whatever reason, Not myArray returns the SafeArray pointer.
' For uninitialized arrays, this returns -1.
' Not (XOR) this to XOR it with -1, thus returning zero.
Public Function IsInitialized(ByRef Arr() As Variant) As Boolean
    IsInitialized = ((Not Not Arr) <> 0)
End Function


' Converts single dimension array into a collection.
Public Function ToCollection(ByRef Arr() As Variant) As Collection

    Dim Output As New Collection
    
    Dim Item As Variant
    For Each Item In Arr
        Output.Add Item
    Next

    Set ToCollection = Output
    Set Output = Nothing
    
End Function


' Copies elements from an Array starting at SourceIndex and pastes them to another
' Array starting at DestinationIndex. Number of elements which will be copied is
' is specified in Length parameter.
Public Sub Copy(ByRef Arr() As Variant, _
                ByVal SourceIndex As Long, _
                ByRef DestinationArray() As Variant, _
                ByVal DestinationIndex As Long, _
                ByVal Length As Long)

    Dim DestNdx As Long
    DestNdx = DestinationIndex
    
    Dim Ndx As Long
    For Ndx = SourceIndex To (Length + SourceIndex - 1)
        DestinationArray(DestNdx) = Arr(Ndx)
        DestNdx = DestNdx + 1
    Next Ndx
    
End Sub


' This function returns the number of dimensions of an array. An unallocated dynamic array
' has 0 dimensions.
Public Function Rank(ByRef Arr() As Variant) As Long
    
    Dim Ndx As Long
    '@Ignore VariableNotUsed
    Dim Res As Long
    On Error Resume Next

    Do
        Ndx = Ndx + 1
        Res = UBound(Arr, Ndx)
    Loop Until Err.Number <> 0
    On Error GoTo 0
    Err.Number = 0
    
    Rank = Ndx - 1
    
End Function


' Returns the number of elements in single dimension of array.
Public Function Length(ByRef Arr() As Variant) As Long
    
    Dim Output As Long
    
    Dim i As Long
    For i = 1 To Rank(Arr)
        Output = UBound(Arr) + 1
    Next i
    
    Length = Output
    
End Function


' Returns the number of elements in the specified dimension (Dimension) of the array in
' Arr. If you omit Dimension, the first dimension is used.
Public Function NumElements(ByRef Arr() As Variant, Optional ByVal Dimension As Long = 1) As Long

    Dim NumDimensions As Long
    
    If IsArray(Arr) = False Then
        NumElements = 0
        Exit Function
    End If
    
    If IsEmpty(Arr) = True Then
        NumElements = 0
        Exit Function
    End If
    
    If Dimension < 0 Then
        NumElements = 0
        Exit Function
    End If
    
    NumDimensions = ArrayExt.Rank(Arr)
    If NumDimensions < Dimension Then
        NumElements = 0
        Exit Function
    End If
    
    NumElements = UBound(Arr, Dimension) - LBound(Arr, Dimension) + 1

End Function


' Clears an range of items in Array starting at Index.
' This method works only for Arrays which have a single dimensions.
Public Sub Clear(ByRef Arr() As Variant, ByVal Index As Long, ByVal Length As Long)
    
    Const MethodName = "Clear"
    
    If Index < 0 Then
        Errors.OnArgumentOutOfRange "Index", ModuleName & "." & MethodName
    End If
    
    If Index < LBound(Arr) Then
        Errors.OnArgumentOutOfRange "Index", ModuleName & "." & MethodName
    End If
    
    If Index + Length > ArrayExt.Length(Arr) Then
        Errors.OnArgumentOutOfRange vbNullString, ModuleName & "." & MethodName
    End If
    
    If ArrayExt.Rank(Arr) <> 1 Then
        Errors.OnArgumentError "Arr", _
                                    "Multi dimensions arrray are not supported. " & ModuleName & MethodName
    End If
    
    Dim Ndx As Long
    Ndx = Index
    
    Dim i As Long
    For i = 1 To Length
        Arr(Ndx) = 0
        Ndx = Ndx + 1
    Next i
    
End Sub


' Returns a String which contains every element in an Array recursively.
' Method will work only if Array contains value type elements (strings, integers, long, ect.)
Public Function ToString(ByRef Arr() As Variant) As String

    Dim Output As String: Output = vbNullString
    ToStringRecursively Arr, Output
    ToString = Output

End Function


Private Sub ToStringRecursively(ByRef Arr() As Variant, ByRef Output As String)
    
    Dim i As Long
    For i = 0 To UBound(Arr)
        If IsArray(Arr(i)) Then
            Dim InnerArr() As Variant
            InnerArr = Arr(i)
            ToStringRecursively InnerArr, Output
        Else
            Output = Output & ", " & Arr(i)
        End If
    Next i
    
    Output = VBA.Replace(Output, ", ", vbNullString, Start:=1, Count:=1)

End Sub


' Searches a section of an array for a given element using a binary search
' algorithm. Elements of the array are compared to the search value using
' the given IComparer interface. If comparer is null,
' elements of the array are compared to the search value using the
' IComparable interface, defualt string comparer is used to compare
' all elements. This method assumes that the array is already sorted;
' if this is not the case, the result will be incorrect.
'
' The method returns the index of the given value in the array. If the
' array does not contain the given value, the method returns a negative
' integer. The bitwise complement operator (NOT) can be applied to a
' negative result to produce the index of the first element (if any) that
' is larger than the given search value.

Public Function BinarySearch(ByRef Arr() As Variant, _
                             ByVal Index As Long, _
                             ByVal Length As Long, _
                             ByRef Value As Variant, _
                             ByRef Comparer As IComparer) As Long
    
    Const MethodName = "BinarySearch"
    
    If ArrayExt.IsInitialized(Arr) = False Then
        Errors.OnArgumentNull "Arr", "Array is not initialized. " & ModuleName & "." & MethodName
    End If
    
    Dim Lb As Long: Lb = ArrayExt.GetLowerBound(Arr, 0)
    
    If Index < Lb Or ArrayExt.Length(Arr) < 0 Then
        Errors.OnArgumentNull "Arr", "Index must be a non-negative numer. " & ModuleName & "." & MethodName
    End If
    
    If ArrayExt.Length(Arr) - (Index - Lb) < Length Then
        Errors.OnArgumentError "Length", "Invalid offset length. " & ModuleName & "." & MethodName
    End If
    
    If ArrayExt.Rank(Arr) <> 1 Then
        Errors.OnArgumentError "Arr", "Multi dimensions array are not supported. " & ModuleName & ".BinarySearch"
    End If
    
    If Comparer Is Nothing Then
        Set Comparer = ComparerStatic.Default
    End If
    
    Dim Low As Long: Low = Index
    Dim High As Long: High = Index + Length - 1
    Dim ObjArr() As Variant: ObjArr = Arr
    Dim i As Long
    Dim c As Long
    
    Do While Low <= High
        i = Application.WorksheetFunction.Median(Low, High)
        c = Comparer.Compare(ObjArr(i), Value)
        
        If c = 0 Then
            BinarySearch = i
            Exit Function
        End If
        
        If c < 0 Then
            Low = i + 1
        Else
            High = i - 1
        End If
    
    Loop
    
    BinarySearch = Not Low
    
End Function


' Return the index of the first element of the specified dimension in the array.
' Dimension: A zero-based dimension of the array whose starting index needs to be determined.
Public Function GetLowerBound(ByRef Arr() As Variant, ByVal Dimension As Long) As Long

    Const MethodName = "GetLowerBound"

    If Dimension < 0 Or Dimension >= ArrayExt.Rank(Arr) Then
        Errors.OnArgumentOutOfRange "Dimension", ModuleName & "." & MethodName
    End If
    
    If ArrayExt.IsInitialized(Arr) = False Then
        Errors.OnArgumentNull "Arr", "Array is not initalized. " & ModuleName & "." & MethodName
    End If
    
    GetLowerBound = LBound(Arr, Dimension + 1)

End Function


' Returns the index of the first occurrence of a given value in a range of
' an array. The array is searched fowards, start from the index StartIndex value up to Count element.
Public Function IndexOf(ByRef Arr() As Variant, ByRef Value As Variant, ByVal StartIndex As Long, ByVal Count As Long) As Long
    
    Const MethodName = "IndexOf"
    
    If ArrayExt.IsInitialized(Arr) = False Then
        Errors.OnArgumentNull "Arr", "Array is not initialized. " & ModuleName & "." & MethodName
    End If
    
    If ArrayExt.Rank(Arr) <> 1 Then
        Errors.OnArgumentError "Arr", "Multi dimensions array are not supported. " & ModuleName & "." & MethodName
    End If
    
    Dim Lb As Long
    Lb = ArrayExt.GetLowerBound(Arr, 0)
    
    If (StartIndex < Lb) Or (StartIndex > ArrayExt.Length(Arr) + Lb) Then
        Errors.OnArgumentOutOfRange "StartIndex", ModuleName & "." & MethodName
    End If
    
    If (StartIndex < Lb) Or (StartIndex > ArrayExt.Length(Arr) + Lb) Then
        Errors.OnArgumentOutOfRange "Count", ModuleName & "." & MethodName
    End If
    
    Dim ObjArray() As Variant
    ObjArray = Arr
    
    Dim EndIndex As Long
    EndIndex = StartIndex + Count
    
    Dim IsEmptyOrNothing As Boolean
    
    If IsEmpty(Value) Then
        IsEmptyOrNothing = True
    ElseIf IsObject(Value) Then
        IsEmptyOrNothing = (Value Is Nothing)
    End If
    
    Dim i As Long
    
    If IsEmptyOrNothing Then
        
        For i = StartIndex To EndIndex - 1
            If IsEmpty(ObjArray(i)) Or (ObjArray(i) Is Nothing) Then
                IndexOf = i
                Exit Function
            End If
        Next i
    Else
    
        For i = StartIndex To EndIndex - 1
            Dim Item As Variant: Item = ObjArray(i)
            If IsEmpty(Item) = False And (Item = Value) Then
                IndexOf = i
                Exit Function
            End If
        Next i
    End If
    
End Function


Public Sub SetValue(ByRef Arr() As Variant, ByRef Value As Variant, ByVal Index As Long)
    
    If Rank(Arr) <> 1 Then
        Errors.OnArgumentError "Arr", "Multi dimensions array are not supported. " & ModuleName & ".SetValue"
    End If
    
    If IsObject(Value) Then
        Set Arr(Index) = Value
    Else
        Arr(Index) = Value
    End If
    
End Sub


' Converts any array to the Variant Array. When input array is not initalized
' method returns not initialized Variant Array.
' Throws ArgumentException when Arr is not an array.
Public Function ToVariantArray(ByRef Arr As Variant) As Variant()

    Const MethodName = "ToVariantArray"
    
    If VBA.IsArray(Arr) = False Then
        Errors.OnArgumentError "Arr", ModuleName & "." & MethodName
    End If

    ' Here, we still can't use defined methods like IsInitalized because
    ' Arr is still a Variant type, as oposed to the Array of Variant elements.
    ' We have to check manually if Array is initiliazed.
    
    '@Ignore AssignmentNotUsed
    Dim Low As Long: Low = -1
    '@Ignore AssignmentNotUsed
    Dim High As Long: High = -1
    
    On Error Resume Next
    Low = LBound(Arr)
    High = UBound(Arr)
    On Error GoTo 0 ' Cancel "Ignore erorrs"
    
    ' At this point we know, that Arr is an array, but it is not initialized.
    ' Return not initalized array back.
    If Low = -1 Or High = -1 Then
        Dim NotInit() As Variant
        ToVariantArray = NotInit
        Exit Function
    End If

    Dim Output() As Variant
    ReDim Output(Low To High)
    
    Dim Ndx As Long
    For Ndx = Low To High
        Output(Ndx) = Arr(Ndx)
    Next Ndx
    
    ToVariantArray = Output

End Function


' Returns an Array of String values. This helper method enables to
' create inline array.
Public Function StringArray(ParamArray Items() As Variant) As String()
    
    Dim UpperLimit As Long
    UpperLimit = UBound(Items)
    
    Dim Output() As String
    ReDim Output(0 To UpperLimit)
    
    Dim i As Long
    For i = 0 To UpperLimit
        Output(i) = Items(i)
    Next i
    
    StringArray = Output
    
End Function





