Attribute VB_Name = "CollectionExt"
'@Folder("Helper")
Option Explicit

Private Const ModuleName = "CollectionHelper"


' Groups a collection's items using PropertyName value.
' PropertyName value has to be a name of a property of grouped items.
Public Function GroupBy(ByVal Items As Collection, ByVal PropertyName As String) As Scripting.Dictionary

    If Items Is Nothing Then
        Errors.OnArgumentNull "Items", ModuleName & ".GroupBy"
    End If
    
    If PropertyName = vbNullString Then
        Errors.OnArgumentError "PropertyName", _
                               "Value of PropertyName can't be an empty string." & _
                               ModuleName & ".GroupBy"
    End If
    
    Dim Output As New Dictionary
    Dim Prop As String
    Dim Item As Object
    
    For Each Item In Items
        Prop = VBA.CallByName(Item, PropertyName, VbGet)
        
        If Not Output.Exists(Prop) Then
            Output.Add Prop, New Collection
        End If
        
        Output.Item(Prop).Add Item
        
    Next Item
    
    Set GroupBy = Output
    Set Output = Nothing
    
End Function


Public Function Concat(ByVal First As Collection, ByVal Second As Collection) As Collection

    If First Is Nothing Then
        Errors.OnArgumentNull "First", ModuleName & ".Concat"
    End If
    
    If Second Is Nothing Then
        Errors.OnArgumentNull "Second", ModuleName & ".Concat"
    End If
    
    Dim Output As New Collection
    Dim Item As Variant
    
    For Each Item In First
        Output.Add Item
    Next Item
    
    For Each Item In Second
        Output.Add Item
    Next Item
    
    Set Concat = Output
    Set Output = Nothing
    
End Function


Public Function ToArray(ByRef Items As Collection) As Variant()

    If Items Is Nothing Then
        Errors.OnArgumentNull "Items", ModuleName & "ToArray"
    End If

    If Items.Count = 0 Then
        Dim EmptyArr() As Variant
        ToArray = EmptyArr
        Exit Function
    End If
    
    Dim i As Long
    Dim Output() As Variant
    ReDim Output(Items.Count - 1)
    
    For i = 1 To Items.Count
        If IsObject(Items.Item(i)) Then
            Set Output(i - 1) = Items.Item(i)
        Else
            Output(i - 1) = Items.Item(i)
        End If
    Next i
    
    ToArray = Output

End Function


' Returns distinct elements from a sequence.
Public Function Distinct(ByVal Source As Collection, Optional ByVal Comparer As IEqualityComparer) As Collection

    Const MethodName = "Distinct"

    If Source Is Nothing Then
        Lapis.Errors.OnArgumentNull "Source", ModuleName & "." & MethodName
    End If

    Dim Output As New Collection
    Dim Item As Variant
    For Each Item In Source
        If CollectionExt.Contains(Output, Item, Comparer) = False Then
            Output.Add Item
        End If
    Next Item
    
    Set Distinct = Output

End Function


' Determines whether a sequence contains a specified element by using a specified IEqualityComparer.
Public Function Contains(ByVal Source As Collection, _
                         ByVal Value As Variant, _
                         Optional ByVal Comparer As IEqualityComparer) As Boolean
    Contains = IndexOf(Source, Value, Comparer) >= 0
End Function


' Returns an index of an Value in collection.
Public Function IndexOf(ByVal Source As Collection, _
                        ByVal Value As Variant, _
                        Optional ByVal Comparer As IEqualityComparer) As Long

    Const MethodName = "IndexOf"
    
    If Source Is Nothing Then
        Errors.OnArgumentNull "Source", ModuleName & "." & MethodName
    End If
    
    If TypeOf Value Is IComparable Then
        IndexOf = IndexOfUsingComparable(Source, Value)
        Exit Function
    End If
    
    ' If comparer was not defined, try to find default comparer based on the
    ' data type of value.
    If System.IsNothing(Comparer) Then
        Set Comparer = EqualityComparers.Default(Value)
    End If
    
    ' If comparer is still not found, throw the error.
    If System.IsNothing(Comparer) Then
        Lapis.Errors.OnArgumentError "Comparer", _
                                     "Default comparer wasn't found for Value argument. " & ModuleName & "." & MethodName
    Else
        IndexOf = IndexOfUsingEqualityComparer(Source, Value, Comparer)
    End If
        
End Function


Private Function IndexOfUsingComparable(ByVal Source As Collection, ByVal Value As IComparable) As Long

    Dim Ndx As Long: Ndx = 1
    Dim Item As Variant
    For Each Item In Source
        If Value.CompareTo(Item) = 0 Then
            IndexOfUsingComparable = Ndx
            Exit Function
        End If
        Ndx = Ndx + 1
    Next Item
    
    IndexOfUsingComparable = -1

End Function


Private Function IndexOfUsingEqualityComparer(ByVal Source As Collection, _
                                              ByVal Value As Variant, _
                                              ByVal Comparer As IEqualityComparer) As Long
                                             
    Dim Ndx As Long: Ndx = 1
    Dim Item As Variant
    For Each Item In Source
        If Comparer.Equals(Item, Value) Then
            IndexOfUsingEqualityComparer = Ndx
            Exit Function
        End If
        Ndx = Ndx + 1
    Next Item
    
    IndexOfUsingEqualityComparer = -1
                                             
End Function


' Info: Adds every item from Items collection to the Container collection.
' Throws ArgumentNullException when either Container or Items collection are set to nothing.
Public Sub AddRange(ByRef Container As Collection, ByVal Items As Collection)

    Const MethodName = "AddRange"
    
    If Container Is Nothing Then
        Errors.OnArgumentNull "Container", ModuleName & "." & MethodName
    End If
    
    If Items Is Nothing Then
        Errors.OnArgumentNull "Items", ModuleName & "." & MethodName
    End If
    
    If Items.Count = 0 Then
        Exit Sub
    End If
    
    Dim Item As Variant
    For Each Item In Items
        Container.Add Item
    Next Item

End Sub


' Sorts the given collection using the MergeSort algorithm.
' O(n log(n)) time
' O(n) space
Public Function Sort(ByVal Items As Collection, ByVal Comparer As Lapis.IComparer) As Collection
    
    Dim Arr1() As Variant
    Arr1 = CollectionExt.ToArray(Items)
    ArrayExt.Sort Arr1, Comparer
    Set Sort = ArrayExt.ToCollection(Arr1)
    
End Function


' Helper method which allows to create collection in one line.
Public Function Make(ParamArray Items() As Variant) As Collection
    
    Dim Output As New Collection
    Dim Item As Variant
    For Each Item In Items
        Output.Add Item
    Next Item
    
    Set Make = Output
    
End Function


' Produces the set difference of two sequences by using the specified IEqualityComparer to compare values.
Public Function Except(ByVal First As Collection, _
                       ByVal Second As Collection, _
                       Optional ByVal Comparer As IEqualityComparer) As Collection
    
    Const MethodName = "Except"
    
    If First Is Nothing Then
        Lapis.Errors.OnArgumentNull "First", ModuleName & "." & MethodName
    End If
    
    If Second Is Nothing Then
        Lapis.Errors.OnArgumentNull "Second", ModuleName & "." & MethodName
    End If
    
    Dim Output As New Collection
    Dim Item As Variant
    For Each Item In First
        If CollectionExt.Contains(Second, Item, Comparer) = False Then
            Output.Add Item
        End If
    Next Item
    
    Set Except = Output

End Function


' Produces the set intersection of two sequences by using the specified IEqualityComparer to compare values.
Public Function Intersect(ByVal First As Collection, _
                          ByVal Second As Collection, _
                          Optional ByVal Comparer As IEqualityComparer) As Collection

    Const MethodName = "Intersect"
    
    If First Is Nothing Then
        Lapis.Errors.OnArgumentNull "First", ModuleName & "." & MethodName
    End If
    
    If Second Is Nothing Then
        Lapis.Errors.OnArgumentNull "Second", ModuleName & "." & MethodName
    End If
    
    Dim Output As New Collection
    Dim Item As Variant
    For Each Item In First
        If CollectionExt.Contains(Second, Item, Comparer) Then
            Output.Add Item
        End If
    Next Item
    
    Set Intersect = Output

End Function


' Returns the minimum value in a sequence of values.
Public Function Min(ByVal Source As Collection, _
                    Optional ByVal Selector As ICallable) As Variant

    Const MethodName = "Min"
    
    If Source Is Nothing Then
        Lapis.Errors.OnArgumentNull "Source", ModuleName & "." & MethodName
    End If
    
    Dim Comparer As IComparer
    If (Selector Is Nothing) = False Then
        Set Comparer = Comparers.Default(Selector.Run(Source.Item(1)))
    Else
        Set Comparer = Comparers.Default(Source.Item(1))
    End If
    
    ' If comparer is still not found, throw the error.
    If System.IsNothing(Comparer) Then
        Lapis.Errors.OnInvalidOperation "Comparer", _
                                        "Default comparer wasn't found for Value argument. " & ModuleName & "." & MethodName
    End If
    
    ' Do the Min comparison.
    
    ' Find first, non-nothing element in source.
    Dim Value As Variant
    System.CopyVariant Value, Source.Item(1)
    Dim Ndx As Long: Ndx = 2
    Do Until System.IsNothing(Value) = False Or Source.Count < Ndx
        Set Value = Source.Item(Ndx)
        Ndx = Ndx + 1
    Loop
    
    ' Check, if entire source is composed of Nothigns.
    If Source.Count < Ndx Then
        Set Min = Nothing
        Exit Function
    End If
    
    Dim Item As Variant
    For Each Item In Source
        If System.IsNothing(Item) Then
            GoTo NextItem
        End If
        
        Dim IsNewMax As Boolean
        If (Selector Is Nothing) = False Then
            IsNewMax = Comparer.Compare(Selector.Run(Item), Selector.Run(Value)) < 0
        Else
            IsNewMax = Comparer.Compare(Item, Value) < 0
        End If
        
        If IsNewMax Then
            System.CopyVariant Value, Item
        End If
NextItem:
    Next Item
    
    System.CopyVariant Min, Value
   
End Function


' Returns the maximum value in a sequence of values.
Public Function Max(ByVal Source As Collection, _
                    Optional ByVal Selector As ICallable) As Variant

    Const MethodName = "Max"
    
    If Source Is Nothing Then
        Lapis.Errors.OnArgumentNull "Source", ModuleName & "." & MethodName
    End If
    
    Dim Comparer As IComparer
    If (Selector Is Nothing) = False Then
        Set Comparer = Comparers.Default(Selector.Run(Source.Item(1)))
    Else
        Set Comparer = Comparers.Default(Source.Item(1))
    End If
    
    ' If comparer is still not found, throw the error.
    If System.IsNothing(Comparer) Then
        Lapis.Errors.OnArgumentError "Comparer", _
                                     "Default comparer wasn't found for Value argument. " & ModuleName & "." & MethodName
    End If
    
    ' Do the Max comparison.
    
    ' Find first, non-nothing element in source.
    Dim Value As Variant
    System.CopyVariant Value, Source.Item(1)
    Dim Ndx As Long: Ndx = 2
    Do Until System.IsNothing(Value) = False Or Source.Count < Ndx
        Set Value = Source.Item(Ndx)
        Ndx = Ndx + 1
    Loop
    
    ' Check, if entire source is composed of Nothigns.
    If Source.Count < Ndx Then
        Set Max = Nothing
        Exit Function
    End If
    
    Dim Item As Variant
    For Each Item In Source
        If System.IsNothing(Item) Then
            GoTo NextItem
        End If
        
        Dim IsNewMax As Boolean
        If (Selector Is Nothing) = False Then
            IsNewMax = Comparer.Compare(Selector.Run(Item), Selector.Run(Value)) > 0
        Else
            IsNewMax = Comparer.Compare(Item, Value) > 0
        End If
        
        If IsNewMax Then
            System.CopyVariant Value, Item
        End If
NextItem:
    Next Item
    
    System.CopyVariant Max, Value
   
End Function


' Generates a sequence of integral numbers within a specified range.
Public Function Range(ByVal Start As Long, ByVal Count As Long) As Collection
    
    Const MethodName = "Range"
    
    If Count < 0 Then
        Errors.OnArgumentOutOfRange "Count", ModuleName & "." & MethodName
    End If
    
    ' When I'm trying to add two long numbers I get Overflow error.
    ' It's because when adding two Long data types, the result is also Long.
    If CDec(Start) + CDec(Count) > System.LongMaxValue Then
        Errors.OnArgumentOutOfRange vbNullString, ModuleName & "." & MethodName
    End If
    
    Dim Output As New Collection
    Dim i As Long
    For i = Start To Start + Count - 1
        Output.Add i
    Next i
    
    Set Range = Output
    
End Function


' Generates a sequence that contains one repeated value.
Public Function Repeat(ByVal Element As Variant, ByVal Count As Long) As Collection
    
    If Count < 0 Then
        Lapis.Errors.OnArgumentOutOfRange "Count", ModuleName & ".Repeat"
    End If
    
    Dim Output As New Collection
    Dim i As Long
    For i = 1 To Count
        Output.Add Element
    Next i
    
    Set Repeat = Output
    
End Function


' Inverts the order of the elements in a sequence.
Public Function Reverse(ByVal Source As Collection) As Collection

    Const MethodName = "Reverse"
    
    If Source Is Nothing Then
        Lapis.Errors.OnArgumentNull "Source", ModuleName & "." & MethodName
    End If
    
    Dim Output As New Collection
    Dim i As Long
    For i = Source.Count To 1 Step -1
        Output.Add Source.Item(i)
    Next i
    
    Set Reverse = Output

End Function


' Returns a specified number of contiguous elements from the start of a sequence.
Public Function Take(ByVal Source As Collection, ByVal Count As Long) As Collection

    Const MethodName = "Take"

    If Source Is Nothing Then
        Lapis.Errors.OnArgumentNull "Source", ModuleName & "." & MethodName
    End If
    
    If Source.Count = 0 Then
        Set Take = New Collection
        Exit Function
    End If

    Dim Output As New Collection
    Dim i As Long
    For i = 1 To Count
        Output.Add Source.Item(i)
    Next i
    
    Set Take = Output
    
End Function


' Bypasses a specified number of elements in a sequence and then returns the remaining elements.
Public Function Skip(ByVal Source As Collection, ByVal Count As Long) As Collection
    
    If Source Is Nothing Then
        Lapis.Errors.OnArgumentNull "Source", ModuleName & ".Skip"
    End If
    
    Dim Output As New Collection
    Dim i As Long
    For i = Application.Max(Count + 1, 1) To Source.Count
        Output.Add Source.Item(i)
    Next i
    
    Set Skip = Output

End Function


' Determines whether two sequences are equal according to an equality comparer.
Public Function SequenceEqual(ByVal First As Collection, _
                              ByVal Second As Collection, _
                              ByVal Comparer As IEqualityComparer) As Boolean
    
    Const MethodName = "SequenceEqual"
    
    If First Is Nothing Then
        Lapis.Errors.OnArgumentNull "First", ModuleName & "." & MethodName
    End If
    
    If Second Is Nothing Then
        Lapis.Errors.OnArgumentNull "Second", ModuleName & "." & MethodName
    End If
    
    If Comparer Is Nothing Then
        Lapis.Errors.OnArgumentNull "Comparer", ModuleName & "." & MethodName
    End If

    If First.Count <> Second.Count Then
        SequenceEqual = False
        Exit Function
    End If
    
    Dim i As Long
    For i = 1 To First.Count
        If Comparer.Equals(First.Item(i), Second.Item(i)) = False Then
            SequenceEqual = False
            Exit Function
        End If
    Next i
    
    SequenceEqual = True
    
End Function


Private Sub Assing(ByVal Source As Variant, ByRef Destination As Variant)
    
    If IsObject(Source) Then
        Set Destination = Source
    Else
        Destination = Source
    End If
    
End Sub


' Determines whether all elements of a sequence satisfy a condition.
Public Function All(ByVal Source As Collection, ByVal Predicate As ICallable) As Boolean

    Const MethodName = "All"
    
    If Source Is Nothing Then
        Lapis.Errors.OnArgumentNull "Source", ModuleName & "." & MethodName
    End If
    
    If Predicate Is Nothing Then
        Lapis.Errors.OnArgumentNull "Predicate", ModuleName & "." & MethodName
    End If
    
    Dim Item As Variant
    For Each Item In Source
        If System.IsNothing(Item) Then
            All = False
            Exit Function
        End If
    
        If Predicate.Run(Item) = False Then
            All = False
            Exit Function
        End If
    Next Item
    
    All = True

End Function


Public Function Where(ByVal Source As Collection, ByVal Predicate As ICallable) As Collection

    Const MethodName = "Where"

    If Source Is Nothing Then
        Lapis.Errors.OnArgumentNull "Source", ModuleName & "." & MethodName
    End If
    
    If Predicate Is Nothing Then
        Lapis.Errors.OnArgumentNull "Predicate", ModuleName & "." & MethodName
    End If
    
    Dim Output As New Collection
    Dim Item As Variant
    For Each Item In Source
        If Predicate.Run(Item) Then
            Output.Add Item
        End If
    Next Item

    Set Where = Output

End Function


' Determines whether some element of a sequence exists or satisfies a condition.
' Better matching word in this case is Any but it is reserved keyword.
Public Function Some(ByVal Source As Collection, Optional ByVal Predicate As ICallable) As Boolean

    If Source Is Nothing Then
        Lapis.Errors.OnArgumentNull "Source", ModuleName & ".Some"
    End If

    If Predicate Is Nothing Then
        Some = Source.Count <> 0
        Exit Function
    End If
    
    Dim Item As Variant
    For Each Item In Source
        If Predicate.Run(Item) Then
            Some = True
            Exit Function
        End If
    Next Item
    
    Some = False

End Function


' Computes the sum of a sequence of numeric values.
Public Function Sum(ByVal Source As Collection, Optional ByVal Selector As ICallable) As Variant

    Const MethodName = "Sum"

    If Source Is Nothing Then
        Lapis.Errors.OnArgumentNull "Source", ModuleName & "." & MethodName
    End If
    
    Dim Item As Variant
    Dim Output As Variant: Output = 0
    If Selector Is Nothing Then
        For Each Item In Source
            Output = Output + VBA.IIf(System.IsNothing(Item), 0, Item)
        Next Item
        
        Sum = Output
        Exit Function
    End If
    
    For Each Item In Source
        If System.IsNothing(Item) = False Then
            Output = Output + Selector.Run(Item)
        End If
    Next Item
    
    Sum = Output

End Function


' Computes the average of a sequence of numeric values.
Public Function Average(ByVal Source As Collection, Optional ByVal Selector As ICallable) As Variant

    Const MethodName = "Average"
    
    If Source Is Nothing Then
        Lapis.Errors.OnArgumentNull "Source", ModuleName & "." & MethodName
    End If

    If Source.Count = 0 Then
        Average = 0
        Exit Function
    End If
    
    ' Do not take into account Nothing values when calculating average.
    Dim NothingCount As Long
    Dim Item As Variant
    For Each Item In Source
        If System.IsNothing(Item) Then
            NothingCount = NothingCount + 1
        End If
    Next Item
    
    ' Case where the entire source contains only Nothing values.
    If Source.Count - NothingCount = 0 Then
        Average = 0
        Exit Function
    End If
    
    Dim Sum As Variant
    Sum = CollectionExt.Sum(Source, Selector)
    Average = Sum / (Source.Count - NothingCount)
    
End Function


' Returns a number that represents how many elements in the specified sequence satisfy a condition.
Public Function Count(ByVal Source As Collection, Optional ByVal Predicate As ICallable) As Long

    Const MethodName = "Count"

    If Source Is Nothing Then
        Lapis.Errors.OnArgumentNull "Source", MethodName & "." & MethodName
    End If
    
    If Predicate Is Nothing Then
        Count = Source.Count
        Exit Function
    End If
    
    Count = CollectionExt.Where(Source, Predicate).Count

End Function


' Projects each element of a sequence into a new form.
Public Function Convert(ByVal Source As Collection, ByVal Selector As ICallable) As Collection

    Const MethodName As String = "Convert"
    
    If Source Is Nothing Then
        Lapis.Errors.OnArgumentNull "Source", ModuleName & "." & MethodName
    End If
    
    If Selector Is Nothing Then
        Lapis.Errors.OnArgumentNull "Selector", ModuleName & "." & MethodName
    End If
    
    Dim Output As New Collection
    Dim Item As Variant
    For Each Item In Source
        Output.Add Selector.Run(Item)
    Next Item
    
    Set Convert = Output

End Function


' Returns the first element in a sequence that satisfies a specified condition.
Public Function First(ByVal Source As Collection, Optional ByVal Predicate As ICallable) As Variant
    
    Const MethodName = "First"
    
    If Source Is Nothing Then
        Lapis.Errors.OnArgumentNull "Source", ModuleName & "." & MethodName
    End If
    
    If Source.Count = 0 Then
        Lapis.Errors.OnInvalidOperation "Source", ModuleName & "." & MethodName
    End If
    
    If Predicate Is Nothing Then
        Assign First, Source.Item(1)
        Exit Function
    End If
    
    Dim Item As Variant
    For Each Item In Source
        If Predicate.Run(Item) Then
            Assign First, Item
            Exit Function
        End If
    Next Item
    
    Lapis.Errors.OnInvalidOperation vbNullString, ModuleName & "." & MethodName

End Function


Private Sub Assign(ByRef Destination As Variant, ByVal Source As Variant)
    System.CopyVariant Destination, Source
End Sub


' Returns the last element of a sequence that satisfies a specified condition.
Public Function Last(ByVal Source As Collection, Optional ByVal Predicate As ICallable) As Variant

    Const MethodName = "Last"
    
    If Source Is Nothing Then
        Lapis.Errors.OnArgumentNull "Source", ModuleName & "." & MethodName
    End If
    
    If Source.Count = 0 Then
        Lapis.Errors.OnInvalidOperation "Source", ModuleName & "." & MethodName
    End If
    
    If Predicate Is Nothing Then
        Assign Last, Source.Item(Source.Count)
        Exit Function
    End If

    Dim Output As Variant
    Dim Item As Variant
    For Each Item In Source
        If Predicate.Run(Item) Then
            Assign Output, Item
        End If
    Next Item
    
    ' No item matches the predicate or source is empty.
    If VBA.IsObject(Output) Then
        If Not (Output Is Nothing) Then
            Assign Last, Output
            Exit Function
        End If
    ElseIf Output = vbEmpty Then
        Lapis.Errors.OnInvalidOperation vbNullString, ModuleName & "." & MethodName
    End If
    
    Assign Last, Output
    
End Function


' Returns the only element of a sequence that satisfies a specified condition,
' and throws an exception if more than one such element exists.
Public Function SelectOne(ByVal Source As Collection, Optional ByVal Predicate As ICallable) As Variant

    Const MethodName = "SelectOne"

    If Source Is Nothing Then
        Lapis.Errors.OnArgumentNull "Source", MethodName & "." & MethodName
    End If
    
    If Source.Count = 0 Then
        Lapis.Errors.OnInvalidOperation "Source", _
                                        "The input sequence is empty. " _
                                        & ModuleName & "." & MethodName
    End If

    
    If Predicate Is Nothing And Source.Count > 1 Then
        Lapis.Errors.OnInvalidOperation vbNullString, _
                                        "The input sequence contains more than one element. " _
                                        & ModuleName & "." & MethodName
    ElseIf Source.Count = 1 Then
        Assign SelectOne, Source.Item(1)
        Exit Function
    End If

    Dim Output As Variant
    Dim Item As Variant
    For Each Item In Source
        If Predicate.Run(Item) Then
            If Output <> vbEmpty Then
                Lapis.Errors.OnInvalidOperation vbNullString, ModuleName & "." & MethodName
            Else
                Assign Output, Item
            End If
        End If
    Next Item
    
    ' No item matches the predicate or source is empty.
    If Output = vbEmpty Then
        Lapis.Errors.OnInvalidOperation vbNullString, ModuleName & "." & MethodName
    End If
    
    Assign SelectOne, Output

End Function

