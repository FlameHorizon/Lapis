Attribute VB_Name = "CollectionExt"
Option Explicit
'@Folder("Helper")

Private Const ModuleName = "CollectionHelper"


Public Function ToString(ByVal Source As Collection, _
                         ByVal Converter As IToString, _
                         Optional ByVal Delimiter As String = ",") As String

    Const MethodName = "ToString"

    If Source Is Nothing Then
        Lapis.Errors.OnArgumentNull "Source", ModuleName & "." & MethodName
    End If

    If Converter Is Nothing Then
        Lapis.Errors.OnArgumentNull "Converter", ModuleName & "." & MethodName
    End If

    If Source.Count = 0 Then
        ToString = vbNullString
        Exit Function
    End If
    
    Dim Output As String
    Dim Item As Variant
    For Each Item In Source
        On Error Resume Next
        Output = Output & Delimiter & Converter.ToString(Item)
        If Err.Number = ErrorNumber.ObjectRequired Or Err.Number = ErrorNumber.TypeMismatch Then
            On Error GoTo 0
            Errors.OnInvalidOperation vbNullString, _
                                      "Given convert was not able to convert value of collection into string. " _
                                      & ModuleName & "." & MethodName
                                       
        ElseIf Err.Number = ErrorNumber.ObjectVariableOrWithBlockVariableNotSet Then
            On Error GoTo 0
            Errors.OnInvalidOperation vbNullString, _
                                      "Given item, inside collection is not set. " _
                                      & ModuleName & "." & MethodName
        
        ElseIf Err.Number = ErrorNumber.ObjectDoesntSupportThisPropertyOrMethod Then
            On Error GoTo 0
            Errors.OnArgumentOutOfRange vbNullString, _
                                        "Given property is not a party of object. " _
                                        & ModuleName & "." & MethodName
        
        End If
        On Error GoTo 0
 
    Next Item
    
    Output = StringExt.RemoveRange(Output, 0, Len(Delimiter))
    ToString = Output

End Function


' Returns a string which represents collection of objects based on the implementation
' of ToString method of each object within Source collecion.
Public Function ToStringByProperty(ByVal Source As Collection, ByVal PropertyName As String) As String
    
    If Source Is Nothing Then
        Errors.OnArgumentNull "Source", ModuleName & ".ToStriToStringByPropertyng"
    End If
    
    Dim Converter As New PropertyToStringConverter
    Converter.PropertyName = PropertyName
    ToStringByProperty = CollectionExt.ToString(Source, Converter)

End Function


' Groups a collection's items using PropertyName value.
' PropertyName value has to be a name of a property of grouped items.
Public Function GroupBy(ByRef Items As Collection, ByVal PropertyName As String) As Scripting.Dictionary

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
Public Function Distinct(ByVal Source As Collection, ByVal Comparer As IEqualityComparer) As Collection

    Const MethodName = "Distinct"

    If Source Is Nothing Then
        Lapis.Errors.OnArgumentNull "Source", ModuleName & "." & MethodName
    End If
    
    If Comparer Is Nothing Then
        Lapis.Errors.OnArgumentNull "Comparer", MethodName & "." & MethodName
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
                         ByVal Comparer As IEqualityComparer) As Boolean
    Contains = IndexOf(Source, Value, Comparer) >= 0
End Function


' Returns a collection which contains distinct values from the Collection.
Public Function DistinctValues(ByRef Items As Collection, ByRef Comparer As IEqualityComparer) As Collection
        
    Const MethodName = "DistinctValues"
        
    If Items Is Nothing Then
        Errors.OnArgumentNull "Items", ModuleName & "." & MethodName
    End If
    
    If Comparer Is Nothing Then
        Errors.OnArgumentNull "Comparer", ModuleName & "." & MethodName
    End If
    
    Dim Output As New Collection
    Dim Item As Variant
    
    For Each Item In Items
        If Not Contains(Output, Item, Comparer) Then
            Output.Add Item
        End If
        
    Next Item
    
    Set DistinctValues = Output
    Set Output = Nothing
    
End Function


' Returns an index of an Value in collection.
Public Function IndexOf(ByVal Source As Collection, _
                        ByVal Value As Variant, _
                        ByVal Comparer As IEqualityComparer) As Long

    Const MethodName = "IndexOf"
    
    If Source Is Nothing Then
        Errors.OnArgumentNull "Source", ModuleName & "." & MethodName
    End If
    
    If Comparer Is Nothing Then
        Errors.OnArgumentNull "Comparer", ModuleName & "." & MethodName
    End If
    
    Dim Ndx As Long: Ndx = 1
    Dim Item As Variant
    
    For Each Item In Source
        If Comparer.Equals(Item, Value) Then
            IndexOf = Ndx
            Exit Function
        End If
        Ndx = Ndx + 1
    Next Item
    
    IndexOf = -1

End Function


' Info: Adds every item from Items collection to the Container collection.
' Throws ArgumentNullException when either Container or Items collection are set to nothing.
Public Sub AddRange(ByRef Container As Collection, ByRef Items As Collection)

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
Public Function Sort(ByVal Items As Collection, Optional ByVal Comparer As Lapis.IComparer) As Collection
    
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
                       ByVal Comparer As IEqualityComparer) As Collection
    
    Const MethodName = "Except"
    Lapis.Errors.OnNotImplemented vbNullString, "CollectionExt.Except"

End Function




