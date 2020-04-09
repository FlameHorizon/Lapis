Attribute VB_Name = "CollectionExt"
Option Explicit
'@Folder("Helper")

Private Const ModuleName = "CollectionHelper"


Public Function ToString(ByVal Items As Collection, _
                         ByVal Converter As IToString, _
                         Optional ByVal Delimiter As String = ",") As String

    Const MethodName = "ToString"

    If Items Is Nothing Then
        Lapis.Errors.OnArgumentNull "Items", ModuleName & "." & MethodName
    End If

    If Converter Is Nothing Then
        Lapis.Errors.OnArgumentNull "Converter", ModuleName & "." & MethodName
    End If

    If Items.Count = 0 Then
        ToString = vbNullString
        Exit Function
    End If
    
    Dim Output As String
    Dim Item As Variant
    For Each Item In Items
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
' of ToString method of each object within Items collecion.
Public Function ToStringByProperty(ByVal Items As Collection, _
                                   ByVal PropertyName As String) As String
    
    If Items Is Nothing Then
        Errors.OnArgumentNull "Items", ModuleName & ".ToStriToStringByPropertyng"
    End If
    
    Dim Converter As New PropertyToStringConverter
    Converter.PropertyName = PropertyName
    ToStringByProperty = CollectionExt.ToString(Items, Converter)

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


Public Function Concat(ByRef Coll1 As Collection, ByRef Coll2 As Collection) As Collection

    If Coll1 Is Nothing Then
        Errors.OnArgumentNull "Coll1", ModuleName & ".Concat"
    End If
    
    If Coll2 Is Nothing Then
        Errors.OnArgumentNull "Coll2", ModuleName & ".Concat"
    End If
    
    Dim Output As New Collection
    Dim Item As Variant
    
    For Each Item In Coll1
        Output.Add Item
    Next Item
    
    For Each Item In Coll2
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


Public Function Distinct(ByRef Items As Collection, ByVal PropertyName As String) As Collection
    
    If Items Is Nothing Then
        Errors.OnArgumentNull "Items", ModuleName & ".Distinct"
    End If

    Dim CompareList As New Dictionary
    Dim Output As New Collection
    Dim PropValue As Variant
    Dim Item As Object
    
    For Each Item In Items
        PropValue = CallByName(Item, PropertyName, VbGet)
        
        If Not CompareList.Exists(PropValue) Then
            CompareList.Add PropValue, PropValue
            Output.Add Item
        End If
        
    Next Item
    
    Set Distinct = Output
    Set Output = Nothing
    Set CompareList = Nothing
    
End Function


Public Function Contains(ByVal Item As Variant, ByRef Items As Collection, ByRef Comparer As IEqualityComparer) As Boolean
    Contains = IndexOf(Item, Items, Comparer) >= 0
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
        If Not Contains(Item, Output, Comparer) Then
            Output.Add Item
        End If
        
    Next Item
    
    Set DistinctValues = Output
    Set Output = Nothing
    
End Function


' Returns an index of an item in collection.
Public Function IndexOf(ByVal Item As Variant, ByRef Items As Collection, ByRef Comparer As IEqualityComparer) As Long

    Const MethodName = "IndexOf"
    
    If Items Is Nothing Then
        Errors.OnArgumentNull "Items", ModuleName & "." & MethodName
    End If
    
    If Comparer Is Nothing Then
        Errors.OnArgumentNull "Comparer", ModuleName & "." & MethodName
    End If
    
    Dim Ndx As Long: Ndx = 1
    Dim Element As Variant
    
    For Each Element In Items
        If Comparer.Equals(Element, Item) Then
            IndexOf = Ndx
            Exit Function
        End If
        Ndx = Ndx + 1
    Next Element
    
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



