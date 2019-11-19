Attribute VB_Name = "CollectionH"
Option Explicit
'@Folder("Helper")

Private Const ModuleName = "CollectionHelper"


Public Function ToCollectionIList(ByRef List As IList) As Collection

    Dim Output As New Collection
    Dim i As Long
    
    For i = 0 To List.Count - 1
        Output.Add List.GetItem(i)
    Next i
    
    Set ToCollectionIList = Output
    Set Output = Nothing

End Function


Public Function ToString(ByRef Items As Collection, ByVal PropertyName As String) As String
    
    If Items Is Nothing Then
        Exception.ArgumentNullException "Items", ModuleName & ".ToString"
    End If
    
    Dim Item As Variant
    Dim Str As String
    
    If Items.Count = 1 Then
        ToString = CallByName(Items.Item(1), PropertyName, VbGet)
        Exit Function
    End If
    
    For Each Item In Items
        Str = Str & ", " & CallByName(Item, PropertyName, VbGet)
    Next Item
    
    Str = Replace(Str, ", ", vbNullString, 1, 1)
    ToString = Str
    
End Function


' Groups a collection's items using PropertyName value.
' PropertyName value has to be a name of a property of grouped items.
Public Function GroupBy(ByRef Items As Collection, ByVal PropertyName As String) As Scripting.Dictionary

    If Items Is Nothing Then
        Exception.ArgumentNullException "Items", ModuleName & ".GroupBy"
    End If
    
    If PropertyName = vbNullString Then
        Exception.ArgumentException "PropertyName", _
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
        
        Output(Prop).Add Item
        
    Next Item
    
    Set GroupBy = Output
    Set Output = Nothing
    
End Function


Public Function Concat(ByRef Coll1 As Collection, ByRef Coll2 As Collection) As Collection

    If Coll1 Is Nothing Then
        Exception.ArgumentNullException "Coll1", ModuleName & ".Concat"
    End If
    
    If Coll2 Is Nothing Then
        Exception.ArgumentNullException "Coll2", ModuleName & ".Concat"
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
        Exception.ArgumentNullException "Items", ModuleName & "ToArray"
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
        Exception.ArgumentNullException "Items", ModuleName & ".Distinct"
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


Public Function Contains(ByVal Item As Variant, ByRef Items As Collection, ByRef Comparer As IVBAEqualityComparer) As Boolean
    Contains = IndexOf(Item, Items, Comparer) >= 0
End Function


' Returns a collection which contains distinct values from the Collection.
Public Function DistinctValues(ByRef Items As Collection, ByRef Comparer As IVBAEqualityComparer) As Collection
        
    Const MethodName = "DistinctValues"
        
    If Items Is Nothing Then
        Exception.ArgumentNullException "Items", ModuleName & "." & MethodName
    End If
    
    If Comparer Is Nothing Then
        Exception.ArgumentNullException "Comparer", ModuleName & "." & MethodName
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
Public Function IndexOf(ByVal Item As Variant, ByRef Items As Collection, ByRef Comparer As IVBAEqualityComparer) As Long

    Const MethodName = "IndexOf"
    
    If Items Is Nothing Then
        Exception.ArgumentNullException "Items", ModuleName & "." & MethodName
    End If
    
    If Comparer Is Nothing Then
        Exception.ArgumentNullException "Comparer", ModuleName & "." & MethodName
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
        Exception.ArgumentNullException "Container", ModuleName & "." & MethodName
    End If
    
    If Items Is Nothing Then
        Exception.ArgumentNullException "Items", ModuleName & "." & MethodName
    End If
    
    If Items.Count = 0 Then
        Exit Sub
    End If
    
    Dim Item As Variant
    For Each Item In Items
        Container.Add Item
    Next Item

End Sub



