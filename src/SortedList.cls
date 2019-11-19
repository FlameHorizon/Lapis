VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SortedList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'@Folder("Collection")
'@Exposed

Implements IVBACollection
Implements IDictionary

Private Const ClassName As String = "SortedList"
Private Const DefaultCapacity As Long = 16

Private pKeys() As Variant
Private pValues() As Variant
Private pSize As Long
Private pComparer As IComparer
'@Ignore MoveFieldCloserToUsage
Private pKeyList As KeyList
'@Ignore MoveFieldCloserToUsage
Private pValueList As ValueList
Private pEmptyArray() As Variant


Public Sub Init1()
    
    ReDim pKeys(0 To DefaultCapacity - 1)
    ReDim pValues(0 To DefaultCapacity - 1)
    pSize = 0
    Set pComparer = New Comparer

End Sub


Public Sub Init2(ByVal InitialCapacity As Long)

    If InitialCapacity < 0 Then
        Exception.ArgumentOutOfRangeException "InitialCapacity", ClassName & ".Init2"
    End If

    ReDim pKeys(0 To InitialCapacity)
    ReDim pValues(0 To InitialCapacity)
    Set pComparer = New Comparer

End Sub


Public Sub Init3(ByRef Comparer As IComparer)
    
    Init1
    If Not Comparer Is Nothing Then
        Set pComparer = Comparer
    End If

End Sub


Public Sub Init4(ByRef Comparer As IComparer, ByVal Capacity As Long)

    Init3 Comparer
    Me.Capacity = Capacity

End Sub


' Adds an entry with the given key and value to this sorted list.
' An ArgumentException is thrown if the key is already present in the
' sorted list
Public Sub Add(ByRef Key As Variant, ByVal Value As Variant)

    Const MethodName = "Add"
    
    If IsObject(Key) Then
        If Key Is Nothing Then
            Exception.ArgumentNullException "Key", ClassName & "." & MethodName
        End If
    Else
        If Key = vbNullString Then
            Exception.ArgumentNullException "Key", ClassName & "." & MethodName
        End If
    End If

    Dim i As Long
    i = ArrayH.BinarySearch(pKeys, 0, pSize, Key, pComparer)

    If i >= 0 Then
        Exception.ArgumentException "Key", "Item you are trying to add already exists. " & ClassName & "." & MethodName
    End If

    Insert Not i, Key, Value

End Sub


Private Sub IDictionary_Add(ByRef Key As Variant, ByVal Value As Variant)
    Add Key, Value
End Sub


Private Sub Insert(ByVal Index As Long, ByVal Key As Variant, ByVal Value As Variant)
    
    If pSize = ArrayH.Length(pKeys) Then
        EnsureCapacity pSize + 1
    End If
    
    If Index < pSize Then
        ArrayH.Copy pKeys, Index, pKeys, Index + 1, pSize - Index
        ArrayH.Copy pValues, Index, pValues, Index + 1, pSize - Index
    End If
    
    pKeys(Index) = Key
    pValues(Index) = Value
    pSize = pSize + 1

End Sub


Private Sub EnsureCapacity(ByVal Min As Long)

    Dim NewCapacity As Long
    NewCapacity = IIf(ArrayH.Length(pKeys) = 0, 16, ArrayH.Length(pKeys) * 2)

    If NewCapacity < Min Then
        NewCapacity = Min
    End If
    
    Capacity = NewCapacity
    
End Sub


Public Property Get Capacity() As Long
    Capacity = ArrayH.Length(pKeys)
End Property


Public Property Let Capacity(ByVal Value As Long)

    If Value < Count Then
        Exception.ArgumentOutOfRangeException "Value", "Value is smaller than current capacity. " & ClassName & ".Let_Capacity"
    End If
    
    If Value <> ArrayH.Length(pKeys) Then
        If Value <> 0 Then
            Dim NewKeys() As Variant: ReDim NewKeys(0 To Value)
            Dim NewValues() As Variant: ReDim NewValues(0 To Value)
            
            If pSize > 0 Then
                ArrayH.Copy pKeys, 0, NewKeys, 0, pSize
                ArrayH.Copy pValues, 0, NewValues, 0, pSize
            End If
            
            pKeys = NewKeys
            pValues = NewValues
        Else
            pKeys = pEmptyArray
            pValues = pEmptyArray
        End If
    End If

End Property


Private Function IDictionary_GetItem(ByVal Key As Variant) As Variant
    Exception.NotImplementedException vbNullString, ClassName & ".IDictionary_GetItem"
End Function


Private Sub IDictionary_SetItem(ByVal Key As Variant, ByVal Item As Variant)
    Exception.NotImplementedException vbNullString, ClassName & ".IDictionary_SetItem"
End Sub


Public Property Get Keys() As IVBACollection
    Set Keys = GetKeyList
End Property


Private Property Get IDictionary_Keys() As IVBACollection
    Set IDictionary_Keys = Keys
End Property


Public Function GetKeyList() As IList
    
    If pKeyList Is Nothing Then
        Set pKeyList = New KeyList
        pKeyList.Init1 Me
    End If
    
    Set GetKeyList = pKeyList
    
End Function


Public Property Get Values() As IVBACollection
    Set Values = GetValueList()
End Property


Private Property Get IDictionary_Values() As IVBACollection
    Set IDictionary_Values = Values
End Property


Public Function GetValueList() As IList
    
    If pValueList Is Nothing Then
        Set pValueList = New ValueList
        pValueList.Init1 Me
    End If
    
    Set GetValueList = pValueList
    
End Function


Public Function Contains(ByRef Key As Variant) As Boolean
    Contains = (IndexOfKey(Key) >= 0)
End Function


Private Function IDictionary_Contains(ByRef Key As Variant) As Boolean
    IDictionary_Contains = Contains(Key)
End Function


Public Function ContainsKey(ByRef Key As Variant) As Boolean
    ContainsKey = (IndexOfKey(Key) >= 0)
End Function


Public Function IndexOfKey(ByRef Key As Variant) As Long
    
    Dim Output As Long
    Output = ArrayH.BinarySearch(pKeys, 0, pSize, Key, pComparer)
    IndexOfKey = IIf(Output >= 0, Output, -1)
    
End Function


Public Function ContainsValue(ByRef Value As Variant) As Boolean
    ContainsValue = IndexOfValue(Value) >= 0
End Function


Public Function IndexOfValue(ByRef Value As Variant) As Long
    IndexOfValue = ArrayH.IndexOf(pValues, Value, 0, pSize)
End Function


' Removes all entries from this SortedList.
Public Sub Clear()
    
    ArrayH.Clear pKeys, 0, pSize
    ArrayH.Clear pValues, 0, pSize
    pSize = 0
    
End Sub


Private Sub IDictionary_Clear()
    Clear
End Sub


Public Property Get IsReadOnly() As Boolean
    IsReadOnly = False
End Property


Private Property Get IDictionary_IsReadOnly() As Boolean
    IDictionary_IsReadOnly = IsReadOnly
End Property


Public Property Get IsFixedSize() As Boolean
    IsFixedSize = False
End Property


Private Property Get IDictionary_IsFixedSize() As Boolean
    IDictionary_IsFixedSize = IsFixedSize
End Property


Public Sub Remove(ByRef Key As Variant)
    
    Dim i As Long: i = IndexOfKey(Key)
    If i >= 0 Then
        RemoveAt i
    End If
    
End Sub


Private Sub IDictionary_Remove(ByRef Key As Variant)
    Remove Key
End Sub


Public Sub RemoveAt(ByVal Index As Long)
    
    If Index < 0 Or Index >= Count Then
        Exception.ArgumentOutOfRangeException "Index", ClassName & ".RemoveAt"
    End If
    
    pSize = pSize - 1
    If Index < pSize Then
        ArrayH.Copy pKeys, Index + 1, pKeys, Index, pSize - Index
        ArrayH.Copy pValues, Index + 1, pValues, Index, pSize - Index
    End If
    
    pKeys(pSize) = Empty
    pValues(pSize) = Empty
    
End Sub


Public Property Get Count() As Long
    Count = pSize
End Property


Private Property Get IVBACollection_Count() As Long
    IVBACollection_Count = Count
End Property


Public Sub CopyTo(ByRef Arr() As Variant, ByVal Index As Long)
    
    Const MethodName = "CopyTo"
    
    If ArrayH.IsInitialized(Arr) = False Then
        Exception.ArgumentNullException "Arr", "Array is not initalized. " & ClassName & "." & MethodName
    End If
    
    If ArrayH.Rank(Arr) <> 1 Then
        Exception.ArgumentException "Arr", "Multi dimensions array are not supported. " & ClassName & "." & MethodName
    End If
    
    If Index < 0 Then
        Exception.ArgumentOutOfRangeException "Index", ClassName & "." & MethodName
    End If
    
    If ArrayH.Length(Arr) - Index < Count Then
        Exception.ArgumentOutOfRangeException "Index", ClassName & "." & MethodName
    End If
    
    Dim i As Long
    Dim Entry As DictionaryEntry
    
    For i = 0 To Count - 1
        Set Entry = New DictionaryEntry
        Entry.Init1 pKeys(i), pValues(i)
        ArrayH.SetValue Arr, Entry, i + Index
    Next i
    
End Sub


Private Sub IVBACollection_CopyTo(ByRef Arr() As Variant, ByVal Index As Long)
    CopyTo Arr, Index
End Sub


Public Function GetKey(ByVal Index As Long) As Variant

    If Index < 0 Or Index >= Count Then
        Exception.ArgumentOutOfRangeException "Index", ClassName & ".GetKey"
    End If

    GetKey = pKeys(Index)

End Function


' Returns the value of the entry at the given index.
Public Function GetByIndex(ByVal Index As Long) As Variant
    
    If Index < 0 Or Index >= Count Then
        Exception.ArgumentOutOfRangeException "Index", ClassName & ".GetByIndex"
    End If

    GetByIndex = pValues(Index)
    
End Function


' Sets the value at an index to a given value. The previous value of
' the entry is overwritten.
Public Sub SetByIndex(ByVal Index As Long, ByVal Value As Variant)
    
    If Index < 0 Or Index >= Count Then
        Exception.ArgumentOutOfRangeException "Index", ClassName & ".SetByIndex"
    End If
    
    pValues(Index) = Value
    
End Sub


Public Property Get Comparer() As IComparer
    Set Comparer = pComparer
End Property


