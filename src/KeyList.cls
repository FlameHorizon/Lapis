VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "KeyList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'@Folder("Collection")

Implements IList
Implements IVBACollection

Private Const ClassName As String = "KeyList"

Private pSortedList As SortedList


Public Sub Init1(ByRef SortedList As SortedList)
    Set pSortedList = SortedList
End Sub


Public Property Get Count() As Long
    Count = pSortedList.Count
End Property


Private Property Get IVBACollection_Count() As Long
    IVBACollection_Count = Count
End Property


Private Property Get IList_Count() As Long
    IList_Count = Count
End Property


Public Property Get IsReadOnly() As Boolean
    IsReadOnly = pSortedList.IsReadOnly
End Property


Private Property Get IList_IsReadOnly() As Boolean
    IList_IsReadOnly = IsReadOnly
End Property


Public Property Get IsFixedSize() As Boolean
    IsFixedSize = pSortedList.IsFixedSize
End Property


Private Property Get IList_IsFixedSize() As Boolean
    IList_IsFixedSize = IsFixedSize
End Property


Public Function Add(ByVal Key As Variant) As Long
    Exception.NotSupportedException vbNullString, "Nested write is not supported inside SortedList. " & ClassName & ".Add"
End Function


Private Function IList_Add(ByVal Value As Variant) As Long
    Add Value
End Function


Public Sub Clear()
    Exception.NotSupportedException vbNullString, "Nested write is not supported inside SortedList. " & ClassName & ".Clear"
End Sub


Private Sub IList_Clear()
    Clear
End Sub


Public Function Contains(ByRef Key As Variant) As Boolean
    Contains = pSortedList.Contains(Key)
End Function


Private Function IList_Contains(ByRef Value As Variant) As Boolean
    IList_Contains = Contains(Value)
End Function


Public Sub CopyTo(ByRef Arr() As Variant, ByVal ArrayIndex As Long)
    
    If (ArrayH.IsInitialized(Arr)) And (ArrayH.Rank(Arr) <> 1) Then
        Exception.ArgumentException "Arr", "Multi dimensions array are not supported. " & ClassName & ".CopyTo"
    End If
    
    ArrayH.Copy CollectionH.ToArray(pSortedList.Keys), 0, Arr, ArrayIndex, pSortedList.Count
    
End Sub


Private Sub IVBACollection_CopyTo(ByRef Arr() As Variant, ByVal Index As Long)
    CopyTo Arr, Index
End Sub


Private Sub IList_CopyTo(ByRef Arr() As Variant, ByVal Index As Long)
    CopyTo Arr, Index
End Sub


Public Sub Insert(ByVal Index As Long, ByVal Value As Variant)
    Exception.NotSupportedException vbNullString, "Nested write is not supported inside SortedList. " & ClassName & ".Insert"
End Sub


Private Sub IList_Insert(ByVal Index As Long, ByVal Value As Variant)
    Insert Index, Value
End Sub


Public Function GetKey(ByVal Index As Long) As Variant
    GetKey = pSortedList.Keys
End Function


Public Function GetItem(ByVal Index As Long) As Variant
    GetItem = pSortedList.GetKey(Index)
End Function


Private Function IList_GetItem(ByVal Index As Long) As Variant
    IList_GetItem = GetItem(Index)
End Function


Private Sub IList_SetItem(ByVal Index As Long, ByVal Value As Variant)
    Exception.NotSupportedException vbNullString, ClassName & ".IList_SetItem"
End Sub


Public Function IndexOf(ByRef Value As Variant) As Long
    
    If IsEmpty(Value) Then
        Exception.ArgumentNullException "Value", ClassName & ".IndexOf"
    End If
    
    Dim i As Long
    i = ArrayH.BinarySearch(ArrayH.ToArrayIList(pSortedList.Keys), 0, pSortedList.Count, Value, pSortedList.Comparer)
    
    If i >= 0 Then
        IndexOf = i
    Else
        IndexOf = -1
    End If
    
End Function


Private Function IList_IndexOf(ByRef Value As Variant) As Long
    IList_IndexOf = IndexOf(Value)
End Function


Public Sub Remove(ByVal Value As Variant)
    Exception.NotSupportedException vbNullString, "Nested write is not supported inside SortedList. " & ClassName & ".Remove"
End Sub


Private Sub IList_Remove(ByVal Value As Variant)
    Remove Value
End Sub


Public Sub RemoveAt(ByVal Value As Variant)
    Exception.NotSupportedException vbNullString, "Nested write is not supported inside SortedList. " & ClassName & ".RemoveAt"
End Sub


Private Sub IList_RemoveAt(ByVal Index As Long)
    RemoveAt Index
End Sub


