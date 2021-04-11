VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LpCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Collection")
Option Explicit

Private Const ModuleName As String = "LpCollection"

Private pInit As Boolean
Private pItems As Collection


Public Sub Init(ByVal Items As Collection)

    Const MethodName = "Init"

    If pInit Then
        Errors.OnInvalidOperation vbNullString, _
                                  "This object is already initialized. " _
                                  & ModuleName & "." & MethodName
    End If
    
    If Items Is Nothing Then
        Lapis.Errors.OnArgumentNull "Items", ModuleName & "." & MethodName
    End If
    
    Set pItems = New Collection
    Dim Item As Variant
    For Each Item In Items
        pItems.Add Item
    Next Item
    
    pInit = True

End Sub


Private Sub CheckInit()

    If pInit = False Then
        Errors.OnInvalidOperation vbNullString, _
                                  "This object is not initialized." _
                                  & ModuleName & ".CheckInit"
    End If

End Sub


Public Sub Init2(ParamArray Items() As Variant)

    Const MethodName = "Init2"

    If pInit Then
        Errors.OnInvalidOperation vbNullString, _
                                  "This object is already initialized. " _
                                  & ModuleName & "." & MethodName
    End If
    
    Set pItems = New Collection
    
    Dim i As Long
    Dim Lower As Long: Lower = LBound(Items)
    Dim Upper As Long: Upper = UBound(Items)

    For i = Lower To Upper
        pItems.Add Items(i)
    Next i

    pInit = True

End Sub


Public Function Add(ByVal Item As Variant) As LpCollection
    
    CheckInit
    pItems.Add Item
    Set Add = Me

End Function


Public Function Item(ByVal Index As Long) As Variant
    
    CheckInit
    System.CopyVariant Item, pItems.Item(Index)
    
End Function


Public Function Remove(ByVal Index As Long) As LpCollection
    
    CheckInit
    pItems.Remove Index
    Set Remove = Me
    
End Function


Public Function GroupBy(ByVal PropertyName As String) As LpCollection

    CheckInit
    Set pItems = CollectionExt.GroupBy(pItems, PropertyName)
    Set GroupBy = Me

End Function


Public Function Concat(ByVal Second As Collection) As LpCollection

    CheckInit
    Set pItems = CollectionExt.Concat(pItems, Second)
    Set Concat = Me

End Function


Public Function ToArray() As Variant()
    
    CheckInit
    ToArray = CollectionExt.ToArray(pItems)
    
End Function


Public Function Distinct(Optional ByVal Comparer As IEqualityComparer) As LpCollection
    
    CheckInit
    Set pItems = CollectionExt.Distinct(pItems, Comparer)
    Set Distinct = Me
    
End Function


Public Function Contains(ByVal Value As Variant, Optional ByVal Comparer As IEqualityComparer) As Boolean
    
    CheckInit
    Contains = CollectionExt.Contains(pItems, Value, Comparer)
    
End Function


Public Function IndexOf(ByVal Value As Variant, Optional ByVal Comparer As IEqualityComparer) As Long
    
    CheckInit
    IndexOf = CollectionExt.IndexOf(pItems, Value, Comparer)
    
End Function


Public Function Sort(ByVal Comparer As IComparer) As LpCollection
    
    CheckInit
    Set pItems = CollectionExt.Sort(pItems, Comparer)
    Set Sort = Me
    
End Function


Public Function AddParamsArray(ParamArray Items() As Variant) As LpCollection
    
    CheckInit
    Dim i As Long
    Dim Lower As Long: Lower = LBound(Items)
    Dim Upper As Long: Upper = UBound(Items)

    For i = Lower To Upper
        pItems.Add Items(i)
    Next i
    Set AddParamsArray = Me
    
End Function


Public Function Except(ByVal Second As Collection, Optional ByVal Comparer As IEqualityComparer) As LpCollection
    
    CheckInit
    Set pItems = CollectionExt.Except(pItems, Second, Comparer)
    Set Except = Me
    
End Function


Public Function Intersect(ByVal Second As Collection, Optional ByVal Comparer As IEqualityComparer) As LpCollection
    
    CheckInit
    Set pItems = CollectionExt.Intersect(pItems, Second, Comparer)
    Set Intersect = Me
    
End Function


Public Function Min(Optional ByVal Selector As ICallable) As LpCollection
    
    CheckInit
    Set pItems = CollectionExt.Min(pItems, Selector)
    Set Min = Me

End Function


Public Function Max(Optional ByVal Selector As ICallable) As LpCollection
    
    CheckInit
    Set pItems = CollectionExt.Max(pItems, Selector)
    Set Max = Me

End Function


Public Function Reverse() As LpCollection
    
    CheckInit
    Set pItems = CollectionExt.Reverse(pItems)
    Set Reverse = Me
    
End Function


Public Function Take(ByVal Count As Long) As LpCollection
    
    CheckInit
    Set pItems = CollectionExt.Take(pItems, Count)
    Set Take = Me

End Function


Public Function Skip(ByVal Count As Long) As LpCollection
    
    CheckInit
    Set pItems = CollectionExt.Skip(pItems, Count)
    Set Skip = Me
    
End Function


Public Function SequenceEqual(ByVal Second As LpCollection, _
                              ByVal Comparer As IEqualityComparer) As Boolean
    
    CheckInit
    SequenceEqual = CollectionExt.SequenceEqual(pItems, Second, Comparer)
    
End Function


Public Function All(ByVal Predicate As ICallable) As Boolean
    
    CheckInit
    All = CollectionExt.All(pItems, Predicate)
    
End Function


Public Function Where(ByVal Predicate As ICallable) As LpCollection
    
    CheckInit
    Set pItems = CollectionExt.Where(pItems, Predicate)
    Set Where = Me
    
End Function


Public Function Some(Optional ByVal Predicate As ICallable) As Boolean
    
    CheckInit
    Some = CollectionExt.Some(pItems, Predicate)
    
End Function
                              

Public Function Sum(Optional ByVal Predicate As ICallable) As Variant
    
    CheckInit
    Sum = CollectionExt.Sum(pItems, Predicate)
    
End Function


Public Function Average(Optional ByVal Predicate As ICallable) As Variant
    
    CheckInit
    Average = CollectionExt.Average(pItems, Predicate)
    
End Function


Public Function Count(Optional ByVal Predicate As ICallable) As Variant
    
    CheckInit
    Count = CollectionExt.Count(pItems, Predicate)
    
End Function


Public Function Convert(ByVal Selector As ICallable) As LpCollection
    
    CheckInit
    Set pItems = CollectionExt.Convert(pItems, Selector)
    Set Convert = Me
    
End Function


Public Function First(Optional ByVal Predicate As ICallable) As Variant
    
    CheckInit
    System.CopyVariant First, CollectionExt.First(pItems, Predicate)
    
End Function


Public Function Last(Optional ByVal Predicate As ICallable) As Variant
    
    CheckInit
    System.CopyVariant Last, CollectionExt.Last(pItems, Predicate)
    
End Function


Public Function SelectOne(Optional ByVal Predicate As ICallable) As Variant
    
    CheckInit
    System.CopyVariant SelectOne, CollectionExt.SelectOne(pItems, Predicate)
    
End Function
