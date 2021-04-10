Attribute VB_Name = "CollectionExt2"
'@Folder("Helper")
Option Explicit

Private Const ModuleName  As String = "CollectionExt2"


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
    Sum = CollectionExt2.Sum(Source, Selector)
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
    
    Count = CollectionExt2.Where(Source, Predicate).Count

End Function
