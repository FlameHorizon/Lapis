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
    Else
        For Each Item In Source
            If System.IsNothing(Item) Then
                GoTo NextItem
            End If
            
            Output = Output + Selector.Run(Item)
            
NextItem:
        Next Item
    End If
    
    Sum = Output

End Function
