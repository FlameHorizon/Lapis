# CollectionExt.Some Method

Determines whether any element of a sequence satisfies a condition. If `Predicate` is not given, then determines whether a sequence contains any elements.

```vb
Public Function Some(ByVal Source As Collection, Optional ByVal Predicate As ICallable) As Boolean
```

### Parameters

**Source** `Collection` <br>
A Collection whose elements to apply the predicate to.

**Predicate** `ICallable` <br>
Optional. A function to test each element for a condition.

### Returns

`Boolean` <br>
`True` if any elements in the source sequence pass the test in the specified predicate; otherwise, `False`.

### Errors

`OnArgumentNull` <br>
`Source` is `Nothing`

## Examples

The following code example demonstrates how to use `Some` to determine whether a sequence contains any elements.

```vb
' Class Module: TestStone
Option Explicit

Public Weight As Single
Public Age As Single
```

```vb
' Standard Module: Main
Option Explicit

Public Sub Start()

    Dim Stones As Collection
    Set Stones = MakeTestStones
    
    If CollectionExt.Any(Stones) Then
        Debug.Print "Set is not empty"
    Else
        Debug.Print "Set is empty."
    End If

End Sub

' This code produces the following output:
' Set is not empty.


Public Function MakeTestStones() As Collection
    
    Dim Output As New Collection
    Output.Add MakeStone(12, 23)
    Output.Add MakeStone(43, 27.5)
    Output.Add MakeStone(23, 11)
    Set MakeTestStones = Output
    
End Function


Public Function MakeStone(ByVal Age As Single, ByVal Weight As Single) As TestStone

    Dim Output As New TestStone
    Output.Age = Age
    Output.Weight = Weight
    Set MakeStone = Output

End Function
```

The following code example demonstrates how to use `Some` to determine whether any element in a sequence satisfies a condition.

```vb
' Class Module: TestStone
Option Explicit

Public Weight As Single
Public Age As Single
```

```vb
' Standard Module: Main
Option Explicit

Public Sub Start()

    Dim Stones As Collection
    Set Stones = MakeTestStones

    Dim SomeMatches As Boolean
    SomeMatches = CollectionExt.Some(Stones, Lambda.Create("$1.Age <= 10 and $1.Weight = 20"))
    
    If SomeMatches Then
        Debug.Print "There are some big young stones"
    Else
        Debug.Print "There are no big young stones"
    End If
    
End Sub

' This code produces the following output:
' There are no big young stones

Public Function MakeTestStones() As Collection
    
    Dim Output As New Collection
    Output.Add MakeStone(12, 23)
    Output.Add MakeStone(43, 27.5)
    Output.Add MakeStone(23, 11)
    Set MakeTestStones = Output
    
End Function


Public Function MakeStone(ByVal Age As Single, ByVal Weight As Single) As TestStone

    Dim Output As New TestStone
    Output.Age = Age
    Output.Weight = Weight
    Set MakeStone = Output

End Function
```


