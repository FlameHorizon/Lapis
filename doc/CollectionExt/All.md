# CollectionExt.All Method

Determines whether all elements of a sequence satisfy a condition.

```vb
Public Function All(ByVal Source As Collection, ByVal Predicate As ICallable) As Boolean
```

### Parameters

**Source** `Collection` <br>
A Collection that contains the elements to apply the predicate to.

**Predicate** `ICallable` <br>
A function to test each element for a condition.

### Returns

`Boolean` <br>
`True` if every element of the source sequence passes the test in the specified predicate, or if the sequence is empty; otherwise, `False`.

### Errors

`OnArgumentNull` <br>
`Source` or `Predicate` is `Nothing`

## Examples

The following code example demonstrates how to use All to determine whether all the elements in a sequence satisfy a condition. Variable `AllOlderThan15Years` is true if all the stones are older than 15.

```vb
' Standard Module: Main
Option Explicit

Public Sub Start()

    Dim Stones As Collection
    Set Stones = MakeTestStones

    Dim AllOlderThan15Years As Boolean
    AllOlderThan15Years = CollectionExt.All(Stones, Lambda.Create("$1.Age >= 15"))
    
    If AllOlderThan15Years Then
        Debug.Print "All Stones are older than 15 years."
    Else
        Debug.Print "Not all Stones are older than 15 years."
    End If

End Sub

' This code produces the following output:
' Not all Stones are older than 15 years.


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

