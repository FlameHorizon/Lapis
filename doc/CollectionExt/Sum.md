# CollectionExt.Sum Method

Computes the sum of a sequence of values. If `Selector` is not defined computes the sum of a sequence values.

```vb
Public Function Sum(ByVal Source As Collection, Optional ByVal Selector As ICallable) As Variant
```

### Parameters

**Source** `Collection` <br>
A sequence of values to calculate the sum of.

**Selector** `ICallable` <br>
Optional. Method how elements of set will be converted into summable values.

### Returns

`Variant` <br>
The sum of the values in the sequence.

### Errors

`OnArgumentNull` <br>
`Source` is `Nothing`

## Examples

The following code example demonstrates how to use Sum to sum the values of a sequence.

```vb
' Standard Module: Main
Option Explicit

Public Sub Start()

    Dim Numbers As Collection
    Set Numbers = CollectionExt.Make(2, 3, 5, 10)
    
    Dim Sum As Long
    Sum = CollectionExt.Sum(Numbers)
    
    Debug.Print "The sum fo the numbers is " & Sum
    
End Sub

' This code produces the following output:
' The sum fo the numbers is 20
```

Computes the sum of the sequence of values that are obtained by invoking a transform function on each element of the input sequence.

```vb
' Standard Module: Main
Option Explicit

Public Function MakeStone(ByVal Age As Single, ByVal Weight As Single) As TestStone

    Dim Output As New TestStone
    Output.Age = Age
    Output.Weight = Weight
    Set MakeStone = Output

End Function


Public Function MakeTestStones() As Collection
    
    Dim Output As New Collection
    Output.Add MakeStone(12, 23)
    Output.Add MakeStone(43, 27.5)
    Output.Add MakeStone(23, 11)
    Set MakeTestStones = Output
    
End Function


Public Sub Start()

    Dim Stones As Collection
    Set Stones = MakeTestStones
    
    Dim Sum As Single
    Sum = CollectionExt.Sum(Stones, Lambda.Create("$1.Weight"))
    
    Debug.Print "The total weight of the stones is: " & Sum
    
End Sub

' This code produces the following output:
' The total weight of the stones is: 61,5
```

