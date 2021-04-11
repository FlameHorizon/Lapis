# CollectionExt.Max Method

Invokes a transform function on each element of a sequence and returns the maximum value.

```vb
Public Function Max(ByVal Source As Collection, Optional ByVal Selector As ICallable) As Variant
```

### Parameters

**Source** `Collection` <br>
A sequence of values to determine the maximum value of `Source`.

**Selector** `ICallable` <br>
Optional. A Comparer to apply to each element.

### Returns

`Variant` <br>
The maximum value in the sequence.

### Errors

`OnArgumentNull` <br>
`Source` is `Nothing`.

`OnInvalidOperation` <br>
Default comparer wasn't found for `Value` argument.

## Examples

The following code example demonstrates how to use `Max` to determine the maximum value in a sequence

```vb
' Standard Module: Main
Option Explicit

Public Sub Start()

    Dim Longs As Collection
    Set Longs = CollectionExt.Make(4294967296#, 466855135#, 81125#)
    
    Dim Max As LongLong
    Max = CollectionExt.Max(Longs)
    Debug.Print "The largest number is " & Max & "."

End Sub

' The example displays the following output:
' The largest number is 4294967296.
```


The following code example demonstrates how to use `Max` to determine the maximum value in a sequence of projected values.

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
    Set Stones = CollectionExt.Make(MakeStone(3, 10), _
                                    MakeStone(8, 20), _
                                    MakeStone(6, 14))

    Dim Max As TestStone
    Set Max = CollectionExt.Max(Stones, Lambda.Create("$1.Age"))

    Debug.Print "The oldest stone is age " & Max.Age

End Sub

' The example displays the following output:
' The oldest stone is age 8


Public Function MakeStone(ByVal Age As Single, ByVal Weight As Single) As TestStone

    Dim Output As New TestStone
    Output.Age = Age
    Output.Weight = Weight
    Set MakeStone = Output

End Function
```

