# CollectionExt.Range Method

Generates a sequence of integral numbers within a specified range.

```vb
Public Function Range(ByVal Start As Long, ByVal Count As Long) As Collection
```

### Parameters

**Start** `Long` <br>
The value of the first integer in the sequence.

**Count** `Long` <br>
The number of sequential integers to generate.

### Returns

`Collection` <br>
A collection that contains a range of sequential integral numbers.

### Errors

`OnArgumentOutOfRange` <br>
`Count` is less than 0.

-or-

`Start` + `Count` -1 is larger than LongMaxValue

## Examples

The following code example demonstrates how to use Range to generate a sequence of values.

```vb
' Main module
Option Explicit

Public Sub Start()

    Dim Squares As Collection
    Set Squares = MakeSquares(CollectionExt.Range(1, 10))
    
    Dim Num As Variant
    For Each Num In Squares
        Debug.Print Num
    Next Num
    
End Sub


Private Function MakeSquares(ByVal Numbers As Collection) As Collection
    
    Dim Output As New Collection
    Dim Num As Variant
    For Each Num In Numbers
        Output.Add Num * Num
    Next Num
    Set MakeSquares = Output
    
End Function

' The example displays the following output:
' 1
' 4
' 9
' 16
' 25
' 36
' 49
' 64
' 81
' 100
```

