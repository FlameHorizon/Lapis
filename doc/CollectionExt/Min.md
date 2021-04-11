# CollectionExt.Min Method

Invokes a transform function on each element of a generic sequence and returns the minimum resulting value.

```vb
Public Function Min(ByVal Source As Collection, Optional ByVal Selector As ICallable) As Variant
```

### Parameters

**Source** `Collection` <br>
A sequence of values to determine the minimum value of

**Selector** `ICallable` <br>
Optional. A transform function to apply to each element.

### Returns

`Variant` <br>
The minimum value in the sequence.

### Errors

`OnArgumentNull` <br>
`Source` is `Nothing`.

`OnInvalidOperation` <br>
Default comparer wasn't found for `Value` argument.

## Examples

The following code example demonstrates how to use `Min` to determine the minimum value in a sequence.

```vb
' Standard Module: Main
Option Explicit

Public Sub Start()
    
    Dim Doubles As Collection
    Set Doubles = CollectionExt.Make(1.5E+104, 9E+103, -2E+103)

    Dim Min As Double
    Min = CollectionExt.Min(Doubles)
    
    Debug.Print "The smallest number is " & Min
    
End Sub

' The example displays the following output:
' The youngest animal is age -2E+103

```

The following code example demonstrates how to use `Min` to determine the minimum value in a sequence of projected values.

```vb
' Pet class module

Option Explicit

Implements IComparable

Public Name As String
Public Age As Long


Public Function CompareTo(ByVal x As Pet) As Long

    Compare = IIf(Age < y.Age, -1, _
              IIf(Age = y.Age, 0, 1))
              
End Function


Private Function IComparable_CompareTo(ByRef x As Variant) As Long
    IComparable_CompareTo = Me.CompareTo(x)
End Function
```

```vb
' Main module
Option Explicit

Public Sub Start()

    Dim Pets As Collection
    Set Pets = CollectionExt.Make(MakePet("Barley", 8), _
                                  MakePet("Boots", 4), _
                                  MakePet("Whiskers", 1))
    
    Dim Min As Pet
    Set Min = CollectionExt.Min(Pets, Lambda.Create("$1.Age")
    
    Debug.Print "The youngest animal is age " & Min.Age
    
End Sub


Private Function MakePet(ByVal Name As String, ByVal Age As Long) As Pet
    
    Dim Output As New Pet
    Output.Name = Name
    Output.Age = Age
    Set MakePet = Output
    
End Function

' The example displays the following output:
' The youngest animal is age 1
```

